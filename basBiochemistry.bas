Attribute VB_Name = "basBiochemistry"
Option Explicit
Public ControlName() As String
Public frmMainCounter As Long
Public frmMainImageCounter As Long
Public pbCounter As Long
Public Function Check_Bio(ByVal ShortName As String, _
                          ByVal Units As String, _
                          ByVal SampleType As String) As String

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo Check_Bio_Error

20        Check_Bio = ""

30        sql = "SELECT Units, SampleType FROM BioTestDefinitions WHERE " & _
                "ShortName = '" & ShortName & "'"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        If tb.EOF Then
70            Check_Bio = "Test Name"
80        Else
90            If ListCodeFor("ST", SampleType) <> tb!SampleType Then
100               Check_Bio = "Sample Type"
110           ElseIf Units <> tb!Units Then
120               Check_Bio = "Units"
130           End If
140       End If

150       Exit Function

Check_Bio_Error:

          Dim strES As String
          Dim intEL As Integer

160       intEL = Erl
170       strES = Err.Description
180       LogError "basBiochemistry", "Check_Bio", intEL, strES, sql

End Function
Public Function AreBioResultsPresent(ByVal SampleID As String) As Long

          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo AreBioResultsPresent_Error

20        AreBioResultsPresent = 0

30        If SampleID = "" Then Exit Function

40        sql = "SELECT count(*) as tot from BioResults WHERE " & _
                "SampleID = '" & SampleID & "'"
50        Set tb = New Recordset
60        Set tb = Cnxn(0).Execute(sql)

70        AreBioResultsPresent = Sgn(tb!Tot)

80        Exit Function

AreBioResultsPresent_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "basBiochemistry", "AreBioResultsPresent", intEL, strES, sql

End Function

Public Function BioLongNameFor(ByVal LongOrShortName As String) As String

          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo BioLongNameFor_Error

20        sql = "SELECT LongName FROM BioTestDefinitions WHERE " & _
                "ShortName = '" & AddTicks(LongOrShortName) & "' " & _
                "OR LongName = '" & AddTicks(LongOrShortName) & "'"

30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If Not tb.EOF Then
60            BioLongNameFor = tb!LongName & ""
70        Else
80            BioLongNameFor = "???"
90        End If

100       Exit Function

BioLongNameFor_Error:

          Dim strES As String
          Dim intEL As Integer

110       intEL = Erl
120       strES = Err.Description
130       LogError "basBiochemistry", "BioLongNameFor", intEL, strES, sql

End Function

Public Function BioShortNameFor(ByVal LongOrShortName As String) As String

          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo BioShortNameFor_Error

20        sql = "SELECT ShortName FROM BioTestDefinitions WHERE " & _
                "ShortName = '" & AddTicks(LongOrShortName) & "' " & _
                "OR LongName = '" & AddTicks(LongOrShortName) & "'"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If Not tb.EOF Then
60            BioShortNameFor = tb!ShortName & ""
70        Else
80            BioShortNameFor = "???"
90        End If

100       Exit Function

BioShortNameFor_Error:

          Dim strES As String
          Dim intEL As Integer

110       intEL = Erl
120       strES = Err.Description
130       LogError "basBiochemistry", "BioShortNameFor", intEL, strES

End Function

Public Function CodeForLongName(ByVal LongName As String) _
       As String

          Dim sql As String
          Dim tb As New Recordset

10        On Error GoTo CodeForLongName_Error

20        CodeForLongName = "???"

30        sql = "SELECT Code from BioTestDefinitions WHERE longname = '" & LongName & "'"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        If Not tb.EOF Then
70            CodeForLongName = Trim(tb!Code)
80        End If

90        Exit Function

CodeForLongName_Error:

          Dim strES As String
          Dim intEL As Integer

100       intEL = Erl
110       strES = Err.Description
120       LogError "basBiochemistry", "CodeForLongName", intEL, strES

End Function

Public Function CodeForShortName(ByVal ShortName As String) _
       As String

          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo CodeForShortName_Error

20        CodeForShortName = "???"

30        sql = "SELECT Code from BioTestDefinitions WHERE " & _
                "ShortName = '" & ShortName & "'"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        If Not tb.EOF Then
70            CodeForShortName = Trim(tb!Code)
80        End If

90        Exit Function

CodeForShortName_Error:

          Dim strES As String
          Dim intEL As Integer

100       intEL = Erl
110       strES = Err.Description
120       LogError "basBiochemistry", "CodeForShortName", intEL, strES, sql

End Function

Public Sub GetControlNames()

          Dim sn As New Recordset
          Dim sql As String
          Dim n As Long

10        On Error GoTo GetControlNames_Error

20        ReDim ControlName(0 To 0)
30        sql = "SELECT distinct controlname from controls"
40        Set sn = New Recordset
50        RecOpenServer 0, sn, sql
60        n = 0
70        Do While Not sn.EOF
80            ReDim Preserve ControlName(0 To n)
90            ControlName(n) = (sn!ControlName & "")
100           sn.MoveNext
110           n = n + 1
120       Loop

130       sn.Close

140       Exit Sub

GetControlNames_Error:

          Dim strES As String
          Dim intEL As Integer

150       intEL = Erl
160       strES = Err.Description
170       LogError "basBiochemistry", "GetControlNames", intEL, strES

End Sub

Sub LogBioAsPrinted(ByVal SampleID As String, _
                    ByVal TestCode As String)

          Dim sql As String

10        On Error GoTo LogBioAsPrinted_Error

20        sql = "UPDATE BioResults " & _
                "set valid = 1, printed = 1 WHERE " & _
                "SampleID = '" & SampleID & "' " & _
                "and code = '" & TestCode & "'"

30        Cnxn(0).Execute sql

40        Exit Sub

LogBioAsPrinted_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "basBiochemistry", "LogBioAsPrinted", intEL, strES, sql

End Sub



Public Function LongNameforCode(ByVal Code As String) _
       As String

          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo LongNameforCode_Error

20        LongNameforCode = "???"

30        sql = "SELECT LongName FROM BioTestDefinitions WHERE " & _
                "Code = '" & Code & "'"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        If Not tb.EOF Then
70            LongNameforCode = Trim(tb!LongName)
80        End If

90        Exit Function

LongNameforCode_Error:

          Dim strES As String
          Dim intEL As Integer

100       intEL = Erl
110       strES = Err.Description
120       LogError "basBiochemistry", "LongNameforCode", intEL, strES

End Function

Public Sub PrintResultBioWin(ByVal SampleID As String)

          Dim tb As New Recordset
          Dim sql As String
          Dim GP As String
          Dim Clin As String
          Dim Ward As String

10        On Error GoTo PrintResultBioWin_Error

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
                "Department = 'B' " & _
                "AND SampleID = '" & SampleID & "'"
140       Set tb = New Recordset
150       RecOpenClient 0, tb, sql
160       If tb.EOF Then
170           tb.AddNew
180       End If
190       tb!SampleID = SampleID
200       tb!Department = "B"
210       tb!Initiator = Username
220       tb!Ward = Ward & ""
230       tb!GP = GP & ""
240       tb!Clinician = Clin & ""
250       tb!pTime = Now
260       tb.Update

270       Exit Sub

PrintResultBioWin_Error:

          Dim strES As String
          Dim intEL As Integer



280       intEL = Erl
290       strES = Err.Description
300       LogError "basBiochemistry", "PrintResultBioWin", intEL, strES

End Sub

Public Sub PrintResultCWin(ByVal SampleID As String)

          Dim tb As New Recordset
          Dim sql As String
          Dim GP As String
          Dim Clin As String
          Dim Ward As String

10        On Error GoTo PrintResultCWin_Error

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
                "Department = 'C' " & _
                "AND SampleID = '" & SampleID & "'"
140       Set tb = New Recordset
150       RecOpenClient 0, tb, sql
160       If tb.EOF Then
170           tb.AddNew
180       End If
190       tb!SampleID = SampleID
200       tb!Department = "C"
210       tb!Initiator = Username
220       tb!pTime = Now
230       tb!Ward = Ward & ""
240       tb!GP = GP & ""
250       tb!Clinician = Clin & ""
260       tb.Update

270       Exit Sub

PrintResultCWin_Error:

          Dim strES As String
          Dim intEL As Integer



280       intEL = Erl
290       strES = Err.Description
300       LogError "basBiochemistry", "PrintResultCWin", intEL, strES

End Sub

Public Function QuickInterpBio(ByVal Result As BIEResult) _
       As String

10        On Error GoTo QuickInterpBio_Error

20        If Not IsNumeric(Result.Result) Then
30            QuickInterpBio = Result.Result
40            Exit Function
50        End If
60        With Result
70            If Val(.Result) > .PlausibleHigh Then
80                QuickInterpBio = "***"
90            ElseIf Val(.Result) < .PlausibleLow Then
100               QuickInterpBio = "***"
110           ElseIf Val(.Result) < .Low Then
120               QuickInterpBio = "Low "
130           ElseIf Val(.Result) > .High Then
140               QuickInterpBio = "High"
150           Else
160               QuickInterpBio = "    "
170           End If
180       End With

190       Exit Function

QuickInterpBio_Error:

          Dim strES As String
          Dim intEL As Integer

200       intEL = Erl
210       strES = Err.Description
220       LogError "basBiochemistry", "QuickInterpBio", intEL, strES

End Function

Public Function QuickInterpBioRes(ByVal Code As String, ByVal Result As String, _
                                  ByVal DaysOld As Integer, _
                                  ByVal sex As String, _
                                  ByVal DefIndex As Integer) _
                                  As String

          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo QuickInterpBioRes_Error

20        sql = "Select * from biotestdefinitions where code = '" & Code & "'"
30        If DefIndex <> 0 Then
40            sql = sql & " and defindex = '" & DefIndex & "'"
50        Else
60            sql = sql & " and inuse = 1"
70        End If

80        Set tb = New Recordset
90        RecOpenServer 0, tb, sql
100       Do While Not tb.EOF
110           If DaysOld >= tb!AgeFromDays And DaysOld <= tb!AgeToDays Then
120               If Left(sex, 1) = "M" And Val(Result) < tb!MaleLow Then
130                   QuickInterpBioRes = "Low "
140               ElseIf Left(sex, 1) = "M" And Val(Result) > tb!MaleHigh Then
150                   QuickInterpBioRes = "High"
160               ElseIf Left(sex, 1) = "F" And Val(Result) < tb!FemaleLow Then
170                   QuickInterpBioRes = "Low "
180               ElseIf Left(sex, 1) = "F" And Val(Result) > tb!FemaleHigh Then
190                   QuickInterpBioRes = "High"
200               ElseIf sex = "" And Val(Result) < tb!FemaleLow Then
210                   QuickInterpBioRes = "Low "
220               ElseIf sex = "" And Val(Result) > tb!MaleLow Then
230                   QuickInterpBioRes = "High"
240               ElseIf Val(Result) > tb!PlausibleHigh Then
250                   QuickInterpBioRes = "*** "
260               ElseIf Val(Result) < tb!PlausibleLow Then
270                   QuickInterpBioRes = "*** "
280               Else
290                   QuickInterpBioRes = "    "
300               End If
310           End If
320           tb.MoveNext
330       Loop

340       Exit Function

QuickInterpBioRes_Error:

          Dim strES As String
          Dim intEL As Integer

350       intEL = Erl
360       strES = Err.Description
370       LogError "basBiochemistry", "QuickInterpBioRes", intEL, strES

End Function
Public Function ShortNameFor(ByVal Code As String) As String

          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo ShortNameFor_Error

20        ShortNameFor = "???"

30        sql = "SELECT ShortName FROM BioTestDefinitions WHERE " & _
                "Code = '" & Code & "'"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        If Not tb.EOF Then
70            ShortNameFor = Trim(tb!ShortName)
80        End If

90        Exit Function

ShortNameFor_Error:

          Dim strES As String
          Dim intEL As Integer

100       intEL = Erl
110       strES = Err.Description
120       LogError "basBiochemistry", "ShortNameFor", intEL, strES

End Function

Function TestAffected(ByVal br As BIEResult) As Boolean

          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo TestAffected_Error

20        TestAffected = False

30        sql = "SELECT COUNT(*) Tot FROM Masks M, BioTestDefinitions D " & _
                "WHERE M.SampleID = '" & br.SampleID & "' " & _
                "AND D.Code = '" & br.Code & "' " & _
                "AND ( (D.H = 1 AND M.H = 1) " & _
                "   OR (D.S = 1 and M.S = 1) " & _
                "   OR (D.L = 1 and M.L = 1) " & _
                "   OR (D.O = 1 and M.O = 1) " & _
                "   OR (D.G = 1 and M.G = 1) " & _
                "   OR (D.J = 1 and M.J = 1) )"

40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql

60        TestAffected = tb!Tot > 0

70        Exit Function

TestAffected_Error:

          Dim strES As String
          Dim intEL As Integer

80        intEL = Erl
90        strES = Err.Description
100       LogError "basBiochemistry", "TestAffected", intEL, strES, sql

End Function

Function TestCode2LongName(ByVal Code As String) As String

          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo TestCode2LongName_Error

20        TestCode2LongName = "???"

30        sql = "SELECT LongName FROM BioTestDefinitions WHERE " & _
                "Code = '" & Code & "'"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        If Not tb.EOF Then
70            TestCode2LongName = Trim(tb!LongName)
80        End If

90        Exit Function

TestCode2LongName_Error:

          Dim strES As String
          Dim intEL As Integer

100       intEL = Erl
110       strES = Err.Description
120       LogError "basBiochemistry", "TestCode2LongName", intEL, strES

End Function

Function TextFor(f() As Long) As String

          Dim T As String
          Dim exclam As Long

10        On Error GoTo TextFor_Error

20        T = ""

30        If f(1) Then T = "Sample slightly Haemolysed: Interference with this test."
40        If f(0) Then T = "Sample Haemolysed: Please Repeat."
50        If f(4) Then T = "Sample grossly haemolysed: Unsuitable for analysis."

60        T = T & Chr(10) & Chr(13)

70        exclam = False
80        If f(2) Then
90            T = T & "!* Sample Lipaemic. "
100           exclam = True
110       End If
120       If f(3) Then
130           If Not exclam Then T = T & "!* "
140           T = T & "Sample old:Unsuitable. "
150           exclam = True
160       End If
170       If f(5) Then
180           If Not exclam Then T = T & "!* "
190           T = T & "Possible Bilirubin interference."
200       End If

210       TextFor = T

220       Exit Function

TextFor_Error:

          Dim strES As String
          Dim intEL As Integer

230       intEL = Erl
240       strES = Err.Description
250       LogError "basBiochemistry", "TextFor", intEL, strES

End Function

Public Function CheckBioFlag(ByVal Code As String, ByVal Res As Double, _
                             ByVal DaysOld As Integer, _
                             ByVal sex As String, _
                             ByVal DefIndex As Integer, ByVal Fasting As String) As String
          Dim sql As String
          Dim tb As Recordset
          Dim s As String

10        On Error GoTo CheckBioFlag_Error

20        sql = "Select * from biotestdefinitions where code = '" & Code & "'"
30        If DefIndex <> 0 Then
40            sql = sql & " and defindex = '" & DefIndex & "'"
50        Else
60            sql = sql & " and inuse = 1"
70        End If

80        Set tb = New Recordset
90        RecOpenServer 0, tb, sql
100       Do While Not tb.EOF
110           If DaysOld >= tb!AgeFromDays And DaysOld <= tb!AgeToDays Then
120               If Left(sex, 1) = "M" Then
130                   If Res > tb!MaleHigh Then
140                       s = "H"
150                   ElseIf Res < tb!MaleLow Then
160                       s = "L"
170                   End If
180               ElseIf Left(sex, 1) = "F" Then
190                   If Res > tb!FemaleHigh Then
200                       s = "H"
210                   ElseIf Res < tb!FemaleLow Then
220                       s = "L"
230                   End If
240               Else
250                   If Res > tb!MaleHigh Then
260                       s = "H"
270                   ElseIf Res < tb!FemaleLow Then
280                       s = "L"
290                   End If
300               End If
310           End If
320           If Res > Val(tb!PlausibleHigh) Then
330               s = "X"
340           ElseIf Res < Val(tb!PlausibleLow) Then
350               s = "X"
360           ElseIf Code = SysOptBioCodeForGlucose(0) Or _
                     Code = SysOptBioCodeForChol(0) Or _
                     Code = SysOptBioCodeForTrig(0) Then
370               If Fasting Then
380                   If Code = SysOptBioCodeForGlucose(0) Or Code = SysOptBioCodeForGlucoseP(0) Then
390                       sql = "SELECT * from fastings WHERE testname = '" & "GLU" & "'"
400                   ElseIf Code = SysOptBioCodeForChol(0) Or Code = SysOptBioCodeForCholP(0) Then
410                       sql = "SELECT * from fastings WHERE testname = '" & "CHO" & "'"
420                   ElseIf Code = SysOptBioCodeForTrig(0) Or Code = SysOptBioCodeForTrigP(0) Then
430                       sql = "SELECT * from fastings WHERE testname = '" & "TRI" & "'"
440                   End If
450                   Set tb = New Recordset
460                   RecOpenServer 0, tb, sql
470                   If Not tb.EOF Then
480                       If Res > tb!FastingHigh Then
490                           s = "H"
500                       ElseIf Res < tb!FastingLow Then
510                           s = "L"
520                       End If
530                   End If
540               End If
550           End If
560           tb.MoveNext
570       Loop

580       CheckBioFlag = s

590       Exit Function

CheckBioFlag_Error:

          Dim strES As String
          Dim intEL As Integer

600       intEL = Erl
610       strES = Err.Description
620       LogError "basBiochemistry", "CheckBioFlag", intEL, strES

End Function


Public Function CheckBioNR(ByVal Code As String, _
                           ByVal DaysOld As Integer, _
                           ByVal sex As String, _
                           ByVal DefIndex As Integer, ByVal Fasting As String) As String
          Dim sql As String
          Dim tb As Recordset
          Dim Nr As String

10        On Error GoTo CheckBioNR_Error

20        sql = "Select * from biotestdefinitions where code = '" & Code & "'"
30        If DefIndex <> 0 Then
40            sql = sql & " and defindex = '" & DefIndex & "'"
50        Else
60            sql = sql & " and inuse = 1"
70        End If

80        Set tb = New Recordset
90        RecOpenServer 0, tb, sql
100       Do While Not tb.EOF
110           If DaysOld >= tb!AgeFromDays And DaysOld <= tb!AgeToDays Then
120               If Left(sex, 1) = "M" Then
130                   Nr = tb!MaleLow & " - " & tb!MaleHigh
140               ElseIf Left(sex, 1) = "F" Then
150                   Nr = tb!FemaleLow & " - " & tb!FemaleHigh
160               Else
170                   Nr = tb!FemaleLow & " - " & tb!MaleHigh
180               End If
190           End If
200           If Code = SysOptBioCodeForGlucose(0) Or _
                 Code = SysOptBioCodeForChol(0) Or _
                 Code = SysOptBioCodeForTrig(0) Then
210               If Fasting Then
220                   If Code = SysOptBioCodeForGlucose(0) Or Code = SysOptBioCodeForGlucoseP(0) Then
230                       sql = "SELECT * from fastings WHERE testname = '" & "GLU" & "'"
240                   ElseIf Code = SysOptBioCodeForChol(0) Or Code = SysOptBioCodeForCholP(0) Then
250                       sql = "SELECT * from fastings WHERE testname = '" & "CHO" & "'"
260                   ElseIf Code = SysOptBioCodeForTrig(0) Or Code = SysOptBioCodeForTrigP(0) Then
270                       sql = "SELECT * from fastings WHERE testname = '" & "TRI" & "'"
280                   End If
290                   Set tb = New Recordset
300                   RecOpenServer 0, tb, sql
310                   If Not tb.EOF Then
320                       Nr = tb!FastingText
330                   End If
340               End If
350           End If
360           tb.MoveNext
370       Loop

380       CheckBioNR = Nr

390       Exit Function

CheckBioNR_Error:

          Dim strES As String
          Dim intEL As Integer

400       intEL = Erl
410       strES = Err.Description
420       LogError "basBiochemistry", "CheckBioNR", intEL, strES

End Function


Public Function BioPrintFormat(ByVal Code As String, ByVal DefIndex As Integer) As Integer

          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo BioPrintFormat_Error

20        sql = "SELECT DP FROM BioTestDefinitions WHERE " & _
                "Code = '" & Code & "'"
30        If DefIndex > 0 Then
40            sql = sql & " AND DefIndex = " & DefIndex & ""
50        End If
60        Set tb = New Recordset
70        RecOpenServer 0, tb, sql
80        If Not tb.EOF Then
90            BioPrintFormat = tb!DP
100       Else
110           BioPrintFormat = 0
120       End If

130       Exit Function

BioPrintFormat_Error:

          Dim strES As String
          Dim intEL As Integer

140       intEL = Erl
150       strES = Err.Description
160       LogError "basBiochemistry", "BioPrintFormat", intEL, strES

End Function


Public Function GetDefIndex(ByVal Code As String, _
                            ByVal DaysOld As Integer, _
                            ByVal sex As String) As Long
          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo GetDefIndex_Error

20        GetDefIndex = 0

30        sql = "Select AgeFromDays, AgeToDays, DefIndex from biotestdefinitions where code = '" & Code & "'  and inuse = 1"

40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        Do While Not tb.EOF
70            If DaysOld >= tb!AgeFromDays And DaysOld <= tb!AgeToDays Then
80                If Left(sex, 1) = "M" Then
90                    GetDefIndex = tb!DefIndex
100               ElseIf Left(sex, 1) = "F" Then
110                   GetDefIndex = tb!DefIndex
120               Else
130                   GetDefIndex = 0
140               End If
150           End If
160           tb.MoveNext
170       Loop

180       Exit Function

GetDefIndex_Error:

          Dim strES As String
          Dim intEL As Integer

190       intEL = Erl
200       strES = Err.Description
210       LogError "basBiochemistry", "GetDefIndex", intEL, strES

End Function

Function TestAffectedNew(ByVal br As BIEResult) As Boolean

          Dim TestName As String
          Dim tb As New Recordset
          Dim sql As String
          Dim sn As New Recordset

10        On Error GoTo TestAffectedNew_Error

20        TestAffectedNew = False
30        TestName = LongNameforCode(br.Code)

40        sql = "SELECT * from masks WHERE sampleid = '" & br.SampleID & "'"
50        Set tb = New Recordset
60        RecOpenServer 0, tb, sql

70        If tb.EOF Then Exit Function

80        sql = "SELECT * FROM BioTestDefinitions WHERE " & _
                "Code = '" & br.Code & "'"

90        Set sn = New Recordset
100       RecOpenServer 0, sn, sql
110       Do While Not sn.EOF
120           With sn
130               If !LongName = TestName And !Code = br.Code Then
140                   If !h And tb!h Then
150                       TestAffectedNew = True
160                       Exit Do
170                   End If
180                   If !s And tb!s Then
190                       TestAffectedNew = True
200                       Exit Do
210                   End If
220                   If !l And tb!l Then
230                       TestAffectedNew = True
240                       Exit Do
250                   End If
260                   If !o And tb!o Then
270                       TestAffectedNew = True
280                       Exit Do
290                   End If
300                   If !g And tb!g Then
310                       TestAffectedNew = True
320                       Exit Do
330                   End If
340                   If !J And tb!J Then
350                       TestAffectedNew = True
360                       Exit Do
370                   End If
380               End If
390           End With
400           sn.MoveNext
410       Loop

420       Exit Function

TestAffectedNew_Error:

          Dim strES As String
          Dim intEL As Integer

430       intEL = Erl
440       strES = Err.Description
450       LogError "basBiochemistry", "TestAffectedNew", intEL, strES

End Function

Public Function BioAnalyserIDForCode(Code As String) As String

      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo BioAnalyserIDForCode_Error

20    sql = "SELECT Analyser FROM BioTestDefinitions WHERE Code = '" & Code & "'"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    If tb.EOF Then
60        BioAnalyserIDForCode = ""
70    Else
80        BioAnalyserIDForCode = tb!Analyser & ""
90    End If

100   Exit Function

BioAnalyserIDForCode_Error:

       Dim strES As String
       Dim intEL As Integer

110    intEL = Erl
120    strES = Err.Description
130    LogError "basBiochemistry", "BioAnalyserIDForCode", intEL, strES, sql
          
End Function
