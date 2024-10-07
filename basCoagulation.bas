Attribute VB_Name = "basCoagulation"
Option Explicit

Public Function AreCoagResultsPresent(ByVal SampleID As String) As Long

          Dim tb As New Recordset
          Dim sql As String


10        On Error GoTo AreCoagResultsPresent_Error

20        AreCoagResultsPresent = 0

30        If SampleID = "" Then Exit Function

40        sql = "SELECT count(*) as tot from CoagResults WHERE " & _
                "SampleID = '" & SampleID & "'"
50        Set tb = New Recordset
60        Set tb = Cnxn(0).Execute(sql)

70        AreCoagResultsPresent = Sgn(tb!Tot)

80        Exit Function

AreCoagResultsPresent_Error:

          Dim strES As String
          Dim intEL As Integer



90        intEL = Erl
100       strES = Err.Description
110       LogError "basCoagulation", "AreCoagResultsPresent", intEL, strES, sql

End Function

Public Function CoagCodeFor(ByVal TestName As String) _
       As String

          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo CoagCodeFor_Error

20        CoagCodeFor = "???"

30        sql = "SELECT * from Coagtestdefinitions WHERE testname = '" & Trim(TestName) & "' and inuse = 1"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql

60        If Not tb.EOF Then
70            CoagCodeFor = tb!Code
80        End If

90        Exit Function

CoagCodeFor_Error:

          Dim strES As String
          Dim intEL As Integer



100       intEL = Erl
110       strES = Err.Description
120       LogError "basCoagulation", "CoagCodeFor", intEL, strES, sql

End Function

Public Function CoagNameFor(ByVal Code As String) _
       As String

          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo CoagNameFor_Error

20        CoagNameFor = Code

30        sql = "SELECT * from Coagtestdefinitions WHERE Code = '" & Trim(Code) & "'"

40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql

60        If Not tb.EOF Then
70            CoagNameFor = Trim(tb!TestName)
80        End If

90        Exit Function

CoagNameFor_Error:

          Dim strES As String
          Dim intEL As Integer



100       intEL = Erl
110       strES = Err.Description
120       LogError "basCoagulation", "CoagNameFor", intEL, strES, sql

End Function

Public Function CoagPrintFormat(ByVal Code As String) _
       As Long

          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo CoagPrintFormat_Error

20        CoagPrintFormat = 1

30        sql = "SELECT * from Coagtestdefinitions WHERE (Code = '" & Trim(Code) & "' OR TestName = '" & Code & "')"

40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql

60        If Not tb.EOF Then
70            CoagPrintFormat = tb!DP
80        End If

90        Exit Function

CoagPrintFormat_Error:

          Dim strES As String
          Dim intEL As Integer



100       intEL = Erl
110       strES = Err.Description
120       LogError "basCoagulation", "CoagPrintFormat", intEL, strES, sql

End Function

Public Function CoagUnitsFor(ByVal Code As String) _
       As String

          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo CoagUnitsFor_Error

20        CoagUnitsFor = "???"

30        sql = "SELECT * from Coagtestdefinitions WHERE Code = '" & Trim(Code) & "' and inuse = 1"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql

60        If Not tb.EOF Then
70            CoagUnitsFor = Trim(tb!Units & "")
80        End If

90        Exit Function

CoagUnitsFor_Error:

          Dim strES As String
          Dim intEL As Integer



100       intEL = Erl
110       strES = Err.Description
120       LogError "basCoagulation", "CoagUnitsFor", intEL, strES, sql

End Function
Public Function ACoagUnitsFor(ByVal Code As String) _
       As String

          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo ACoagUnitsFor_Error

20        ACoagUnitsFor = "???"

30        sql = "SELECT * FROM CoagTestDefinitions WHERE " & _
                "TestName = '" & Trim(Code) & "'"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql

60        If Not tb.EOF Then
70            ACoagUnitsFor = tb!Units & ""
80        End If

90        If Trim(ACoagUnitsFor) = "ÆG/ML" Then ACoagUnitsFor = "ug/mL"

100       Exit Function

ACoagUnitsFor_Error:

          Dim strES As String
          Dim intEL As Integer



110       intEL = Erl
120       strES = Err.Description
130       LogError "basCoagulation", "ACoagUnitsFor", intEL, strES

End Function
Public Function InterC(ByVal Value As Single, ByVal LowLim As Single, ByVal HighLim As Single) As String

10        On Error GoTo InterC_Error

20        InterC = ""

30        If Not IsNull(LowLim) Then
40            If Value < LowLim Then InterC = "CL"
50        End If

60        If Not IsNull(HighLim) Then
70            If HighLim <> 0 And Value > HighLim Then InterC = "CH"
80        End If

90        Exit Function

InterC_Error:

          Dim strES As String
          Dim intEL As Integer



100       intEL = Erl
110       strES = Err.Description
120       LogError "basCoagulation", "InterC", intEL, strES

End Function

Public Function InterpCoag(ByVal sex As String, _
                           ByVal TestCode As String, _
                           ByVal Result As String, _
                           ByVal strDaysOld As String, Optional GreaterSignflag As Boolean = False) _
                           As String

          Dim tb As New Recordset
          Dim Low As String
          Dim High As String
          Dim sql As String

10        On Error GoTo InterpCoag_Error

20        InterpCoag = ""

          'Zyam 14-3-24
          If Not GreaterSignflag Then
30              If Val(Result) = 0 Then Exit Function
          End If
          'Zyam


          'sql = "SELECT * from coagtestdefinitions WHERE code = '" & TestCode & "' " & _
           "and agefromdays = '0' and agetodays > '43819' " & _
           "and hospital = '" & HospName(0) & "'"

40        sql = "SELECT * from coagtestdefinitions WHERE (code = '" & TestCode & "' OR TestName = '" & TestCode & "') " & _
                "and agefromdays <= " & strDaysOld & " and agetodays >= " & strDaysOld & " " & _
                ""

50        Set tb = New Recordset
60        RecOpenServer 0, tb, sql

70        If tb.EOF Then
80            Exit Function
90        End If

100       Select Case UCase(Left(sex, 1))
          Case "M":
110           Low = tb!MaleLow
120           High = tb!MaleHigh
130       Case "F":
140           Low = tb!FemaleLow
150           High = tb!FemaleHigh
160       Case Else:
170           Low = tb!FemaleLow
180           High = tb!MaleHigh
190       End Select

200       If Val(Result) > tb!PlausibleHigh Then
210           InterpCoag = "X"
220           Exit Function
230       ElseIf Val(Result) < tb!PlausibleLow Then
240           InterpCoag = "X"
250           Exit Function
260       End If
          'Zyam 14-3-24
          If GreaterSignflag Then
                Result = Right(Result, Len(Result) - 1)
                Result = Val(Result)
           End If
270         If Val(Result) > Val(High) Then
280             InterpCoag = "H"
290         ElseIf Val(Result) < Val(Low) Then
300             InterpCoag = "L"
310         End If
          'Zyam 14-3-24

320       Exit Function

InterpCoag_Error:

          Dim strES As String
          Dim intEL As Integer



330       intEL = Erl
340       strES = Err.Description
350       LogError "basCoagulation", "InterpCoag", intEL, strES, sql

End Function





Public Function nrCoag(ByVal TestCode As String, _
                       ByVal sex As String, _
                       ByVal Dob As String) As String

          Dim l As String * 4
          Dim h As String * 4
          Dim fMat As String
          Dim DaysOld As Long
          Dim TestedLow As Long
          Dim TestedHigh As Long
          Dim PF As Long
          Dim sql As String
          Dim tb As New Recordset

10        On Error GoTo nrCoag_Error

20        nrCoag = "(    -    )"

30        sql = "SELECT * from coagtestdefinitions WHERE code = '" & Trim(TestCode) & "' " & _
                "and agefromdays = '0' and agetodays = " & MaxAgeToDays & "'"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        If tb.EOF Then Exit Function
70        PF = tb!DP
80        Select Case PF
          Case 0: fMat = "0"
90        Case 1: fMat = "0.0"
100       Case 2: fMat = "0.00"
110       Case 3: fMat = "0.000"
120       End Select

130       If IsDate(Dob) Then
140           DaysOld = Abs(DateDiff("d", Now, Dob))

150           TestedLow = MaxAgeToDays
160           TestedHigh = 0

170           sql = "SELECT * from coagtestdefinitions WHERE code = '" & TestCode & "' " & _
                    "and agefromdays = '0' and agetodays = " & MaxAgeToDays & "'"
180           Set tb = New Recordset
190           RecOpenServer 0, tb, sql

200           Do While Not tb.EOF
210               If tb!Code = TestCode Then
220                   If tb!AgeFromDays <= TestedLow And tb!AgeToDays >= TestedHigh Then
230                       If DaysOld >= tb!AgeFromDays And DaysOld <= tb!AgeToDays Then
240                           TestedLow = tb!AgeFromDays
250                           TestedHigh = tb!AgeFromDays
260                       End If
270                   End If
280               End If
290               tb.MoveNext
300           Loop
310       End If
320       sql = "SELECT * from coagtestdefinitions WHERE code = '" & TestCode & "' " & _
                "and agefromdays = '0' and agetodays = " & MaxAgeToDays & "'"
330       Set tb = New Recordset
340       RecOpenServer 0, tb, sql

350       If Not tb.EOF Then
360           Select Case sex
              Case "M":
370               If tb!MaleHigh = 999 Then
380                   nrCoag = "           "
390                   Exit Function
400               End If
410               RSet l = Format$(tb!MaleLow, fMat)
420               Mid$(nrCoag, 2, 4) = l
430               LSet h = Format$(tb!MaleHigh, fMat)
440               Mid$(nrCoag, 7, 4) = h
450           Case "F":
460               If tb!FemaleHigh = 999 Then
470                   nrCoag = "           "
480                   Exit Function
490               End If
500               RSet l = Format$(tb!FemaleLow, fMat)
510               Mid$(nrCoag, 2, 4) = l
520               LSet h = Format$(tb!FemaleHigh, fMat)
530               Mid$(nrCoag, 7, 4) = h
540           Case Else:
550               If tb!MaleHigh = 999 Then
560                   nrCoag = "           "
570                   Exit Function
580               End If
590               RSet l = Format$(tb!FemaleLow, fMat)
600               Mid$(nrCoag, 2, 4) = l
610               LSet h = Format$(tb!MaleHigh, fMat)
620               Mid$(nrCoag, 7, 4) = h
630           End Select
640       Else
650           nrCoag = "           "
660       End If

670       Exit Function

nrCoag_Error:

          Dim strES As String
          Dim intEL As Integer



680       intEL = Erl
690       strES = Err.Description
700       LogError "basCoagulation", "nrCoag", intEL, strES

End Function

Public Function UnitConv(ByVal Unit As String)

10        On Error GoTo UnitConv_Error

20        If Unit = "" Then Exit Function

30        UnitConv = Trim(Unit)

40        If UCase(Trim(Unit)) = "ÆG/ML" Then UnitConv = "ug/ml"

50        If Trim(Unit) = "INR" Then UnitConv = ""

60        Exit Function

UnitConv_Error:

          Dim strES As String
          Dim intEL As Integer



70        intEL = Erl
80        strES = Err.Description
90        LogError "basCoagulation", "UnitConv", intEL, strES

End Function


