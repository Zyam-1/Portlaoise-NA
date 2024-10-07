Attribute VB_Name = "basHaematology"
Option Explicit

Public Function AreHaemResultsPresent(ByVal SampleID As String) As Long

          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo AreHaemResultsPresent_Error

20        AreHaemResultsPresent = 0

30        If SampleID = "" Then Exit Function

40        sql = "SELECT count(*) as tot from HaemResults WHERE " & _
                "SampleID = '" & SampleID & "'"
50        Set tb = New Recordset
60        Set tb = Cnxn(0).Execute(sql)

70        AreHaemResultsPresent = Sgn(tb!Tot)

80        Exit Function

AreHaemResultsPresent_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "basHaematology", "AreHaemResultsPresent", intEL, strES, sql

End Function

Sub buildinterp(tb As Recordset, i() As String)

          Dim n As Long
          Dim R As String
          Dim l As Long

10        On Error GoTo buildinterp_Error

20        l = True
30        n = 0

40        If Val(tb!NeutA & "") <> 0 And Val(tb!NeutP & "") <> 0 Then
50            R = Interp(0, tb!NeutA)
60            If R <> "" Then
70                i(0) = R
80                l = False
90            Else
100               R = Interp(1, tb!NeutP)
110               If R <> "" Then
120                   i(0) = R
130                   l = False
140               End If
150           End If
160           R = Interp(2, tb!NeutA)
170           If R <> "" Then
180               i(0) = R
190               l = False
200           Else
210               R = Interp(3, tb!NeutP)
220               If R <> "" Then
230                   i(0) = R
240                   l = False
250               End If
260           End If
270       End If

280       If Val(tb!LymA & "") <> 0 And Val(tb!LymP & "") <> 0 Then
290           R = Interp(4, tb!LymA)
300           If R <> "" Then
310               i(0) = i(0) & R
320               l = Not l
330               If l Then n = n + 1
340           Else
350               R = Interp(5, tb!LymP)
360               If R <> "" Then
370                   i(0) = i(0) & R
380                   l = Not l
390                   If l Then n = n + 1
400               End If
410           End If
420           R = Interp(6, tb!LymA)
430           If R <> "" Then
440               i(0) = i(0) & R
450               l = Not l
460               If l Then n = n + 1
470           Else
480               R = Interp(7, tb!LymP)
490               If R <> "" Then
500                   i(0) = i(0) & R
510                   l = Not l
520                   If l Then n = n + 1
530               End If
540           End If
550       End If

560       If Val(tb!MonoA & "") <> 0 And Val(tb!MonoP & "") <> 0 Then
570           R = Interp(8, tb!MonoA)
580           If R <> "" Then
590               i(n) = i(n) & R
600               l = Not l
610               If l Then n = n + 1
620           Else
630               R = Interp(9, tb!MonoP)
640               If R <> "" Then
650                   i(n) = i(n) & R
660                   l = Not l
670                   If l Then n = n + 1
680               End If
690           End If
700       End If

710       If Val(tb!EosA & "") <> 0 And Val(tb!EosP & "") <> 0 Then
720           R = Interp(10, tb!EosA)
730           If R <> "" Then
740               i(n) = i(n) & R
750               l = Not l
760               If l Then n = n + 1
770           Else
780               R = Interp(11, tb!EosP)
790               If R <> "" Then
800                   i(n) = i(n) & R
810                   l = Not l
820                   If l Then n = n + 1
830               End If
840           End If
850       End If

860       If Val(tb!BasA & "") <> 0 And Val(tb!BasP & "") <> 0 Then
870           R = Interp(12, tb!BasA)
880           If R <> "" Then
890               i(n) = i(n) & R
900               l = Not l
910               If l Then n = n + 1
920           Else
930               R = Interp(13, tb!BasP)
940               If R <> "" Then
950                   i(n) = i(n) & R
960                   l = Not l
970                   If l Then n = n + 1
980               End If
990           End If
1000      End If

1010      If Val(tb!wbc & "") <> 0 Then
1020          R = Interp(14, tb!wbc)
1030          If R <> "" Then
1040              i(n) = i(n) & R
1050              l = Not l
1060              If l Then n = n + 1
1070          Else
1080              R = Interp(15, tb!wbc)
1090              If R <> "" Then
1100                  i(n) = i(n) & R
1110                  l = Not l
1120                  If l Then n = n + 1
1130              End If
1140          End If
1150      End If

1160      If (Val(tb!rdwsd & "") <> 0) And (Val(tb!RDWCV & "") <> 0) Then
1170          R = Interp(16, tb!rdwsd)
1180          If R <> "" Then
1190              i(n) = i(n) & R
1200              l = Not l
1210              If l Then n = n + 1
1220          Else
1230              R = Interp(17, tb!RDWCV)
1240              If R <> "" Then
1250                  i(n) = i(n) & R
1260                  l = Not l
1270                  If l Then n = n + 1
1280              End If
1290          End If
1300      End If

1310      If Val(tb!MCV & "") <> 0 Then
1320          R = Interp(18, tb!MCV)
1330          If R <> "" Then
1340              i(n) = i(n) & R
1350              l = Not l
1360              If l Then n = n + 1
1370          Else
1380              R = Interp(19, tb!MCV)
1390              If R <> "" Then
1400                  i(n) = i(n) & R
1410                  l = Not l
1420                  If l Then n = n + 1
1430              End If
1440          End If
1450      End If

1460      If Val(tb!mchc & "") <> 0 Then
1470          R = Interp(20, tb!mchc)
1480          If R <> "" Then
1490              i(n) = i(n) & R
1500              l = Not l
1510              If l Then n = n + 1
1520          End If
1530      End If

1540      If Val(tb!Hgb & "") <> 0 Then
1550          R = Interp(21, tb!Hgb)
1560          If R <> "" Then
1570              i(n) = i(n) & R
1580              l = Not l
1590              If l Then n = n + 1
1600          End If
1610      End If

1620      If Val(tb!rbc & "") <> 0 Then
1630          R = Interp(22, tb!rbc)
1640          If R <> "" Then
1650              i(n) = i(n) & R
1660              l = Not l
1670              If l Then n = n + 1
1680          End If
1690      End If

1700      If Val(tb!Plt & "") <> 0 Then
1710          R = Interp(23, tb!Plt)
1720          If R <> "" Then
1730              i(n) = i(n) & R
1740              l = Not l
1750              If l Then n = n + 1
1760          Else
1770              R = Interp(24, tb!Plt)
1780              If R <> "" Then
1790                  i(n) = i(n) & R
1800                  l = Not l
1810                  If l Then n = n + 1
1820              End If
1830          End If
1840      End If

1850      Exit Sub

buildinterp_Error:

          Dim strES As String
          Dim intEL As Integer

1860      intEL = Erl
1870      strES = Err.Description
1880      LogError "basHaematology", "buildinterp", intEL, strES

End Sub

Public Sub FillInterpTable()

          Dim n As Long
          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo FillInterpTable_Error

20        sql = "SELECT * from interp"

30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If Not tb.EOF Then
60            For n = 0 To 24
70                InterpList(n) = tb(n)
80            Next
90        End If

          'InterpList(0) = "Neutropaenia   "
          'InterpList(1) = "Neutropaenia   "
          'InterpList(2) = "Neutrophilia   "
          'InterpList(3) = "Neutrophilia   "
          'InterpList(4) = "Lymphopaenia   "
          'InterpList(5) = "Lymphopaenia   "
          'InterpList(6) = "Lymphocytosis  "
          'InterpList(7) = "Lymphocytosis  "
          'InterpList(8) = "Monocytosis    "
          'InterpList(9) = "Monocytosis    "
          'InterpList(10) = "Eosinophilia   "
          'InterpList(11) = "Eosinophilia   "
          'InterpList(12) = "Basophilia     "
          'InterpList(13) = "Basophilia     "
          'InterpList(14) = "Leucocytosis   "
          'InterpList(15) = "Leucopaenia    "
          'InterpList(16) = "Anisocytosis   "
          'InterpList(17) = "Anisocytosis   "
          'InterpList(18) = "Microcytosis   "
          'InterpList(19) = "Macrocytosis   "
          'InterpList(20) = "Hypochromia    "
          'InterpList(21) = "Anaemia        "
          'InterpList(22) = "Erythrocytosis "
          'InterpList(23) = "Thrombopaenia  "
          'InterpList(24) = "Thrombocytosis "

100       Exit Sub

FillInterpTable_Error:

          Dim strES As String
          Dim intEL As Integer

110       intEL = Erl
120       strES = Err.Description
130       LogError "basHaematology", "FillInterpTable", intEL, strES, sql

End Sub

Function Interp(ByVal P As Long, ByVal v As Single) As String

10        On Error GoTo Interp_Error

20        Select Case P
          Case 0: If v < InterpList(0) Then Interp = "Neutropaenia   "
30        Case 1: If v < InterpList(1) Then Interp = "Neutropaenia   "
40        Case 2: If v > InterpList(2) Then Interp = "Neutrophilia   "
50        Case 3: If v > InterpList(3) Then Interp = "Neutrophilia   "
60        Case 4: If v < InterpList(4) Then Interp = "Lymphopaenia   "
70        Case 5: If v < InterpList(5) Then Interp = "Lymphopaenia   "
80        Case 6: If v > InterpList(6) Then Interp = "Lymphocytosis  "
90        Case 7: If v > InterpList(7) Then Interp = "Lymphocytosis  "
100       Case 8: If v > InterpList(8) Then Interp = "Monocytosis    "
110       Case 9: If v > InterpList(9) Then Interp = "Monocytosis    "
120       Case 10: If v > InterpList(10) Then Interp = "Eosinophilia   "
130       Case 11: If v > InterpList(11) Then Interp = "Eosinophilia   "
140       Case 12: If v > InterpList(12) Then Interp = "Basophilia     "
150       Case 13: If v > InterpList(13) Then Interp = "Basophilia     "
160       Case 14: If v > InterpList(14) Then Interp = "Leucocytosis   "
170       Case 15: If v < InterpList(15) Then Interp = "Leucopaenia    "
180       Case 16: If v > InterpList(16) Then Interp = "Anisocytosis   "
190       Case 17: If v > InterpList(17) Then Interp = "Anisocytosis   "
200       Case 18: If v < InterpList(18) Then Interp = "Microcytosis   "
210       Case 19: If v > InterpList(19) Then Interp = "Macrocytosis   "
220       Case 20: If v < InterpList(20) Then Interp = "Hypochromia    "
230       Case 21: If v < InterpList(21) Then Interp = "Anaemia        "
240       Case 22: If v > InterpList(22) Then Interp = "Erythrocytosis "
250       Case 23: If v < InterpList(23) Then Interp = "Thrombopaenia  "
260       Case 24: If v > InterpList(24) Then Interp = "Thrombocytosis "
270       End Select

280       Exit Function

Interp_Error:

          Dim strES As String
          Dim intEL As Integer

290       intEL = Erl
300       strES = Err.Description
310       LogError "basHaematology", "Interp", intEL, strES

End Function

Function InterpH(ByVal Value As Single, _
                 ByVal Analyte As String, _
                 ByVal sex As String, _
                 ByVal Dob As String, _
                 ByVal n As Long, _
                 Optional Rundate As String) _
                 As String

          Dim sql As String
          Dim tb As New Recordset
          Dim DaysOld As Long
          Dim SexSQL As String

10        On Error GoTo InterpH_Error

20        Select Case Left$(UCase$(sex), 1)
          Case "M"
30            SexSQL = "MaleLow as Low, MaleHigh as High "
40        Case "F"
50            SexSQL = "FemaleLow as Low, FemaleHigh as High "
60        Case Else
70            SexSQL = "FemaleLow as Low, MaleHigh as High "
80        End Select

90        If IsDate(Dob) Then

100           If Rundate <> "" Then
110               DaysOld = Abs(DateDiff("d", Format(Rundate, "dd/MMM/yyyy"), Dob))
120           Else
130               DaysOld = Abs(DateDiff("d", Now, Dob))
140           End If
150           sql = "SELECT top 1 PlausibleLow, PlausibleHigh,   " & _
                    SexSQL & _
                    "from HaemTestDefinitions WHERE " & _
                    "AnalyteName = '" & Analyte & "' and AgeFromDays <= '" & DaysOld & "' " & _
                    "and AgeToDays >= '" & DaysOld & "' " & _
                    "order by AgeFromDays desc, AgeToDays asc"
160       Else
170           sql = "SELECT top 1 PlausibleLow, PlausibleHigh, " & _
                    SexSQL & _
                    "from HaemTestDefinitions WHERE Analytename = '" & Analyte & "' " & _
                    "and AgeFromDays <= '9125' " & _
                    "and AgeToDays >= '9125'"
180       End If

190       Set tb = New Recordset
200       RecOpenServer n, tb, sql
210       If Not tb.EOF Then

220           If Value > tb!PlausibleHigh Then
230               InterpH = "X"
240               Exit Function
250           ElseIf Value < tb!PlausibleLow Then
260               InterpH = "X"
270               Exit Function
280           End If

290           If Value > tb!High Then
300               InterpH = "H"
310           ElseIf Value < tb!Low Then
320               InterpH = "L"
330           Else
340               InterpH = " "
350           End If
360       Else
370           InterpH = " "
380       End If

390       Exit Function

InterpH_Error:

          Dim strES As String
          Dim intEL As Integer

400       intEL = Erl
410       strES = Err.Description
420       LogError "basHaematology", "InterpH", intEL, strES, sql

End Function

Public Sub PrintResultHaemWin(ByVal SampleID As String)

          Dim tb As New Recordset
          Dim sql As String
          Dim GP As String
          Dim Clin As String
          Dim Ward As String

10        On Error GoTo PrintResultHaemWin_Error

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
                "Department = 'H' " & _
                "AND SampleID = '" & SampleID & "'"
140       Set tb = New Recordset
150       RecOpenClient 0, tb, sql
160       If tb.EOF Then
170           tb.AddNew
180       End If
190       tb!SampleID = SampleID
200       tb!Department = "H"
210       tb!Initiator = Username
220       tb!Ward = Ward & ""
230       tb!GP = GP & ""
240       tb!Clinician = Clin & ""
250       tb!pTime = Now
260       tb.Update

270       Exit Sub

PrintResultHaemWin_Error:

          Dim strES As String
          Dim intEL As Integer

280       intEL = Erl
290       strES = Err.Description
300       LogError "basHaematology", "PrintResultHaemWin", intEL, strES, sql

End Sub

Function iNr(ByVal Analyte As String, _
             ByVal sex As String, _
             ByVal Dob As String, _
             ByVal n As Long, _
             Optional Rundate As String) _
             As String

          Dim sql As String
          Dim tb As New Recordset
          Dim DaysOld As Long
          Dim SexSQL As String
          Dim fMat As String
          Dim l As String * 4
          Dim h As String * 4

10        On Error GoTo iNr_Error

20        Select Case UCase$(Analyte)
          Case "WBC", "RBC", "HGB", "HCT", "MCV", "MCH", "MCHC", "RDWCV", "FIB", "NEUTA", "LYMA", "MONOA", "EOSA", "BASA": fMat = "#0.0"
30        Case Else: fMat = "####"
40        End Select

50        iNr = "(    -    )"

60        Select Case Left$(UCase$(sex), 1)
          Case "M"
70            SexSQL = "MaleLow as Low, MaleHigh as High "
80        Case "F"
90            SexSQL = "FemaleLow as Low, FemaleHigh as High "
100       Case Else
110           SexSQL = "FemaleLow as Low, MaleHigh as High "
120       End Select

130       If IsDate(Dob) Then

140           If Rundate <> "" Then
150               DaysOld = Abs(DateDiff("d", Format(Rundate, "dd/MMM/yyyy"), Dob))
160           Else
170               DaysOld = Abs(DateDiff("d", Now, Dob))
180           End If
190           sql = "SELECT top 1 PlausibleLow, PlausibleHigh,   " & _
                    SexSQL & _
                    "from HaemTestDefinitions WHERE " & _
                    "AnalyteName = '" & Analyte & "' and AgeFromDays <= '" & DaysOld & "' " & _
                    "and AgeToDays >= '" & DaysOld & "' " & _
                    "order by AgeFromDays desc, AgeToDays asc"
200       Else
210           sql = "SELECT top 1 PlausibleLow, PlausibleHigh, " & _
                    SexSQL & _
                    "from HaemTestDefinitions WHERE Analytename = '" & Analyte & "' " & _
                    "and AgeFromDays = '0' " & _
                    "and AgeToDays = '43830'"
220       End If

230       Set tb = New Recordset
240       RecOpenServer n, tb, sql
250       If Not tb.EOF Then
260           RSet l = Format$(tb!Low, fMat)
270           Mid$(iNr, 2, 4) = l
280           LSet h = Format$(tb!High, fMat)
290           Mid$(iNr, 7, 4) = h
300       End If

310       Exit Function

iNr_Error:

          Dim strES As String
          Dim intEL As Integer

320       intEL = Erl
330       strES = Err.Description
340       LogError "basHaematology", "iNr", intEL, strES, sql

End Function

Public Function HNR(ByVal Analyte As String, _
                    ByVal DaysOld As Long, _
                    ByVal sex As String) As String

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo HNR_Error

20        If sex = "" Then        'QMS reference number #817982
30            HNR = ""
40            Exit Function
50        End If
60        sql = "SELECT * FROM HaemTestDefinitions WHERE " & _
                "AnalyteName = '" & Analyte & "' " & _
                "AND " & DaysOld & " >= agefromdays and " & DaysOld & " <= agetodays "
70        Set tb = New Recordset
80        RecOpenServer 0, tb, sql
90        Do While Not tb.EOF
100           If DaysOld >= tb!AgeFromDays And DaysOld <= tb!AgeToDays Then
110               If Left(sex, 1) = "M" Then
120                   HNR = tb!MaleLow & " - " & tb!MaleHigh
130               ElseIf Left(sex, 1) = "F" Then
140                   HNR = tb!FemaleLow & " - " & tb!FemaleHigh
150               Else
160                   HNR = tb!FemaleLow & " - " & tb!MaleHigh
170               End If
180           End If
190           tb.MoveNext
200       Loop

210       Exit Function

HNR_Error:

          Dim strES As String
          Dim intEL As Integer

220       intEL = Erl
230       strES = Err.Description
240       LogError "basHaematology", "HNR", intEL, strES, sql

End Function
