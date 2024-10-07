Attribute VB_Name = "basImmunology"
Option Explicit
Public Function Check_Imm(ByVal Test As String, _
                          ByVal Units As String, _
                          ByVal SampleType As String) As String

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo Check_Imm_Error

20        Check_Imm = ""

30        sql = "SELECT SampleType, Units FROM ImmTestDefinitions WHERE " & _
                "ShortName = '" & Test & "'"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        If tb.EOF Then
70            Check_Imm = "Test Name"
80        Else
90            If ListCodeFor("ST", SampleType) <> tb!SampleType Then
100               Check_Imm = "Sample Type"
110           ElseIf Units <> tb!Units Then
120               Check_Imm = "Units"
130           End If
140       End If

150       Exit Function

Check_Imm_Error:

          Dim strES As String
          Dim intEL As Integer

160       intEL = Erl
170       strES = Err.Description
180       LogError "basImmunology", "Check_Imm", intEL, strES, sql

End Function
Public Function AreImmResultsPresent(ByVal SampleID As String) As Long

          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo AreImmResultsPresent_Error

20        sql = "SELECT count(*) as tot from ImmResults WHERE " & _
                "SampleID = '" & Val(SampleID) & "'"
30        Set tb = New Recordset
40        Set tb = Cnxn(0).Execute(sql)

50        AreImmResultsPresent = Sgn(tb!Tot)

60        Exit Function

AreImmResultsPresent_Error:

          Dim strES As String
          Dim intEL As Integer

70        intEL = Erl
80        strES = Err.Description
90        LogError "basImmunology", "AreImmResultsPresent", intEL, strES, sql

End Function

Public Function iCodeForLongName(ByVal LongName As String) _
       As String

          Dim sql As String
          Dim tb As New Recordset

10        On Error GoTo iCodeForLongName_Error

20        iCodeForLongName = "???"

30        sql = "SELECT Code from ImmTestDefinitions WHERE " & _
                "LongName = '" & LongName & "' " & _
                "AND InUse = 1"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        If Not tb.EOF Then
70            iCodeForLongName = Trim(tb!Code)
80        End If

90        Exit Function

iCodeForLongName_Error:

          Dim strES As String
          Dim intEL As Integer

100       intEL = Erl
110       strES = Err.Description
120       LogError "basImmunology", "iCodeForLongName", intEL, strES, sql

End Function

Public Function ICodeForShortName(ByVal ShortName As String) _
       As String

          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo ICodeForShortName_Error

20        ICodeForShortName = "???"

30        sql = "SELECT Code from ImmTestDefinitions WHERE " & _
                "ShortName = '" & ShortName & "' " & _
                "AND InUse = '1'"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        If Not tb.EOF Then
70            ICodeForShortName = Trim(tb!Code)
80        End If

90        Exit Function

ICodeForShortName_Error:

          Dim strES As String
          Dim intEL As Integer

100       intEL = Erl
110       strES = Err.Description
120       LogError "basImmunology", "ICodeForShortName", intEL, strES, sql

End Function

Public Function ImmLongNameFor(ByVal Code As String) As String

          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo ImmLongNameFor_Error

20        ImmLongNameFor = "???"

30        sql = "SELECT LongName from ImmTestDefinitions WHERE " & _
                "Code = '" & Code & "'"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        If Not tb.EOF Then
70            ImmLongNameFor = Trim(tb!LongName)
80        End If

90        Exit Function

ImmLongNameFor_Error:

          Dim strES As String
          Dim intEL As Integer

100       intEL = Erl
110       strES = Err.Description
120       LogError "basImmunology", "ImmLongNameFor", intEL, strES, sql

End Function

Public Function ImmShortNameFor(ByVal Code As String) As String

          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo ImmShortNameFor_Error

20        ImmShortNameFor = "???"

30        sql = "SELECT ShortName FROM ImmTestDefinitions WHERE " & _
                "Code = '" & Code & "'"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        If Not tb.EOF Then
70            ImmShortNameFor = Trim(tb!ShortName)
80        End If

90        Exit Function

ImmShortNameFor_Error:

          Dim strES As String
          Dim intEL As Integer

100       intEL = Erl
110       strES = Err.Description
120       LogError "basImmunology", "ImmShortNameFor", intEL, strES, sql

End Function

Function ImmTestAffected(ByVal br As BIEResult) As Boolean

          Dim TestName As String
          Dim tb As New Recordset
          Dim sql As String
          Dim sn As New Recordset

10        On Error GoTo ImmTestAffected_Error

20        ImmTestAffected = False
30        TestName = Trim(br.LongName)

40        sql = "SELECT * from masks WHERE sampleid = '" & br.SampleID & "'"
50        Set tb = New Recordset
60        RecOpenServer 0, tb, sql

70        If tb.EOF Then Exit Function

80        sql = "SELECT * from Immtestdefinitions WHERE code = '" & br.Code & "'"
90        Set sn = New Recordset
100       RecOpenServer 0, sn, sql
110       Do While Not sn.EOF
120           With sn
130               If !LongName = TestName And !Code = br.Code Then
140                   If !h And tb!h Then
150                       ImmTestAffected = True
160                       Exit Do
170                   End If
180                   If !s And tb!s Then
190                       ImmTestAffected = True
200                       Exit Do
210                   End If
220                   If !l And tb!l Then
230                       ImmTestAffected = True
240                       Exit Do
250                   End If
260                   If !o And tb!o Then
270                       ImmTestAffected = True
280                       Exit Do
290                   End If
300                   If !g And tb!g Then
310                       ImmTestAffected = True
320                       Exit Do
330                   End If
340                   If !J And tb!J Then
350                       ImmTestAffected = True
360                       Exit Do
370                   End If
380               End If
390           End With
400           sn.MoveNext
410       Loop

420       Exit Function

ImmTestAffected_Error:

          Dim strES As String
          Dim intEL As Integer

430       intEL = Erl
440       strES = Err.Description
450       LogError "basImmunology", "ImmTestAffected", intEL, strES, sql

End Function

Public Sub PrintResultImmWin(ByVal SampleID As String)

          Dim tb As New Recordset
          Dim sql As String
          Dim GP As String
          Dim Clin As String
          Dim Ward As String

10        On Error GoTo PrintResultImmWin_Error

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
                "Department = 'I' " & _
                "AND SampleID = '" & SampleID & "'"
140       Set tb = New Recordset
150       RecOpenClient 0, tb, sql
160       If tb.EOF Then
170           tb.AddNew
180       End If
190       tb!SampleID = SampleID
200       tb!Department = "I"
210       tb!Initiator = Username
220       tb!pTime = Now
230       tb!Ward = Ward & ""
240       tb!GP = GP & ""
250       tb!Clinician = Clin & ""
260       tb.Update

270       Exit Sub

PrintResultImmWin_Error:

          Dim strES As String
          Dim intEL As Integer



280       intEL = Erl
290       strES = Err.Description
300       LogError "basImmunology", "PrintResultImmWin", intEL, strES, sql

End Sub

Public Function QuickInterpImm(ByVal Result As BIEResult) _
       As String

10        On Error GoTo QuickInterpImm_Error

20        With Result
30            If IsNumeric(.Result) Then
40                If Val(.Result) > .PlausibleHigh Then
50                    QuickInterpImm = "***"
60                ElseIf Val(.Result) < .PlausibleLow Then
70                    QuickInterpImm = "***"
80                ElseIf Val(.Result) < .Low Then
90                    QuickInterpImm = "Low "
100               ElseIf Val(.Result) > .High Then
110                   QuickInterpImm = "High"
120               Else
130                   QuickInterpImm = "    "
140               End If
150           Else
160               QuickInterpImm = "    "
170           End If
180       End With

190       Exit Function

QuickInterpImm_Error:

          Dim strES As String
          Dim intEL As Integer

200       intEL = Erl
210       strES = Err.Description
220       LogError "basImmunology", "QuickInterpImm", intEL, strES

End Function



Public Function GetImmUnit(ByVal Code As String) As String

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo GetImmUnit_Error

20        GetImmUnit = ""

30        sql = "SELECT Units FROM ImmTestDefinitions WHERE " & _
                "Code = '" & Code & "'"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        If Not tb.EOF Then
70            GetImmUnit = tb!Units
80        End If

90        Exit Function

GetImmUnit_Error:

          Dim strES As String
          Dim intEL As Integer

100       intEL = Erl
110       strES = Err.Description
120       LogError "basImmunology", "GetImmUnit", intEL, strES, sql

End Function
