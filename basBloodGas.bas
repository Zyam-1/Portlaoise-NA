Attribute VB_Name = "basBloodGas"
Option Explicit

Private Type hl
    Low As Single
    High As Single
End Type


Private Type BGR
    pH As hl
    PCO2 As hl
    PO2 As hl
    HCO3 As hl
    O2SAT As hl
    BE As hl
    TotCO2 As hl
End Type

Public BGRanges As BGR

Public Function AreBgaResultsPresent(ByVal SampleID As String) As Long
      'check if bloodgas present
          Dim tb As New Recordset
          Dim sql As String


10        On Error GoTo AreBgaResultsPresent_Error

20        sql = "SELECT count(*) as tot from BgaResults WHERE " & _
                "SampleID = '" & SampleID & "'"
30        Set tb = New Recordset
40        Set tb = Cnxn(0).Execute(sql)

50        AreBgaResultsPresent = Sgn(tb!Tot)






60        Exit Function

AreBgaResultsPresent_Error:

          Dim strES As String
          Dim intEL As Integer



70        intEL = Erl
80        strES = Err.Description
90        LogError "basBloodGas", "AreBgaResultsPresent", intEL, strES, sql


End Function

Public Function BgaCodeForShortName(ByVal ShortName As String) _
       As String

          Dim tb As New Recordset
          Dim sql As String


10        On Error GoTo BgaCodeForShortName_Error

20        BgaCodeForShortName = "???"

30        sql = "SELECT * from bgaTestDefinitions WHERE ShortName = '" & ShortName & "'"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        If Not tb.EOF Then
70            BgaCodeForShortName = Trim(tb!Code)
80        End If




90        Exit Function

BgaCodeForShortName_Error:

          Dim strES As String
          Dim intEL As Integer



100       intEL = Erl
110       strES = Err.Description
120       LogError "basBloodGas", "BgaCodeForShortName", intEL, strES, sql


End Function

Public Function LongNameforBgaCode(ByVal Code As String) _
       As String

          Dim tb As New Recordset
          Dim sql As String


10        On Error GoTo LongNameforBgaCode_Error

20        LongNameforBgaCode = "???"

30        sql = "SELECT * from BgaTestDefinitions WHERE Code = '" & Code & "'"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        If Not tb.EOF Then
70            LongNameforBgaCode = Trim(tb!LongName)
80        End If



90        Exit Function

LongNameforBgaCode_Error:

          Dim strES As String
          Dim intEL As Integer



100       intEL = Erl
110       strES = Err.Description
120       LogError "basBloodGas", "LongNameforBgaCode", intEL, strES, sql


End Function

Public Sub PrintResultBGA(ByVal SampleID As String)

          Dim tb As New Recordset
          Dim sql As String
          Dim Ward As String
          Dim GP As String
          Dim Clin As String

10        On Error GoTo PrintResultBGA_Error

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
                "Department = 'G' " & _
                "AND SampleID = '" & SampleID & "'"
140       Set tb = New Recordset
150       RecOpenClient 0, tb, sql
160       If tb.EOF Then
170           tb.AddNew
180       End If
190       tb!SampleID = SampleID
200       tb!Department = "G"
210       tb!Initiator = UserName
220       tb!Ward = Ward & ""
230       tb!GP = GP & ""
240       tb!Clinician = Clin & ""

250       tb!pTime = Now
260       tb.Update

270       Exit Sub

PrintResultBGA_Error:

          Dim strES As String
          Dim intEL As Integer



280       intEL = Erl
290       strES = Err.Description
300       LogError "basBloodGas", "PrintResultBGA", intEL, strES, sql

End Sub

