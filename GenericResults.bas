Attribute VB_Name = "GenericResults"
Public Type GenericResult
    SampleID As String
    TestName As String
    Result As String
    Username As String
    HealthLink As String
    TestDateTime As String
    DateTimeOfRecord As String
    Valid As Boolean
    Printed As Boolean
End Type


Public Function LoadGenericResult(SampleID As String, TestName As String) As GenericResult

          Dim tb As Recordset
          Dim sql As String
          Dim GR As GenericResult

10        On Error GoTo LoadGenericResult_Error

20        sql = "Select * From GenericResults Where SampleID = '%sampleid' And TestName = '%testname'"
30        sql = Replace(sql, "%sampleid", SampleID)
40        sql = Replace(sql, "%testname", TestName)

50        Set tb = New Recordset
60        RecOpenClient 0, tb, sql
70        If Not tb.EOF Then
80            GR.SampleID = tb!SampleID & ""
90            GR.TestName = tb!TestName & ""
100           GR.Result = tb!Result & ""
110           GR.Username = tb!Username & ""
120           GR.HealthLink = tb!HealthLink & ""
130           GR.TestDateTime = tb!TestDateTime & ""
140           GR.DateTimeOfRecord = tb!DateTimeOfRecord & ""
150           GR.Valid = IIf(IsNull(tb!Valid), 0, tb!Valid)
160           GR.Printed = IIf(IsNull(tb!Printed), 0, tb!Printed)
170       End If

180       LoadGenericResult = GR


190       Exit Function

LoadGenericResult_Error:

          Dim strES As String
          Dim intEL As Integer

200       intEL = Erl
210       strES = Err.Description
220       LogError "GenericResults", "LoadGenericResult", intEL, strES, sql

End Function
