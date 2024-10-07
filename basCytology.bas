Attribute VB_Name = "basCytology"
Option Explicit

Public Function AreCytoResultsPresent(ByVal SampleID As String, ByVal Year As String) As Long

          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo AreCytoResultsPresent_Error

20        sql = "SELECT count(*) as tot from Cytoresults WHERE " & _
                "SampleID = '" & SampleID & "' and hyear = '" & Year & "'"
30        Set tb = New Recordset
40        Set tb = Cnxn(0).Execute(sql)

50        AreCytoResultsPresent = Sgn(tb!Tot)

60        Exit Function

AreCytoResultsPresent_Error:

          Dim strES As String
          Dim intEL As Integer



70        intEL = Erl
80        strES = Err.Description
90        LogError "basCytology", "AreCytoResultsPresent", intEL, strES, sql

End Function

