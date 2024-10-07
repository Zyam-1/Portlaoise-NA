Attribute VB_Name = "modSql"
Option Explicit

Public Function Create_Graph(ByVal Chart As String, ByVal From As String, ByVal ToDate As String) As String
      Dim sql As String

10    On Error GoTo Create_Graph_Error

20    sql = "SELECT rundate, sampleid from demographics " & _
            "WHERE chart = '" & Chart & "' " & _
            "and rundate between '" & _
            Format(From, "dd/mmm/yyyy") & "' and '" & _
            Format(ToDate, "dd/mmm/yyyy") & "' " & _
            "order by rundate"

30    Create_Graph = sql

40    Exit Function

Create_Graph_Error:

      Dim strES As String
      Dim intEL As Integer


50    intEL = Erl
60    strES = Err.Description
70    LogError "modSql", "Create_Graph", intEL, strES


End Function

