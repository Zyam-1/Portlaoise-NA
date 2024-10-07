Attribute VB_Name = "basExternals"
Option Explicit
Public SysExt(100) As Boolean
Public SysAandE(100) As Boolean
Public SysMRN(100) As Boolean
Public SysWard(100) As Boolean
Public PrnAll(100) As Boolean

Public Function AreExtResultsPresent(ByVal SampleID As String) As Long

          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo AreExtResultsPresent_Error

20        AreExtResultsPresent = 0

30        If SampleID = "" Then Exit Function

40        sql = "SELECT count(*) as tot from extResults WHERE " & _
                "SampleID = '" & SampleID & "'"
50        Set tb = New Recordset
60        Set tb = Cnxn(0).Execute(sql)

70        AreExtResultsPresent = Sgn(tb!Tot)

80        Exit Function

AreExtResultsPresent_Error:

          Dim strES As String
          Dim intEL As Integer



90        intEL = Erl
100       strES = Err.Description
110       LogError "basExternals", "AreExtResultsPresent", intEL, strES, sql

End Function

Function eName2Normal(ByVal s As String, _
                      ByVal Department As String) _
                      As String

          Dim sql As String
          Dim tb As New Recordset

10        On Error GoTo eName2Normal_Error

20        eName2Normal = ""

30        sql = "SELECT NormalRange FROM ExternalDefinitions WHERE " & _
                "AnalyteName = '" & AddTicks(s) & "' " & _
                "AND Department = '" & Department & "'"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        If Not tb.EOF Then
70            eName2Normal = tb!NormalRange & ""
80        End If

90        Exit Function

eName2Normal_Error:

          Dim strES As String
          Dim intEL As Integer

100       intEL = Erl
110       strES = Err.Description
120       LogError "basExternals", "eName2Normal", intEL, strES, sql

End Function

Function eName2SendTo(ByVal s As String, _
                      Optional ByVal Department As String = "") _
                      As String

          Dim sql As String
          Dim tb As New Recordset

10        On Error GoTo eName2SendTo_Error

20        eName2SendTo = ""
          'Zyam
          If Trim(Department) = "" Then
30            sql = "SELECT SendTo FROM ExternalDefinitions WHERE " & _
                "AnalyteName = '" & AddTicks(s) & "' "
          Else
31        sql = "SELECT SendTo FROM ExternalDefinitions WHERE " & _
                "AnalyteName = '" & AddTicks(s) & "' " & _
                "AND Department = '" & Department & "'"

          End If


40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        If Not tb.EOF Then
70            eName2SendTo = tb!SendTo & ""
80        Else
90            eName2SendTo = s
100       End If

110       Exit Function

eName2SendTo_Error:

          Dim strES As String
          Dim intEL As Integer

120       intEL = Erl
130       strES = Err.Description
140       LogError "basExternals", "eName2SendTo", intEL, strES

End Function

Function eName2Units(ByVal s As String, _
                     ByVal Department As String) _
                     As String

          Dim sql As String
          Dim tb As New Recordset

10        On Error GoTo eName2Units_Error

20        eName2Units = 0

30        sql = "SELECT Units FROM ExternalDefinitions WHERE " & _
                "AnalyteName = '" & AddTicks(s) & "' " & _
                "AND Department = '" & Department & "'"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        If Not tb.EOF Then
70            eName2Units = tb!Units & ""
80        End If

90        Exit Function

eName2Units_Error:

          Dim strES As String
          Dim intEL As Integer

100       intEL = Erl
110       strES = Err.Description
120       LogError "basExternals", "eName2Units", intEL, strES, sql

End Function

Function eNumber2Name(ByVal x As String, _
                      ByVal Department As String) _
                      As String

          Dim sql As String
          Dim tb As New Recordset

10        On Error GoTo eNumber2Name_Error

20        eNumber2Name = "???"

30        sql = "SELECT AnalyteName FROM ExternalDefinitions WHERE " & _
                "MBCode = '" & x & "' " & _
                "AND Department = '" & Department & "'"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        If Not tb.EOF Then
70            eNumber2Name = Trim(tb!AnalyteName & "")
80        End If

90        Exit Function

eNumber2Name_Error:

          Dim strES As String
          Dim intEL As Integer

100       intEL = Erl
110       strES = Err.Description
120       LogError "basExternals", "eNumber2Name", intEL, strES, sql

End Function
