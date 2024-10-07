VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CoagResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private mcolCoagResults As New Collection
Private mCoagResults As New CoagResults

Public Function Add(CR As CoagResult) As CoagResult

10        On Error GoTo Add_Error

20        mcolCoagResults.Add CR

30        Set Add = CR

40        Exit Function

Add_Error:

          Dim strES As String
          Dim intEL As Integer


50        intEL = Erl
60        strES = Err.Description
70        LogError "CoagResults", "Add", intEL, strES

End Function

Public Property Get CoagResults() As CoagResults

10        On Error GoTo CoagResults_Error

20        Set CoagResults = mCoagResults

30        Exit Property

CoagResults_Error:

          Dim strES As String
          Dim intEL As Integer


40        intEL = Erl
50        strES = Err.Description
60        LogError "CoagResults", "CoagResults", intEL, strES

End Property

Public Sub Clear()

          Dim n As Long

10        On Error GoTo Clear_Error

20        For n = 1 To mcolCoagResults.Count
30            mcolCoagResults.Remove 1
40        Next

50        Exit Sub

Clear_Error:

          Dim strES As String
          Dim intEL As Integer


60        intEL = Erl
70        strES = Err.Description
80        LogError "CoagResults", "Clear", intEL, strES

End Sub

Public Function Count() As Long

10        On Error GoTo Count_Error

20        Count = mcolCoagResults.Count

30        Exit Function

Count_Error:

          Dim strES As String
          Dim intEL As Integer


40        intEL = Erl
50        strES = Err.Description
60        LogError "CoagResults", "Count", intEL, strES

End Function


Public Sub DELETE(ByVal SampleID As String, _
                  ByVal Code As String, ByVal Units As String)

          Dim sql As String

10        On Error GoTo DELETE_Error

20        sql = "DELETE from CoagResults WHERE " & _
                "SampleID = '" & SampleID & "' " & _
                "and Code = '" & Code & "' and units = '" & Units & "'"

30        Cnxn(0).Execute sql

40        Exit Sub

DELETE_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "CoagResults", "DELETE", intEL, strES, sql

End Sub

Public Function Item(ByVal X As Long) _
       As CoagResult
Attribute Item.VB_UserMemId = 0

10        On Error GoTo Item_Error

20        Set Item = mcolCoagResults(X)

30        Exit Function

Item_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "CoagResults", "Item", intEL, strES

End Function


Public Function Load(ByVal SampleID As String, _
                     ByVal v As Long, _
                     ByVal P As Long, _
                     Optional ByVal Units As String, Optional ByVal n As Long) _
                     As CoagResults

          Dim CRs As New CoagResults
          Dim CR As CoagResult
          Dim sql As String
          Dim tb As New Recordset

10        On Error GoTo Load_Error

20        sql = "SELECT DISTINCT R.*, " & _
                "D.PrintPriority, D.CheckTime, D.DoDelta, D.DeltaLimit " & _
                "FROM  CoagResults R INNER JOIN " & _
                "CoagTestDefinitions D ON (R.Code = D.Code OR R.Code = D.TestName) " & _
                "WHERE     (R.SampleID = '" & Val(SampleID) & "') "

30        If v = gVALID And P = gNOTPRINTED Then
40            sql = sql & "AND Valid = 1 AND Printed = 0 "
50        ElseIf v = gVALID And P = gDONTCARE Then
60            sql = sql & "AND Valid = 1 "
70        ElseIf v = gNOTVALID And P = gDONTCARE Then
80            sql = sql & "AND Valid = 0 "
90        End If

100       sql = sql & " ORDER BY D.PrintPriority"

110       Set tb = New Recordset
120       RecOpenServer n, tb, sql

130       Do While Not tb.EOF
140           Set CR = New CoagResult
150           With CR
160               .SampleID = tb!SampleID & ""
170               .Code = tb!Code & ""
180               .Result = tb!Result & ""
190               .Rundate = Format$(tb!Rundate, "dd/mm/yyyy")
200               .RunTime = Format$(tb!RunTime, "dd/mm/yyyy hh:mm")
210               .Printed = tb!Printed
220               .Valid = tb!Valid
230               .Units = tb!Units & ""
240               .Analyser = tb!Analyser & ""
250               If tb!DoDelta & "" <> "" Then .DoDelta = tb!DoDelta Else .DoDelta = False
260               .DeltaLimit = IIf(IsNull(tb!DeltaLimit), 9999, tb!DeltaLimit)
270               If IsNull(tb!CheckTime) Then
280                   .CheckTime = 1
290               Else
300                   .CheckTime = tb!CheckTime
310               End If
320               CRs.Add CR
330           End With
340           tb.MoveNext
350       Loop

360       If CRs.Count <> 0 Then
370           Set Load = CRs
380       Else
390           Set Load = Nothing
400       End If

410       Exit Function

Load_Error:

          Dim strES As String
          Dim intEL As Integer

420       intEL = Erl
430       strES = Err.Description
440       LogError "CoagResults", "Load", intEL, strES, sql

End Function
Public Function LoadRepeats(ByVal SampleID As String, _
                            ByVal v As Long, _
                            ByVal P As Long, _
                            Optional ByVal Units As String) As CoagResults

          Dim CRs As New CoagResults
          Dim CR As CoagResult
          Dim sql As String
          Dim tb As New Recordset

10        On Error GoTo LoadRepeats_Error

20        sql = "SELECT CoagRepeats.SampleID, CoagRepeats.Result,  " & _
                "CoagRepeats.RunDate, CoagRepeats.Valid, " & _
                "CoagRepeats.Printed, CoagTestDefinitions.code "

30        If Units = "" Or Units = "False" Then sql = sql & ",CoagTestDefinitions.Units " Else sql = sql & ", CoagRepeats.units"

40        sql = sql & ", PrintPriority from CoagRepeats, CoagTestDefinitions  " & _
                "WHERE SampleID = '" & SampleID & "' " & _
                "and CoagRepeats.Code  = CoagTestDefinitions.Code  "

50        If Units <> "" And Units <> "False" Then
60            sql = sql & " and (CoagRepeats.units = coagtestdefinitions.units)"
70        End If

80        If v = gVALID And P = gNOTPRINTED Then
90            sql = sql & "and valid = 1 and printed = 0 "
100       ElseIf v = gVALID And P = gDONTCARE Then
110           sql = sql & "and valid = 1 "
120       ElseIf v = gNOTVALID And P = gDONTCARE Then
130           sql = sql & "and valid = 0 "
140       End If

150       sql = "SELECT CoagRepeats.* " & _
                "FROM         CoagRepeats " & _
                "WHERE     (CoagRepeats.SampleID = '" & SampleID & "') "

160       If v = gVALID And P = gNOTPRINTED Then
170           sql = sql & "and valid = 1 and printed = 0 "
180       ElseIf v = gVALID And P = gDONTCARE Then
190           sql = sql & "and valid = 1 "
200       ElseIf v = gNOTVALID And P = gDONTCARE Then
210           sql = sql & "and valid = 0 "
220       End If

230       Set tb = New Recordset
240       RecOpenServer 0, tb, sql
250       If tb.EOF Then Exit Function
260       Do While Not tb.EOF
270           Set CR = New CoagResult
280           With CR
290               .SampleID = tb!SampleID & ""
300               .Code = tb!Code & ""
310               .Result = tb!Result & ""
320               .Rundate = Format$(tb!Rundate, "dd/mm/yyyy")
330               .Printed = tb!Printed
340               .Valid = tb!Valid
350               .Units = tb!Units
360               CRs.Add CR
370           End With
380           tb.MoveNext
390       Loop

400       If CRs.Count <> 0 Then
410           Set LoadRepeats = CRs
420       Else
430           Set LoadRepeats = Nothing
440       End If

450       Exit Function

LoadRepeats_Error:

          Dim strES As String
          Dim intEL As Integer

460       intEL = Erl
470       strES = Err.Description
480       LogError "CoagResults", "LoadRepeats", intEL, strES, sql

End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"

10        On Error GoTo NewEnum_Error

20        Set NewEnum = mcolCoagResults.[_NewEnum]

30        Exit Function

NewEnum_Error:

          Dim strES As String
          Dim intEL As Integer


40        intEL = Erl
50        strES = Err.Description
60        LogError "CoagResults", "NewEnum", intEL, strES

End Function



Public Sub Remove(ByVal Index As Integer)

10        On Error GoTo Remove_Error

20        mcolCoagResults.Remove (Index)

30        Exit Sub

Remove_Error:

          Dim strES As String
          Dim intEL As Integer


40        intEL = Erl
50        strES = Err.Description
60        LogError "CoagResults", "Remove", intEL, strES

End Sub

Public Sub Save(ByVal CRs As CoagResults)

          Dim sql As String
          Dim tb As New Recordset
          Dim CR As CoagResult
          Dim lngSampleID As Long
          Dim Rundate As String

10        On Error GoTo Save_Error

20        For Each CR In CRs
30            With CR
40                If IsNumeric(.Code) Then
50                    lngSampleID = Val(.SampleID)
60                    sql = "SELECT * from CoagResults WHERE " & _
                            "SampleID = '" & lngSampleID & "' " & _
                            "and  Code = '" & .Code & "' and units = '" & .Units & "'"
70                    Set tb = New Recordset
80                    RecOpenServer 0, tb, sql
90                    If Not tb.EOF Then
100                       sql = "SELECT * from CoagRepeats WHERE " & _
                                "SampleID = '" & lngSampleID & "' " & _
                                "and  Code = '" & CR.Code & "' and units = '" & .Units & "'"
110                       Set tb = New Recordset
120                       RecOpenServer 0, tb, sql
130                   End If
140                   tb.AddNew
150                   tb!SampleID = lngSampleID
160                   Rundate = Format$(.Rundate, "dd/mmm/yyyy")
170                   tb!Rundate = Rundate
180                   If Trim$(.RunTime) <> "" Then
190                       tb!RunTime = Format$(.RunTime, "dd/mmm/yyyy hh:mm")
200                   End If
210                   tb!Code = Left$(.Code, 4)
220                   If IsNumeric(.Result) Then
230                       tb!Result = Left$(Format$(Val(.Result)), 6)
240                   Else
250                       tb!Result = Left$(.Result, 6)
260                   End If
270                   tb!Printed = IIf(.Printed, 1, 0)
280                   tb!Valid = IIf(.Valid, 1, 0)
290                   tb!Units = .Units
300                   tb.Update
310                   sql = "DELETE from CoagRequests WHERE " & _
                            "sampleid = '" & lngSampleID & "' " & _
                            "and Code = '" & .Code & "' and units = '" & .Units & "'"
320                   Cnxn(0).Execute sql
330               End If
340           End With
350       Next

360       Set CR = Nothing

370       Exit Sub

Save_Error:

          Dim strES As String
          Dim intEL As Integer

380       intEL = Erl
390       strES = Err.Description
400       LogError "CoagResults", "Save", intEL, strES, sql

End Sub


Public Sub LogAsPrinted(ByVal Tn As Long, ByVal SampleID As String)

          Dim sql As String

10        On Error GoTo LogAsPrinted_Error

20        sql = "UPDATE CoagResults set Printed = 1 WHERE " & _
                "SampleID = '" & SampleID & "'"
30        Cnxn(Tn).Execute sql

40        Exit Sub

LogAsPrinted_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "CoagResults", "LogAsPrinted", intEL, strES, sql

End Sub



