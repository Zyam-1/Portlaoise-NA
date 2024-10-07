Attribute VB_Name = "modAutoComments"
Option Explicit

Public Function CheckAutoComments(ByVal SampleID As String, ByVal Discipline As String) As String

          Dim tb As Recordset
          Dim sql As String
          Dim ShortDisc As String
          Dim RetVal As String
          Dim i As Integer
          Dim Obs As Observations
          Dim Ob As Observation
          Dim CurrentComment As String
          Dim SampleValidated As Boolean

10        On Error GoTo CheckAutoComments_Error

20        RetVal = ""
30        SampleValidated = False

40        If Discipline = "Biochemistry" Then
50            ShortDisc = "Bio"
60        ElseIf Discipline = "Coagulation" Then
70            ShortDisc = "Coag"
80        ElseIf Discipline = "Immunology" Then
90            ShortDisc = "Imm"
100       ElseIf Discipline = "Endocrinology" Then
110           ShortDisc = "End"
120       End If
          
130       Set Obs = New Observations
140           Set Obs = Obs.Load(SampleID, Discipline)
150           If Not Obs Is Nothing Then
160               For Each Ob In Obs
170                   If Ob.Discipline = Discipline Then
180                       CurrentComment = Ob.Comment
190                   End If
200               Next
210           End If

220       sql = "SELECT Parameter, COALESCE(Valid, 0) Valid, COALESCE(Printed,0) Printed, 'Output' = " & _
                "CASE WHEN ISNUMERIC(R.Result) = 1 AND R.Result <> '.' AND R.Result <> '+' AND R.Result <> '-' " & _
                "  THEN " & _
                "    CASE " & _
                "      WHEN Criteria = 'Present' THEN A.Comment " & _
                "      WHEN Criteria = 'Equal to' AND CONVERT(float, R.Result) = CONVERT(float, A.Value0) THEN A.Comment " & _
                "      WHEN Criteria = 'Less than' AND CONVERT(float, R.Result) < CONVERT(float, A.Value0) THEN A.Comment " & _
                "      WHEN Criteria = 'Greater than' AND CONVERT(float, R.Result) > CONVERT(float, A.Value0) THEN A.Comment " & _
                "      WHEN Criteria = 'Between' AND CONVERT(float, R.Result) > CONVERT(float, A.Value0) AND CONVERT(float, R.Result) < CONVERT(float, A.Value1) THEN A.Comment " & _
                "      WHEN Criteria = 'Not between' AND (CONVERT(float, R.Result) < CONVERT(float, A.Value0) OR CONVERT(float, R.Result) > CONVERT(float, A.Value1)) THEN A.Comment " & _
                "      ELSE '' " & _
                "    END " & _
                "  ELSE " & _
                "    CASE " & _
                "      WHEN Criteria = 'Contains Text' AND CHARINDEX( A.Value0 COLLATE Latin1_General_CI_AS, R.Result) > 0 THEN A.Comment " & _
                "      WHEN Criteria = 'Starts with' AND LEFT(R.Result, 1) = A.Value0 COLLATE Latin1_General_CI_AS THEN A.Comment " & _
                "      ELSE '' " & _
                "    END " & _
                "END, COALESCE(CommentType, 0) CommentType "
230       sql = sql & "FROM AutoComments A JOIN " & ShortDisc & "Results R ON " & _
                "R.Code = (SELECT TOP 1 Code FROM " & ShortDisc & "TestDefinitions " & _
                "          WHERE ShortName = A.Parameter COLLATE Latin1_General_CI_AS " & _
                "          AND InUse = 1 ) " & _
                "WHERE A.Discipline = '" & Discipline & "' " & _
                "AND R.SampleID = '" & SampleID & "' " & _
                "AND R.RunTime Between A.DateStart AND A.DateEnd"

240       Set tb = New Recordset
250       RecOpenClient 0, tb, sql
260       Do While Not tb.EOF
270           If Trim$(tb!Output & "") <> "" Then
280               If tb!CommentType = 0 Then
290                   If InStr(RetVal, tb!Output) = 0 And InStr(CurrentComment, tb!Output) = 0 Then
300                       RetVal = RetVal & tb!Output & vbCrLf
310                   End If
320                   If tb!Valid = 1 Or tb!Printed = 1 Then
330                       SampleValidated = True
340                   End If
350               ElseIf tb!CommentType = 1 And tb!Valid = 0 And tb!Printed = 0 Then
360                   sql = "UPDATE " & ShortDisc & "Results " & _
                            "SET Comment = '" & Left$(tb!Output, 100) & "' WHERE " & _
                            "SampleID = '" & SampleID & "' " & _
                            "AND Code = " & _
                            "  (SELECT Top 1 Code FROM " & ShortDisc & "TestDefinitions WHERE " & _
                            "   ShortName = '" & tb!Parameter & "' ) " & _
                            "AND COALESCE(Valid, 0) = 0 AND COALESCE(Printed, 0) = 0"
370                   Cnxn(0).Execute sql
380               End If

390           End If

400           tb.MoveNext
410       Loop

420       If Trim$(RetVal) <> "" Then
      '        Set Obs = New Observations
      '        Set Obs = Obs.Load(SampleID, Discipline)
      '        If Not Obs Is Nothing Then
      '            For Each Ob In Obs
      '                If Ob.Discipline = Discipline Then
      '                    CurrentComment = Ob.Comment
      '                End If
      '            Next
      '        End If
      '        If InStr(CurrentComment, RetVal) = 0 And SampleValidated = False Then
      '            Set Obs = New Observations
      '            Obs.Save SampleID, False, Discipline, RetVal
      '        End If
430           Set Obs = New Observations
440               Obs.Save SampleID, False, Discipline, RetVal
450       End If

460       CheckAutoComments = Trim$(RetVal)

470       Exit Function

CheckAutoComments_Error:

          Dim strES As String
          Dim intEL As Integer

480       intEL = Erl
490       strES = Err.Description
500       LogError "modAutoComments", "CheckAutoComments", intEL, strES, sql

End Function


