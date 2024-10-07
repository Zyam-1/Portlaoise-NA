VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BIEResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private mcolBIEResults As New Collection
Private mBIEResults As New BIEResults

Public Function Add(br As BIEResult) As BIEResult

10        On Error GoTo Add_Error

20        mcolBIEResults.Add br

30        Set Add = br

40        Exit Function

Add_Error:

          Dim strES As String
          Dim intEL As Integer


50        intEL = Erl
60        strES = Err.Description
70        LogError "BIEResults", "Add", intEL, strES


End Function

Public Property Get BIEResults() As BIEResults

10        On Error GoTo BIEResults_Error

20        Set BIEResults = mBIEResults

30        Exit Property

BIEResults_Error:

          Dim strES As String
          Dim intEL As Integer


40        intEL = Erl
50        strES = Err.Description
60        LogError "BIEResults", "BIEResults", intEL, strES


End Property

Public Sub Clear()

          Dim n As Long

10        On Error GoTo Clear_Error

20        For n = 1 To mcolBIEResults.Count
30            mcolBIEResults.Remove 1
40        Next

50        Exit Sub

Clear_Error:

          Dim strES As String
          Dim intEL As Integer


60        intEL = Erl
70        strES = Err.Description
80        LogError "BIEResults", "Clear", intEL, strES


End Sub

Public Function Count() As Long

10        On Error GoTo Count_Error

20        Count = mcolBIEResults.Count

30        Exit Function

Count_Error:

          Dim strES As String
          Dim intEL As Integer


40        intEL = Erl
50        strES = Err.Description
60        LogError "BIEResults", "Count", intEL, strES


End Function

Public Function Item(ByVal X As Long) _
       As BIEResult
Attribute Item.VB_UserMemId = 0

10        On Error GoTo Item_Error

20        On Error Resume Next

30        Set Item = mcolBIEResults(X)

40        Exit Function

Item_Error:

          Dim strES As String
          Dim intEL As Integer


50        intEL = Erl
60        strES = Err.Description
70        LogError "BIEResults", "Item", intEL, strES


End Function

'Public Function Load(ByVal Discipline As String, _
'                     ByVal SampleID As String, _
'                     ByVal ResultOrRepeat As String, _
'                     ByVal v As Integer, _
'                     ByVal P As Integer, _
'                     Optional ByVal Cat As String, _
'                     Optional ByVal Connection As Integer, _
'                     Optional ByVal Sex As String = "", _
'                     Optional ByVal Dob As String = "") _
'                     As BIEResults
''Discipline is either "Bio", "Imm" or "End"
'
'    Dim BRs As New BIEResults
'    Dim br As BIEResult
'    Dim tb As Recordset
'    Dim tbRange As Recordset
'    Dim tbNewIDX As Recordset
'    Dim sql As String
'    Dim DaysOld As Long
'    Dim SELECTNormalRange As String
'    Dim SELECTFlagRange As String
'    Dim TableName As String
'    Dim SelectSexLow As String
'    Dim SelectSexHigh As String
'
'    On Error GoTo Load_Error
'
'    If Cat = "" Then Cat = "Human"
'
'    If Val(SampleID) = 0 Then
'        Set BRs = Nothing
'        Exit Function
'    End If
'
'    TableName = Discipline & ResultOrRepeat
'
'    If UCase$(Discipline) = "BIO" Then
'        sql = "UPDATE BioResults " & _
'              "SET SampleType = 'S' " & _
'              "WHERE SampleID = '" & SampleID & "' " & _
'              "AND COALESCE(SampleType, '') = ''"
'        Cnxn(0).Execute sql
'    End If
'
'    If Sex = "" And Dob = "" Then
'        sql = "SELECT DoB, Sex FROM Demographics WHERE " & _
'              "SampleID = '" & SampleID & "'"
'        Set tb = Cnxn(Connection).Execute(sql)
'        If Not tb.EOF Then
'            Dob = tb!Dob & ""
'            Sex = tb!Sex & ""
'        End If
'    End If
'
'    If IsDate(Dob) Then
'        Dob = Format$(Dob, "dd/mmm/yyyy")
'        DaysOld = DateDiff("d", Dob, Now)
'    End If
'
'    Select Case Left$(UCase$(Trim$(Sex)), 1)
'    Case "M": SelectSexLow = "Male": SelectSexHigh = "Male"
'    Case "F": SelectSexLow = "Female": SelectSexHigh = "Female"
'    Case Else: SelectSexLow = "Female": SelectSexHigh = "Male"
'    End Select
'
'    SELECTNormalRange = " COALESCE(" & SelectSexLow & "Low, 0) Low, " & _
'                        " COALESCE(" & SelectSexHigh & "High, 9999) High, "
'    SELECTFlagRange = " COALESCE(Flag" & SelectSexLow & "Low, 0) FlagLow, " & _
'                      " COALESCE(Flag" & SelectSexHigh & "High, 9999) FlagHigh, "
'
'    sql = "UPDATE " & TableName & " " & _
'          "SET Result = 'XXXXX' " & _
'          "WHERE SampleID = '" & SampleID & "' " & _
'          "AND ( Result LIKE 'I_' " & _
'          "      OR Result LIKE 'L_' " & _
'          "      OR Result LIKE 'H_' )"
'    Cnxn(0).Execute sql
'
'    sql = "SELECT R.DefIndex, COALESCE(X.NormalLow, 0) Low, " & _
'          "COALESCE(X.NormalHigh, 9999) High, COALESCE(X.FlagLow, 0) FlagLow, COALESCE(X.FlagHigh, 9999) FlagHigh, " & _
'          "COALESCE(X.PlausibleLow, 0) PlausibleLow, COALESCE(X.PlausibleHigh, 9999) PlausibleHigh, " & _
'          "LongName, ShortName, DoDelta, DeltaLimit, " & _
'          "DP, PrintPriority, " & _
'          "R.SampleID, R.Code, R.Result, " & _
'          "COALESCE(R.Valid, 0) AS Valid, COALESCE(R.Printed, 0) Printed, " & _
'          "R.RunTime, R.RunDate, R.Operator, R.Flags, R.Units, " & _
'          "R.SampleType, R.Analyser, R.Faxed, R.Authorised, " & _
'          "R.Comment AS Comment, " & _
'          "COALESCE(PrintRefRange, 1) PrntRefRange , D.DeltaDaysBackLimit " & _
'          "FROM " & TableName & " R JOIN " & Discipline & "TestDefinitions D ON R.Code = D.Code " & _
'          "LEFT JOIN " & Discipline & "DefIndex X ON R.DefIndex = X.DefIndex " & _
'          "WHERE " & _
'          "SampleID = '" & SampleID & "' AND D.category = '" & Cat & "' " & _
'          "AND R.Code = D.Code AND D.Inuse = 1 " & _
'          "AND AgeFromDays <= " & DaysOld & " " & _
'          "AND AgeToDays >= " & DaysOld & " "
'
'    '290   If UCase(HospName(0)) <> "CAVAN" Then
'    sql = sql & "AND R.SampleType = D.SampleType "
'    '310   End If
'    If P = gNOTPRINTED And v = gNOTVALID Then
'        sql = sql & "AND Printed = 0 AND Valid = 0 "
'    ElseIf P = gNOTPRINTED And v = gVALID Then
'        sql = sql & "AND Printed = 0 AND Valid = 1 "
'    ElseIf P = gNOTPRINTED And v = gDONTCARE Then
'        sql = sql & "AND Printed = 0 "
'    ElseIf P = gPRINTED And v = gNOTVALID Then
'        sql = sql & "AND Printed = 1 AND Valid = 0 "
'    ElseIf P = gPRINTED And v = gVALID Then
'        sql = sql & "AND Printed = 1 AND Valid = 1 "
'    ElseIf P = gPRINTED And v = gDONTCARE Then
'        sql = sql & "AND Printed = 1 "
'    ElseIf P = gDONTCARE And v = gNOTVALID Then
'        sql = sql & "AND Valid = 0 "
'    ElseIf P = gDONTCARE And v = gVALID Then
'        sql = sql & "AND Valid = 1 "
'    End If
'
'    sql = sql & "ORDER BY PrintPriority ASC"
'    Set tb = New Recordset
'    RecOpenServer Connection, tb, sql    '  RecOpenClient 0,tb, Sql
'    Do While Not tb.EOF
'        Set br = New BIEResult
'        With br
'            .SampleID = tb!SampleID & ""
'            .Code = tb!Code & ""
'            .Result = tb!Result & ""
'            .Operator = tb!Operator & ""
'            .Rundate = Format$(tb!Rundate, "dd/mm/yyyy")
'            .RunTime = Format$(tb!RunTime, "dd/mm/yyyy hh:mm:ss")
'            .Units = tb!Units & ""
'            .Printed = tb!Printed
'            .Valid = IIf(tb!Valid, True, False)
'            If Not IsNull(tb!Authorised) Then
'                .Authorised = IIf(tb!Authorised, True, False)
'            Else
'                .Authorised = False
'            End If
'            '    .Faxed = tb!Faxed
'            '    If TB!h & "" <> "" Then .h = TB!h Else .h = False
'            '    If TB!g & "" <> "" Then .g = TB!g Else .g = False
'            '    If TB!o & "" <> "" Then .o = TB!o Else .o = False
'            '    If TB!J & "" <> "" Then .J = TB!h Else .J = False
'            '    If TB!l & "" <> "" Then .l = TB!h Else .l = False
'            '    If TB!s & "" <> "" Then .s = TB!h Else .s = False
'            .Flags = tb!Flags & ""
'            .SampleType = tb!SampleType & ""
'            .Printformat = IIf(IsNull(tb!DP), 0, tb!DP)
'            .ShortName = tb!ShortName & ""
'            .LongName = tb!LongName & ""
'            If tb!DoDelta & "" <> "" Then .DoDelta = tb!DoDelta Else .DoDelta = False
'            .DeltaLimit = IIf(IsNull(tb!DeltaLimit), 9999, tb!DeltaLimit)
'            .Analyser = tb!Analyser & ""
'            .Comment = tb!Comment & ""
'            .Low = tb!Low
'            .High = tb!High
'            .FlagLow = tb!FlagLow
'            .FlagHigh = tb!FlagHigh
'            .PlausibleLow = tb!PlausibleLow
'            .PlausibleHigh = tb!PlausibleHigh
'            .PrintRefRange = tb!PrntRefRange
'            .DefIndex = tb!DefIndex
'            .DeltaDaysBackLimit = IIf(IsNull(tb!DeltaDaysBackLimit), 0, tb!DeltaDaysBackLimit)
'            If .DefIndex = 0 Then
'                If Dob <> "" And Sex <> "" Then
'                    sql = "SELECT " & _
'                          SELECTNormalRange & SELECTFlagRange & _
'                          "COALESCE(PlausibleLow, 0) PlausibleLow, " & _
'                          "COALESCE(PlausibleHigh, 9999) PlausibleHigh " & _
'                          "FROM " & Discipline & "TestDefinitions  " & _
'                          "WHERE category = '" & Cat & "' " & _
'                          "AND Code = '" & .Code & "' " & _
'                          "AND AgeFromDays <= " & DaysOld & " " & _
'                          "AND AgeToDays >= " & DaysOld & " "
'                    Set tbRange = New Recordset
'                    RecOpenServer 0, tbRange, sql
'                    If Not tbRange.EOF Then
'                        .Low = tbRange!Low
'                        .High = tbRange!High
'                        .FlagLow = tbRange!FlagLow
'                        .FlagHigh = tbRange!FlagHigh
'                        .PlausibleLow = tbRange!PlausibleLow
'                        .PlausibleHigh = tbRange!PlausibleHigh
'
'                        sql = "SELECT * FROM " & Discipline & "DefIndex " & _
'                              "WHERE NormalLow = '" & .Low & "' " & _
'                              "AND NormalHigh = '" & .High & "' " & _
'                              "AND FlagLow = '" & .FlagLow & "' " & _
'                              "AND FlagHigh = '" & .FlagHigh & "' " & _
'                              "AND PlausibleLow = '" & .PlausibleLow & "' " & _
'                              "AND PlausibleHigh = '" & .PlausibleHigh & "' "
'
'                        Set tbNewIDX = New Recordset
'                        RecOpenClient 0, tbNewIDX, sql
'                        If Not tbNewIDX.EOF Then
'                            .DefIndex = tbNewIDX!DefIndex
'                        Else
'
'                            sql = "INSERT INTO " & Discipline & "DefIndex " & _
'                                  "( NormalLow, NormalHigh, FlagLow, FlagHigh, " & _
'                                  "  PlausibleLow, PlausibleHigh, AutoValLow, AutoValHigh) " & _
'                                  "VALUES ( " & _
'                                  .Low & ", " & .High & ", " & .FlagLow & ", " & .FlagHigh & ", " & _
'                                  .PlausibleLow & ", " & .PlausibleHigh & ", 0,9999) "
'                            Cnxn(0).Execute sql
'
'                            sql = "SELECT MAX(DefIndex) NewIndex FROM " & Discipline & "DefIndex"
'                            Set tbNewIDX = New Recordset
'                            RecOpenClient 0, tbNewIDX, sql
'                            .DefIndex = tbNewIDX!NewIndex
'
'                        End If
'
'                        sql = "UPDATE " & TableName & " " & _
'                              "SET DefIndex = '" & .DefIndex & "' " & _
'                              "WHERE SampleID = '" & .SampleID & "' " & _
'                              "AND Code = '" & .Code & "'"
'                        Cnxn(0).Execute sql
'
'                    End If
'                End If
'            End If
'
'
'            BRs.Add br
'        End With
'        tb.MoveNext
'    Loop
'
'    If BRs.Count <> 0 Then
'        Set Load = BRs
'    Else
'        Set Load = Nothing
'    End If
'    Set br = Nothing
'    Set BRs = Nothing
'
'    Exit Function
'
'Load_Error:
'
'    Dim strES As String
'    Dim intEL As Integer
'
'    intEL = Erl
'    strES = Err.Description
'    LogError "BIEResults", "Load", intEL, strES, sql
'
'End Function

Public Function Load(ByVal Discipline As String, _
                     ByVal SampleID As String, _
                     ByVal ResultOrRepeat As String, _
                     ByVal v As Long, _
                     ByVal P As Long, _
                     ByVal Cn As Long, _
                     ByVal Cat As String, _
                     ByVal Rundate As String) _
                     As BIEResults
      'Discipline is either "Bio", "Imm" or "End"

      Dim BRs As New BIEResults
      Dim br As BIEResult
      Dim tb As New Recordset
      Dim sql As String
      Dim Dob As String
      Dim DaysOld As Long
      Dim SELECTNormalRange As String
      Dim SELECTFlagRange As String
      Dim TableName As String
      Dim sex As String
      Dim tbRange As Recordset
      Dim tbNewIDX  As Recordset

10    On Error GoTo Load_Error

20    DaysOld = 9125

30    If SampleID = "" Then Exit Function

40    If UCase(Discipline) = "BIO" Or UCase(Discipline) = "BGA" Then Cat = ""

50    TableName = Discipline & ResultOrRepeat

60    sql = "SELECT DoB, Sex,rundate from Demographics WHERE " & _
            "SampleID = '" & Val(SampleID) & "'"
70    Set tb = Cnxn(Cn).Execute(sql)
80    If Not tb.EOF Then
90        If IsDate(tb!Dob) Then
100           Dob = Format$(tb!Dob, "dd/mmm/yyyy")
110           DaysOld = DateDiff("d", Dob, tb!Rundate)
120           If DaysOld = 0 Then DaysOld = 1
130       End If
140       sex = Left$(UCase$(Trim$(tb!sex & "")), 1)
150       Select Case Left$(UCase$(Trim$(tb!sex & "")), 1)
          Case "M": SELECTNormalRange = " MaleLow as Low, MaleHigh as High, "
160           SELECTFlagRange = " FlagMaleLow as FlagLow, FlagMaleHigh as FlagHigh, "
170       Case "F": SELECTNormalRange = " FemaleLow as Low, FemaleHigh as High, "
180           SELECTFlagRange = " FlagFemaleLow as FlagLow, FlagFemaleHigh as FlagHigh, "
190       Case Else: SELECTNormalRange = " FemaleLow as Low, MaleHigh as High, "
200           SELECTFlagRange = " FlagFemaleLow as FlagLow, FlagMaleHigh as FlagHigh, "
210       End Select
220   Else
230       SELECTNormalRange = " FemaleLow as Low, MaleHigh as High, "
240       SELECTFlagRange = " FlagFemaleLow as FlagLow, FlagMaleHigh as FlagHigh, "
250   End If


260   sql = "SELECT COALESCE(R.DefIndex, 0) DefIndex, COALESCE(X.NormalLow, 0) Low, " & _
            "COALESCE(X.NormalHigh, 9999) High, COALESCE(X.FlagLow, 0) FlagLow, COALESCE(X.FlagHigh, 9999) FlagHigh, " & _
            "COALESCE(X.PlausibleLow, 0) PlausibleLow, COALESCE(X.PlausibleHigh, 9999) PlausibleHigh, " & _
            "LongName, ShortName, DoDelta, DeltaLimit, CheckTime, Printable, " & _
            "DP, PrintPriority, " & _
            "R.SampleID, R.Code, R.Result, " & _
            "COALESCE(R.Valid, 0) AS Valid, COALESCE(R.Printed, 0) Printed, " & _
            "R.RunTime, R.RunDate, R.Operator, R.Flags, R.Units, " & _
            "R.SampleType, R.Analyser, R.Faxed, R.Authorised, " & _
            "R.Comment AS Comment, R.PC "
270   If UCase$(Discipline) = "IMM" Then
280       sql = sql & ", prnrr "
290   End If
300   sql = sql & "FROM " & TableName & " R JOIN " & Discipline & "TestDefinitions D ON R.Code = D.Code " & _
            "LEFT JOIN " & Discipline & "DefIndex X ON R.DefIndex = X.DefIndex " & _
            "WHERE " & _
            "SampleID = '" & SampleID & "' AND D.category = '" & Cat & "' " & _
            "AND R.Code = D.Code " & _
            "AND AgeFromDays <= " & DaysOld & " " & _
            "AND AgeToDays >= " & DaysOld & " "




310   sql = sql & "and R.SampleType = D.SampleType "
320   If P = gNOTPRINTED And v = gNOTVALID Then
330       sql = sql & "and Printed = 0 and Valid = 0 "
340   ElseIf P = gNOTPRINTED And v = gVALID Then
350       sql = sql & "and Printed = 0 and Valid = 1 "
360   ElseIf P = gNOTPRINTED And v = gDONTCARE Then
370       sql = sql & "and Printed = 0 "
380   ElseIf P = gPRINTED And v = gNOTVALID Then
390       sql = sql & "and Printed = 1 and Valid = 0 "
400   ElseIf P = gPRINTED And v = gVALID Then
410       sql = sql & "and Printed = 1 and Valid = 1 "
420   ElseIf P = gPRINTED And v = gDONTCARE Then
430       sql = sql & "and Printed = 1 "
440   ElseIf P = gDONTCARE And v = gNOTVALID Then
450       sql = sql & "and Valid = 0 "
460   ElseIf P = gDONTCARE And v = gVALID Then
470       sql = sql & "and Valid = 1 "
480   End If
490   sql = sql & "Order by PrintPriority asc"
500   Set tb = New Recordset
510   RecOpenServer Cn, tb, sql    '  RecOpenClient 0,tb, Sql
520   Do While Not tb.EOF
530       Set br = New BIEResult
540       With br
550           .SampleID = Trim(tb!SampleID & "")
560           .DefIndex = tb!DefIndex
570           .Code = Trim(tb!Code & "")
580           .Result = Trim(tb!Result & "")
590           .Operator = Trim(tb!Operator & "")
600           .Rundate = Format$(tb!Rundate, "dd/mm/yyyy")
610           .RunTime = Format$(tb!RunTime, "dd/mm/yyyy hh:mm")
620           .Units = Trim(tb!Units & "")
630           If Trim(tb!Printed & "") <> "" Then .Printed = IIf(tb!Printed, True, False)
640           If Trim(tb!Valid & "") <> "" Then .Valid = IIf(tb!Valid, True, False)
650           .Flags = Trim(tb!Flags & "")
660           .SampleType = Trim(tb!SampleType & "")
670           .Low = IIf(IsNull(tb!Low), 0, tb!Low)
680           .FlagLow = IIf(IsNull(tb!FlagLow), 0, tb!FlagLow)
690           .PlausibleLow = IIf(IsNull(tb!PlausibleLow), 0, tb!PlausibleLow)
700           .High = IIf(IsNull(tb!High), 9999, tb!High)
710           .FlagHigh = IIf(IsNull(tb!FlagHigh), 9999, tb!FlagHigh)
720           .PlausibleHigh = IIf(IsNull(tb!PlausibleHigh), 99999, tb!PlausibleHigh)
730           .Printformat = IIf(IsNull(tb!DP), 0, tb!DP)
740           .ShortName = Trim(tb!ShortName & "")
750           .LongName = Trim(tb!LongName & "")
760           If tb!DoDelta & "" <> "" Then .DoDelta = tb!DoDelta Else .DoDelta = False
770           .DeltaLimit = IIf(IsNull(tb!DeltaLimit), 9999, tb!DeltaLimit)
780           .Analyser = Trim(tb!Analyser & "")
790           If Discipline = "Imm" Then
800               .PrnRR = IIf(IsNull(tb!PrnRR), True, tb!PrnRR)
810           End If
820           .Comment = Trim(tb!Comment & "")
830           .Pc = Trim(tb!Pc & "")
840           If IsNull(tb!CheckTime) Then
850               .CheckTime = 1
860           Else
870               .CheckTime = tb!CheckTime
880           End If
890           .Printable = tb!Printable
              '   .NormalLow = tb!NormalLow
              '   .NormalHigh = tb!NormalHigh
              '   .NormalUsed = tb!NormalUsed

900           If .DefIndex = 0 Then
910               If Dob <> "" And sex <> "" Then
920                   sql = "SELECT " & _
                            SELECTNormalRange & SELECTFlagRange & _
                            "COALESCE(PlausibleLow, 0) PlausibleLow, " & _
                            "COALESCE(PlausibleHigh, 9999) PlausibleHigh " & _
                            "FROM " & Discipline & "TestDefinitions  " & _
                            "WHERE category = '" & Cat & "' " & _
                            "AND Code = '" & .Code & "' " & _
                            "AND AgeFromDays <= " & DaysOld & " " & _
                            "AND AgeToDays >= " & DaysOld & " "
930                   Set tbRange = New Recordset
940                   RecOpenServer 0, tbRange, sql
950                   If Not tbRange.EOF Then
960                       .Low = tbRange!Low
970                       .High = tbRange!High
980                       .FlagLow = tbRange!FlagLow
990                       .FlagHigh = tbRange!FlagHigh
1000                      .PlausibleLow = tbRange!PlausibleLow
1010                      .PlausibleHigh = tbRange!PlausibleHigh

1020                      sql = "SELECT * FROM " & Discipline & "DefIndex " & _
                                "WHERE NormalLow = '" & .Low & "' " & _
                                "AND NormalHigh = '" & .High & "' " & _
                                "AND FlagLow = '" & .FlagLow & "' " & _
                                "AND FlagHigh = '" & .FlagHigh & "' " & _
                                "AND PlausibleLow = '" & .PlausibleLow & "' " & _
                                "AND PlausibleHigh = '" & .PlausibleHigh & "' "

1030                      Set tbNewIDX = New Recordset
1040                      RecOpenClient 0, tbNewIDX, sql
1050                      If Not tbNewIDX.EOF Then
1060                          .DefIndex = tbNewIDX!DefIndex
1070                      Else

1080                          sql = "INSERT INTO " & Discipline & "DefIndex " & _
                                    "( NormalLow, NormalHigh, FlagLow, FlagHigh, " & _
                                    "  PlausibleLow, PlausibleHigh, AutoValLow, AutoValHigh) " & _
                                    "VALUES ( " & _
                                    .Low & ", " & .High & ", " & .FlagLow & ", " & .FlagHigh & ", " & _
                                    .PlausibleLow & ", " & .PlausibleHigh & ", 0,9999) "
1090                          Cnxn(0).Execute sql

1100                          sql = "SELECT MAX(DefIndex) NewIndex FROM " & Discipline & "DefIndex"
1110                          Set tbNewIDX = New Recordset
1120                          RecOpenClient 0, tbNewIDX, sql
1130                          .DefIndex = tbNewIDX!NewIndex

1140                      End If

1150                      sql = "UPDATE " & TableName & " " & _
                                "SET DefIndex = '" & .DefIndex & "' " & _
                                "WHERE SampleID = '" & .SampleID & "' " & _
                                "AND Code = '" & .Code & "'"
1160                      Cnxn(0).Execute sql

1170                  End If
1180              End If
1190          End If
1200          BRs.Add br
1210      End With
1220      tb.MoveNext
1230  Loop

1240  If BRs.Count <> 0 Then
1250      Set Load = BRs
1260  Else
1270      Set Load = Nothing
1280  End If
1290  Set br = Nothing
1300  Set BRs = Nothing


1310  Exit Function



1320  Exit Function

Load_Error:

      Dim strES As String
      Dim intEL As Integer


1330  intEL = Erl
1340  strES = Err.Description
1350  LogError "BIEResults", "Load", intEL, strES, sql


End Function
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"

10        On Error GoTo NewEnum_Error

20        Set NewEnum = mcolBIEResults.[_NewEnum]

30        Exit Function

NewEnum_Error:

          Dim strES As String
          Dim intEL As Integer


40        intEL = Erl
50        strES = Err.Description
60        LogError "BIEResults", "NewEnum", intEL, strES


End Function



Public Sub RemoveItem(ByVal X As Long)

10        On Error GoTo RemoveItem_Error

20        mcolBIEResults.Remove X

30        Exit Sub

RemoveItem_Error:

          Dim strES As String
          Dim intEL As Integer


40        intEL = Erl
50        strES = Err.Description
60        LogError "BIEResults", "RemoveItem", intEL, strES


End Sub


Public Sub Save(ByVal Discipline As String, _
                ByVal BRs As BIEResults)
      'Discipline is either "Bio", "Imm" or "End"

          Dim tb As New Recordset
          Dim sql As String
          Dim br As BIEResult
          Dim lngSampleID As Long

10        On Error GoTo Save_Error

20        For Each br In BRs
30            With br
40                lngSampleID = Format$(Val(.SampleID))
50                sql = "SELECT * from " & Discipline & "Results WHERE " & _
                        "SampleID = '" & lngSampleID & "' " & _
                        "and Code = '" & br.Code & "'"
60                Set tb = New Recordset
70                RecOpenClient 0, tb, sql
80                If Not tb.EOF Then
90                    sql = "SELECT * from " & Discipline & "Repeats WHERE " & _
                            "SampleID = '" & lngSampleID & "'"
100                   Set tb = New Recordset
110                   RecOpenClient 0, tb, sql
120               End If
130               tb.AddNew
140               tb!SampleID = lngSampleID
150               If IsDate(.Rundate) Then
160                   tb!Rundate = Format$(.Rundate, "dd/mmm/yyyy")
170               Else
180                   tb!Rundate = Format$(.RunTime, "dd/mmm/yyyy")
190               End If
200               If Trim$(.RunTime) <> "" Then
210                   tb!RunTime = Format$(.RunTime, "dd/mmm/yyyy hh:mm")
220               End If
230               tb!Code = .Code
240               tb!Result = .Result
250               tb!Units = .Units
260               tb!Printed = .Printed
270               tb!Valid = .Valid
280               tb!Faxed = 0
290               tb!Analyser = .Analyser
300               tb!SampleType = .SampleType
310               tb.Update

320               sql = "DELETE from " & Discipline & "Requests WHERE " & _
                        "SampleID = '" & lngSampleID & "' " & _
                        "and Code = '" & br.Code & "'"
330               Cnxn(0).Execute sql

340           End With
350       Next

360       Set br = Nothing


370       Exit Sub

Save_Error:

          Dim strES As String
          Dim intEL As Integer


380       intEL = Erl
390       strES = Err.Description
400       LogError "BIEResults", "Save", intEL, strES, sql


End Sub




