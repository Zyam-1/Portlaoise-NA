Attribute VB_Name = "basLibrary"
Option Explicit

Public Xnxn As Connection   'Used in Ide to save Sql statements for analysis
Private Declare Function GetComputerName Lib "Kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Public Declare Function SetCursorPos Lib "user32.dll" (ByVal X As Long, ByVal Y As Long) As Long

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Const gYES = True
Public Const gNO = False
Public Const gNOCHANGE = 4

Public Enum InputValidation
    Numericfullstopdash = 0
    Char = 1
    YorN = 2
    AlphaNumeric_NoApos = 3
    AlphaNumeric_AllowApos = 4
    Numeric_Only = 5
    AlphaNumeric_WithApos = 6
End Enum

Public Enum SearchType
    ExactMatch = 0
    LeadingCharacters = 1
    TrailingCharacters = 2
End Enum

Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Public Function SearchForIrishNames(ByVal strName As String, Optional ByVal ST As SearchType = ExactMatch) As String

'Step 1: Is 1st letter "O"

Dim LeadingChar As String
Dim TrailingChar As String

On Error GoTo SearchForIrishNames_Error

Select Case ST
    Case SearchType.ExactMatch
        LeadingChar = ""
        TrailingChar = ""
    Case SearchType.LeadingCharacters
        LeadingChar = "%"
        TrailingChar = ""
    Case SearchType.TrailingCharacters
        LeadingChar = ""
        TrailingChar = "%"
End Select
    

If UCase(Left(Trim(strName), 1)) = "O" Then
    'Step 2: Is 2nd character Apostrophe or space
    If UCase(Mid(strName, 2, 1)) = "'" Then
        'Search with Apostrophe , Apostrophe replaced with space, Space removed from surname
        SearchForIrishNames = " (PatName LIKE '" & TrailingChar & AddTicks(strName) & LeadingChar & "' or " & _
            "PatName LIKE '" & TrailingChar & Replace(strName, "'", " ", , 1) & LeadingChar & "' or " & _
            "PatName LIKE '" & TrailingChar & Replace(strName, "'", "", , 1) & LeadingChar & "')"
    ElseIf UCase(Mid(strName, 2, 1)) = " " Then
        SearchForIrishNames = " (PatName LIKE '" & TrailingChar & AddTicks(strName) & LeadingChar & "' or PatName LIKE '" & TrailingChar & Replace(strName, " ", "''", , 1) & LeadingChar & "' or PatName LIKE '" & TrailingChar & Replace(strName, " ", "", , 1) & LeadingChar & "')"
    Else    'OBrien
        SearchForIrishNames = " (PatName LIKE '" & TrailingChar & strName & LeadingChar & "' or PatName LIKE '" & TrailingChar & "O''" & Mid(strName, 2) & LeadingChar & "' or PatName LIKE '" & TrailingChar & "O " & Mid(strName, 2) & LeadingChar & "')"
    End If
Else
    SearchForIrishNames = " PatName LIKE '" & TrailingChar & AddTicks(strName) & LeadingChar & "'"
End If

Exit Function

SearchForIrishNames_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "basLibrary", "SearchForIrishNames", intEL, strES

End Function

Public Function CheckDemographics(ByVal TrialID As String) As String

          Dim sn As New Recordset
          Dim sql As String
          Dim n As Long
          Dim pName(1 To 4) As String
          Dim pAddress(1 To 4) As String
          Dim pDoB(1 To 4) As String
          Dim IDFound(1 To 4) As Boolean
          Dim Found As Long
          Dim f As Form

10        On Error GoTo CheckDemographics_Error

20        If TrialID = "" Then Exit Function

30        Set sn = New Recordset
40        With sn
50            Found = 0
60            For n = 1 To 2
70                IDFound(n) = False
80                sql = "SELECT * from patientifs WHERE " & _
                        Choose(n, "CHART", "AandE") & " = '" & TrialID & "'"
90                RecOpenServer 0, sn, sql
100               If Not .EOF Then
110                   Do While Not sn.EOF
120                       IDFound(n) = True
130                       Found = Found + 1
140                       pName(n) = initial2upper(!PatName & "")
150                       If Not IsNull(!Dob) Then pDoB(n) = Format(!Dob, "dd/MM/yyyy")
160                       pAddress(n) = initial2upper(!Address0 & "") & " " & initial2upper(!Address1 & "")
170                       sn.MoveNext
180                   Loop
190               End If
200               .Close
210           Next
220       End With

230       If Found = 0 Then
240           CheckDemographics = ""
250       ElseIf Found = 1 Then
260           For n = 1 To 2
270               If IDFound(n) Then
280                   CheckDemographics = Choose(n, "CHART", "AandE")
290                   Exit For
300               End If
310           Next
320       Else
330           Set f = New frmDemogCheck
340           With f
350               For n = 1 To 2
360                   If IDFound(n) Then
370                       .bSelect(n).Visible = True
380                       .lName(n) = initial2upper(pName(n))
390                       .lAddress(n) = initial2upper(pAddress(n))
400                       .lDoB(n) = pDoB(n)
410                   End If
420               Next
430               .Show 1
440               CheckDemographics = .IDType
450           End With
460           Unload f
470           Set f = Nothing
480       End If

490       Exit Function

CheckDemographics_Error:

          Dim strES As String
          Dim intEL As Integer

500       intEL = Erl
510       strES = Err.Description
520       LogError "basLibrary", "CheckDemographics", intEL, strES, sql


End Function

Public Function CheckPhoneLog1(ByVal SID As String) As PhoneLog

      'Returns PhoneLog.SampleID = 0 if no entry in phone log

          Dim tb As Recordset
          Dim sql As String
          Dim PL As PhoneLog

10        On Error GoTo CheckPhoneLog1_Error

20        sql = "Select * from PhoneLog where " & _
                "SampleID = '" & Val(SID) & "'"
30        Set tb = Cnxn(0).Execute(sql)
40        If tb.EOF Then
50            CheckPhoneLog1.SampleID = 0
60        Else
70            With PL
80                .SampleID = Val(SID)
90                .Comment = tb!Comment & ""
100               .Datetime = tb!Datetime
110               .Discipline = tb!Discipline & ""
120               .PhonedBy = tb!PhonedBy & ""
130               .PhonedTo = tb!PhonedTo & ""
140               .Title = tb!Title & ""
150               .PersonName = tb!PersonName & ""

160           End With
170           CheckPhoneLog1 = PL
180       End If

190       Exit Function

CheckPhoneLog1_Error:

          Dim strES As String
          Dim intEL As Integer

200       intEL = Erl
210       strES = Err.Description
220       LogError "basLibrary", "CheckPhoneLog1", intEL, strES

End Function

Public Function CheckPhoneLog(ByVal SID As String) As Boolean

      'Returns True if an entry in phone log

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo CheckPhoneLog_Error

20        sql = "Select * from PhoneLog where " & _
                "SampleID = '" & Val(SID) & "'"
30        Set tb = Cnxn(0).Execute(sql)

40        CheckPhoneLog = Not tb.EOF

50        Exit Function

CheckPhoneLog_Error:

          Dim strES As String
          Dim intEL As Integer



60        intEL = Erl
70        strES = Err.Description
80        LogError "basLibrary", "CheckPhoneLog", intEL, strES, sql


End Function

Public Sub SetToolTip(ByVal OnOff As Boolean, ByVal Frm As Form)

          Dim ax As Control

10        On Error GoTo SetToolTip_Error

20        If OnOff = False Then
30            For Each ax In Frm.Controls
40                If TypeOf ax Is TextBox Or TypeOf ax Is Label _
                     Or TypeOf ax Is MSFlexGrid Or TypeOf ax Is Frame _
                     Or TypeOf ax Is ComboBox Or TypeOf ax Is ListBox _
                     Or TypeOf ax Is CheckBox Or TypeOf ax Is OptionButton _
                     Or TypeOf ax Is SSPanel Or TypeOf ax Is CommandButton Then

50                    ax.ToolTipText = ""
60                End If
70            Next
80        End If

90        Exit Sub

SetToolTip_Error:

          Dim strES As String
          Dim intEL As Integer



100       intEL = Erl
110       strES = Err.Description
120       LogError "basLibrary", "SetToolTip", intEL, strES

End Sub
Public Function AddTicks(ByVal s As String) As String
      'add single comma's to names

10        On Error GoTo AddTicks_Error

20        s = Trim$(s)

30        s = Replace(s, "'", "''")

40        AddTicks = s

50        Exit Function

AddTicks_Error:

          Dim strES As String
          Dim intEL As Integer



60        intEL = Erl
70        strES = Err.Description
80        LogError "basLibrary", "AddTicks", intEL, strES

End Function

Public Function BetweenDates(ByVal Index As Integer, _
                             ByRef upto As String) _
                             As String

          Dim From As String
          Dim m As Long

10        On Error GoTo BetweenDates_Error

20        Select Case Index
          Case 0:    'last week
30            From = Format$(DateAdd("ww", -1, Now), "dd/mm/yyyy")
40            upto = Format$(Now, "dd/mm/yyyy")
50        Case 1:    'last month
60            From = Format$(DateAdd("m", -1, Now), "dd/mm/yyyy")
70            upto = Format$(Now, "dd/mm/yyyy")
80        Case 2:    'last fullmonth
90            From = Format$(DateAdd("m", -1, Now), "dd/mm/yyyy")
100           From = "01/" & Mid$(From, 4)
110           upto = DateAdd("m", 1, From)
120           upto = Format$(DateAdd("d", -1, upto), "dd/mm/yyyy")
130       Case 3:    'last quarter
140           From = Format$(DateAdd("q", -1, Now), "dd/mm/yyyy")
150           upto = Format$(Now, "dd/mm/yyyy")
160       Case 4:    'last full quarter
170           From = Format$(DateAdd("q", -1, Now), "dd/mm/yyyy")
180           m = Val(Mid$(From, 4, 2))
190           m = ((m - 1) \ 3) * 3 + 1
200           From = "01/" & Format$(m, "00") & Mid$(From, 6)
210           upto = DateAdd("q", 1, From)
220           upto = Format$(DateAdd("d", -1, upto), "dd/mm/yyyy")
230       Case 5:    'year to date
240           From = "01/01/" & Format$(Now, "yyyy")
250           upto = Format$(Now, "dd/mm/yyyy")
260       Case 6:    'today
270           From = Format$(Now, "dd/mm/yyyy")
280           upto = From
290       End Select

300       BetweenDates = From

310       Exit Function

BetweenDates_Error:

          Dim strES As String
          Dim intEL As Integer



320       intEL = Erl
330       strES = Err.Description
340       LogError "basLibrary", "BetweenDates", intEL, strES

End Function
Public Function CalcpAge(ByVal Dob As String) As String

          Dim diff As Long
          Dim DobYr As Single

10        On Error GoTo CalcpAge_Error

20        Dob = Format$(Dob, "dd/mm/yyyy")
30        If IsDate(Dob) Then
40            diff = DateDiff("d", (Dob), (Now))
50            DobYr = diff / 365.25
60            If DobYr > 1 Then
70                CalcpAge = Int(DobYr)
80            ElseIf diff < 30.43 Then
90                CalcpAge = diff
100           Else
110               CalcpAge = Int(diff / 30.43)
120           End If
130       Else
140           CalcpAge = ""
150       End If

160       Exit Function

CalcpAge_Error:

          Dim strES As String
          Dim intEL As Integer



170       intEL = Erl
180       strES = Err.Description
190       LogError "basLibrary", "CalcpAge", intEL, strES

End Function
Public Function CalcAge(ByVal Dob As String, ByVal strSampleDate As String) As String

          Dim diff As Long
          Dim DobYr As Single

10        On Error GoTo CalcAge_Error

20        Dob = Format$(Dob, "dd/mm/yyyy")
30        If IsDate(Dob) Then
40            diff = DateDiff("d", (Dob), strSampleDate)
50            If diff < 0 Then
60                CalcAge = ""
70                Exit Function
80            End If

90            DobYr = diff / 365.25
100           If DobYr > 1 Then    'Year
110               CalcAge = Format$(Int(DobYr), "###\Yr")
120           ElseIf diff < 30.43 Then    'Day
130               If diff = 0 Then
140                   CalcAge = "0D"
150               Else
160                   CalcAge = Format$(diff, "##\D")
170               End If
180           Else    ' Month
190               CalcAge = Format$(Int(diff / 30.43), "##\M")
200           End If
210       Else
220           CalcAge = ""
230       End If

240       Exit Function

CalcAge_Error:

          Dim strES As String
          Dim intEL As Integer

250       intEL = Erl
260       strES = Err.Description
270       LogError "basLibrary", "CalcAge", intEL, strES

End Function

Public Function CalcOldAge(ByVal Dob As String, ByVal Old As String) As String

          Dim diff As Long
          Dim DobYr As Single

10        On Error GoTo CalcOldAge_Error

20        Dob = Format$(Dob, "dd/mm/yyyy")
30        If IsDate(Dob) Then
40            diff = DateDiff("d", (Dob), (Old))
50            DobYr = diff / 365.25
60            If DobYr > 1 Then
70                CalcOldAge = Format$(Int(DobYr), "###\Yr")
80            ElseIf diff < 30.43 Then
90                CalcOldAge = Format$(diff, "##\D")
100           Else
110               CalcOldAge = Format$(Int(diff / 30.43), "##\M")
120           End If
130       Else
140           CalcOldAge = ""
150       End If

160       Exit Function

CalcOldAge_Error:

          Dim strES As String
          Dim intEL As Integer



170       intEL = Erl
180       strES = Err.Description
190       LogError "basLibrary", "CalcOldAge", intEL, strES

End Function


Public Sub ClearFGrid(ByVal g As MSFlexGrid)

10        On Error GoTo ClearFGrid_Error

20        With g
30            .Rows = .FixedRows + 1
40            .AddItem ""
50            .RemoveItem .FixedRows
60            .Visible = False
70        End With

80        Exit Sub

ClearFGrid_Error:

          Dim strES As String
          Dim intEL As Integer



90        intEL = Erl
100       strES = Err.Description
110       LogError "basLibrary", "ClearFGrid", intEL, strES

End Sub
Public Function Convert62Date(ByVal s As String, _
                              ByVal Direction As Long) _
                              As String

          Dim D As String

10        On Error GoTo Convert62Date_Error

20        If Len(s) <> 6 Then
30            Convert62Date = s
40            Exit Function
50        End If

60        D = Left$(s, 2) & "/" & Mid$(s, 3, 2) & "/" & Right$(s, 2)
70        If IsDate(D) Then
80            Select Case Direction
              Case BACKWARD:
90                If DateValue(D) > DateValue(Now) Then
100                   D = DateAdd("yyyy", -100, D)
110               End If
120               Convert62Date = Format$(D, "dd/mm/yyyy")
130           Case FORWARD:
140               If DateValue(D) < Now Then
150                   D = DateAdd("yyyy", 100, D)
160               End If
170               Convert62Date = Format$(D, "dd/mm/yyyy")
180           Case gDONTCARE:
190               Convert62Date = Format$(D, "dd/mm/yyyy")
200           End Select
210       Else
220           Convert62Date = s
230       End If




240       Exit Function

Convert62Date_Error:

          Dim strES As String
          Dim intEL As Integer



250       intEL = Erl
260       strES = Err.Description
270       LogError "basLibrary", "Convert62Date", intEL, strES


End Function

Public Function CountDays(ByVal Number As Long, ByVal Interval As String) As Long

10        On Error GoTo CountDays_Error

20        Select Case Interval
          Case "Days": CountDays = Number
30        Case "Months": CountDays = Number * (365.25 / 12)
40        Case "Years": CountDays = Number * 365.25
50        End Select

60        Exit Function

CountDays_Error:

          Dim strES As String
          Dim intEL As Integer

70        intEL = Erl
80        strES = Err.Description
90        LogError "basLibrary", "CountDays", intEL, strES

End Function


Function date2code(ByVal s As String) As String

10        On Error GoTo date2code_Error

20        If Not IsDate(s) Then
30            date2code = ""
40        Else
50            s = Format$(s, "dd/mm/yyyy")
60            date2code = Right$(s, 4) & Mid$(s, 4, 2) & Left$(s, 2)
70        End If

80        Exit Function

date2code_Error:

          Dim strES As String
          Dim intEL As Integer



90        intEL = Erl
100       strES = Err.Description
110       LogError "basLibrary", "date2code", intEL, strES

End Function



Public Function dmyFromCount(ByVal Days As Long) As String

          Dim D As Long
          Dim m As Long
          Dim Y As Long
          Dim s As String

10        On Error GoTo dmyFromCount_Error

20        Y = Int(Days / 365)

30        Days = Days - Fix(Y * 365.25)

40        m = Days \ 30.42

50        D = Days - (m * 30.42)

60        If Y > 0 Then
70            s = Format$(Y) & "Y "
80        End If

90        If m > 0 Then
100           s = s & Format$(m) & "M "
110       End If

120       dmyFromCount = s & Format$(D, "0") & "D"

130       Exit Function

dmyFromCount_Error:

          Dim strES As String
          Dim intEL As Integer



140       intEL = Erl
150       strES = Err.Description
160       LogError "basLibrary", "dmyFromCount", intEL, strES

End Function

Public Sub FixGridColWidth(ByVal g As MSFlexGrid, ByVal Frm As Form)

      Dim intCol     As Integer
      Dim intRow     As Integer

10    On Error GoTo FixGridColWidth_Error

20    For intCol = 0 To g.Cols - 1
30        For intRow = 0 To g.Rows - 1    'makes the colum with to fit text
40            If g.ColWidth(intCol) > 0 Then
50                If g.ColWidth(intCol) < Frm.TextWidth(g.TextMatrix(intRow, intCol)) + 100 Then
60                    g.ColWidth(intCol) = Frm.TextWidth(g.TextMatrix(intRow, intCol)) + 100
70                End If
80            End If
90        Next
100   Next

110   Exit Sub
FixGridColWidth_Error:

120   LogError "basLibrary", "FixGridColWidth", Erl, Err.Description


End Sub

Public Sub FixG(ByVal g As MSFlexGrid)
      Dim intRow As Integer
      Dim intCol As Integer
10    On Error GoTo FixG_Error

20    With g
30        .Visible = True
40        If .Rows > .FixedRows + 1 And .TextMatrix(.FixedRows, 0) = "" Then
50            .RemoveItem .FixedRows
60        End If
      '    For intcol = 0 To .Cols - 1
      '        For introw = 0 To .Rows - 1 'makes the colum with to fit text
      '            If .ColWidth(intcol) < frmMicroSurveillanceSearches.TextWidth(.TextMatrix(introw, intcol)) + 100 Then
      '                .ColWidth(intcol) = frmMicroSurveillanceSearches.TextWidth(.TextMatrix(introw, intcol)) + 100
      '            End If
      '            If intcol > 1 And .TextMatrix(0, intcol) = "" Then 'deletes colum if there is no heading
      '                .ColWidth(intcol) = 0
      '            End If
      '        Next
      '    Next
70    End With

80    Exit Sub

FixG_Error:

      Dim strES As String
      Dim intEL As Integer

90    intEL = Erl
100   strES = Err.Description
110   LogError "basLibrary", "FixG", intEL, strES

End Sub

Public Sub FlashNoPrevious(ByVal lbl As Label)

          Dim T As Single
          Dim n As Long

10        On Error GoTo FlashNoPrevious_Error

20        With lbl
30            For n = 1 To 5
40                .Visible = True
50                .Refresh
60                T = Timer
70                Do While Timer - T < 0.1: DoEvents: Loop
80                .Visible = False
90                .Refresh
100               T = Timer
110               Do While Timer - T < 0.1: DoEvents: Loop
120           Next
130       End With

140       Exit Sub

FlashNoPrevious_Error:

          Dim strES As String
          Dim intEL As Integer

150       intEL = Erl
160       strES = Err.Description
170       LogError "basLibrary", "FlashNoPrevious", intEL, strES

End Sub

Function ForeName(ByVal s As String) As String

          Dim P As Long

10        On Error GoTo ForeName_Error

20        s = Trim$(s)
30        If s = "" Then
40            ForeName = ""
50        Else
60            P = InStr(s, " ")
70            If P = 0 Then
80                ForeName = ""
90            Else
100               ForeName = Trim$(Mid$(s, P))
110           End If
120       End If

130       Exit Function

ForeName_Error:

          Dim strES As String
          Dim intEL As Integer

140       intEL = Erl
150       strES = Err.Description
160       LogError "basLibrary", "ForeName", intEL, strES

End Function

Public Sub GetAgeText(ByVal AgeInDays As Long, _
                      ByRef DMYNumber As Long, _
                      ByRef DMYInterval As String)

          Dim temp As Long

10        On Error GoTo GetAgeText_Error

20        temp = AgeInDays / 365.25
30        If temp >= 1 Then
40            DMYInterval = "Years"
50            DMYNumber = temp
60        Else
70            temp = AgeInDays / (365.25 / 12)
80            If temp >= 1 Then
90                DMYInterval = "Months"
100               DMYNumber = temp
110           Else
120               DMYInterval = "Days"
130               DMYNumber = AgeInDays
140           End If
150       End If

160       Exit Sub

GetAgeText_Error:

          Dim strES As String
          Dim intEL As Integer



170       intEL = Erl
180       strES = Err.Description
190       LogError "basLibrary", "GetAgeText", intEL, strES

End Sub

Function initial2upper(ByVal s As String) As String

          Dim n As Long

10        On Error GoTo initial2upper_Error

20        s = Trim$(s & "")
30        If s = "" Then
40            initial2upper = ""
50            Exit Function
60        End If

70        If InStr(UCase$(s), "MAC") > 0 Or InStr(UCase$(s), "MC") > 0 Or InStr(s, "'") > 0 Then
80            s = LCase$(s)
90            s = UCase$(Left$(s, 1)) & Mid(s, 2)

100           For n = 1 To Len(s) - 1
110               If Mid(s, n, 1) = " " Or Mid(s, n, 1) = "'" Or Mid(s, n, 1) = "." Then
120                   s = Left$(s, n) & UCase$(Mid(s, n + 1, 1)) & Mid(s, n + 2)
130               End If
140               If n > 1 Then
150                   If Mid(s, n, 1) = "c" And Mid(s, n - 1, 1) = "M" Then
160                       s = Left$(s, n) & UCase$(Mid(s, n + 1, 1)) & Mid(s, n + 2)
170                   End If
180               End If
190           Next
200       Else
210           s = StrConv(s, vbProperCase)
220       End If
230       initial2upper = s

240       Exit Function

initial2upper_Error:

          Dim strES As String
          Dim intEL As Integer



250       intEL = Erl
260       strES = Err.Description
270       LogError "basLibrary", "initial2upper", intEL, strES

End Function

Public Function IsRoutine() As Boolean

      'Returns True if time now is between
      '09:30 and 17:30 Mon to Fri
      'else returns False

10        On Error GoTo IsRoutine_Error

20        IsRoutine = False

30        If Weekday(Now) <> vbSaturday And Weekday(Now) <> vbSunday Then
40            If TimeValue(Now) > TimeValue("09:29") And _
                 TimeValue(Now) < TimeValue("17:31") Then
50                IsRoutine = True
60            End If
70        End If

80        Exit Function

IsRoutine_Error:

          Dim strES As String
          Dim intEL As Integer



90        intEL = Erl
100       strES = Err.Description
110       LogError "basLibrary", "IsRoutine", intEL, strES

End Function

Function LowAge(ByVal Age As String) As Long

10        On Error GoTo LowAge_Error

20        LowAge = False
30        If Val(Age) <> 0 Then
40            If InStr(Age, "D") Then
50                LowAge = True
60            ElseIf InStr(Age, "M") Then
70                LowAge = True
80            ElseIf Val(Age) < 15 Then
90                LowAge = True
100           End If
110       End If

120       Exit Function

LowAge_Error:

          Dim strES As String
          Dim intEL As Integer



130       intEL = Erl
140       strES = Err.Description
150       LogError "basLibrary", "LowAge", intEL, strES

End Function

Public Function ParseForeName(ByVal Name As String) As String

          Dim n As Long
          Dim temp As String

10        On Error GoTo ParseForeName_Error

20        Name = Trim$(UCase$(Name))

30        If InStr(Name, "B/O") Or _
             InStr(Name, "BABY") Then
40            ParseForeName = ""
50            Exit Function
60        End If

70        temp = Trim(Mid$(Name, 2))

80        n = InStr(temp, " ")
90        If n = 0 Then
100           ParseForeName = ""
110           Exit Function
120       End If

130       temp = Mid$(temp, n + 1)

          Rem Code Change 16/01/2006
          'checks if a double barreled name
140       n = InStr(temp, " ")
150       temp = Mid$(temp, n + 1)
160       If Trim(temp) = "" Then
170           Exit Function
180       End If

190       If InStr(temp, " ") Or _
             temp Like "*[!A-Z]*" Or _
             Len(temp) = 1 Then
200           ParseForeName = ""
210       Else
220           ParseForeName = temp
230       End If

240       Exit Function

ParseForeName_Error:

          Dim strES As String
          Dim intEL As Integer



250       intEL = Erl
260       strES = Err.Description
270       LogError "basLibrary", "ParseForeName", intEL, strES

End Function

Public Sub RecClose(ByVal rs As Recordset)

10        On Error GoTo RecClose_Error

20        rs.Close
30        Set rs = Nothing

40        Exit Sub

RecClose_Error:

          Dim strES As String
          Dim intEL As Integer



50        intEL = Erl
60        strES = Err.Description
70        LogError "basLibrary", "RecClose", intEL, strES

End Sub

Public Sub RecOpenClient(ByVal n As Long, ByVal RecSet As Recordset, ByVal sql As String)

          Dim T As Single
          Dim X As Single

10        T = Timer

20        With RecSet
30            .CursorLocation = adUseClient
40            .CursorType = adOpenDynamic
50            .LockType = adLockOptimistic
60            .ActiveConnection = Cnxn(n)
70            .Source = sql
80            .Open
90        End With

100       X = Timer - T
110       If X > 0.1 Then
120           Debug.Print X; " "; sql
130       End If

End Sub

Public Sub RecOpenClientBB(ByVal RecSet As Recordset, ByVal sql As String)

10        With RecSet
20            .CursorLocation = adUseClient
30            .CursorType = adOpenDynamic
40            .LockType = adLockOptimistic
50            .ActiveConnection = CnxnBB
60            .Source = sql
70            .Open
80        End With

End Sub

Public Sub RecOpenRemoteClient(ByVal RecSet As Recordset, ByVal sql As String)

10        With RecSet
20            .CursorLocation = adUseClient
30            .CursorType = adOpenDynamic
40            .LockType = adLockOptimistic
50            .ActiveConnection = CnxnRemote
60            .Source = sql
70            .Open
80        End With

End Sub

Public Sub RecOpenRemoteServer(ByVal RecSet As Recordset, ByVal sql As String)

10        With RecSet
20            .CursorLocation = adUseServer
30            .CursorType = adOpenDynamic
40            .LockType = adLockOptimistic
50            .ActiveConnection = CnxnRemote
60            .Source = sql
70            .Open
80        End With

End Sub

Public Sub RecOpenServer(ByVal n As Long, ByVal RecSet As Recordset, ByVal sql As String)

          Dim T As Single
          Dim X As Single

10        T = Timer
20        With RecSet
30            .CursorLocation = adUseServer
40            .CursorType = adOpenDynamic
50            .LockType = adLockOptimistic
60            .ActiveConnection = Cnxn(n)
70            .Source = sql
80            .Open
90        End With

100       X = Timer - T
110       If X > 0.1 Then
120           Debug.Print X; " "; sql
130       End If

End Sub

Public Sub RecOpenServerBB(ByVal Cx As Long, _
                           ByVal RecSet As Recordset, _
                           ByVal sql As String)

10        With RecSet
20            .CursorLocation = adUseServer
30            .CursorType = adOpenDynamic
40            .LockType = adLockOptimistic
50            .ActiveConnection = CnxnBB(Cx)
60            .Source = sql
70            .Open
80        End With

End Sub

Function SurName(ByVal s As String) As String

          Dim P As Long

10        On Error GoTo SurName_Error

20        s = Trim$(s)
30        If s = "" Then
40            SurName = ""
50        Else
60            P = InStr(s, " ")
70            If P = 0 Then
80                SurName = s
90            Else
100               SurName = Left$(s, P - 1)
110           End If
120       End If

130       Exit Function

SurName_Error:

          Dim strES As String
          Dim intEL As Integer



140       intEL = Erl
150       strES = Err.Description
160       LogError "basLibrary", "SurName", intEL, strES

End Function

Function time2code(ByVal s As String) As String

      'convert dd/mm/yyyy hh:mm:ss
      'to 14 character string "yyyymmddhhmmss"

          Dim Code As String

10        On Error GoTo time2code_Error

20        If Not IsDate(s) Then
30            time2code = ""
40            Exit Function
50        End If

60        s = Format$(s, "dd/mm/yyyy hh:mm:ss")

70        Code = Mid$(s, 7, 4) & Mid$(s, 4, 2)
80        Code = Code & Left$(s, 2) & Mid$(s, 12, 2)
90        Code = Code & Mid$(s, 15, 2) & Mid$(s, 18, 2)

100       time2code = Code

110       Exit Function

time2code_Error:

          Dim strES As String
          Dim intEL As Integer



120       intEL = Erl
130       strES = Err.Description
140       LogError "basLibrary", "time2code", intEL, strES

End Function

Public Function Split_Comm(ByVal Comm As String) As String

          Dim n As Long
          Dim s As String
          Dim Cnt As Long

10        On Error GoTo Split_Comm_Error

20        For n = 1 To Len(Comm)
30            If Asc(Mid(Comm, n, 1)) = Asc(vbCr) Or Asc(Mid(Comm, n, 1)) = 10 Then
40                If Cnt = 0 Then
50                    If n > 1 Then s = s & vbCrLf
60                    Cnt = 1
70                End If
80            Else
90                s = s & Mid(Comm, n, 1)
100               Cnt = 0
110           End If
120       Next

130       Split_Comm = s

140       Exit Function

Split_Comm_Error:

          Dim strES As String
          Dim intEL As Integer



150       intEL = Erl
160       strES = Err.Description
170       LogError "basLibrary", "Split_Comm", intEL, strES

End Function



Public Function vbGetComputerName() As String

      'Gets the name of the machine
          Const MAXSIZE As Integer = 256
          Dim sTmp As String * MAXSIZE
          Dim lLen As Long

10        On Error GoTo vbGetComputerName_Error

20        lLen = MAXSIZE - 1
30        If (GetComputerName(sTmp, lLen)) Then
40            vbGetComputerName = Left$(sTmp, lLen)
50        Else
60            vbGetComputerName = ""
70        End If

80        Exit Function

vbGetComputerName_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "basLibrary", "vbGetComputerName", intEL, strES


End Function


Public Function VI(KeyAscii As Integer, _
                   iv As InputValidation, _
                   Optional NextFieldOnEnter As Boolean) As Integer

      Dim sTemp As String

10    sTemp = chr$(KeyAscii)
20    If KeyAscii = 13 Then    'Enter Key
30        If NextFieldOnEnter = True Then
40            VI = 9    'Return Tab Keyascii if User Selected NextFieldOnEnter Option
50        Else
60            VI = 13
70        End If
80        Exit Function
90    ElseIf KeyAscii = 8 Then    'BackSpace
100       VI = 8
110       Exit Function
120   End If

      ' turn input to upper case

130   Select Case iv
      Case 0:    'NumbersFullstopDash
140       Select Case sTemp
          Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", ".", "-", "<", ">"
150           VI = Asc(sTemp)
160       Case Else
170           VI = 0
180       End Select

190   Case 1:    'Characters Only
200       Select Case sTemp
          Case " ", "-"
210           VI = Asc(sTemp)
220       Case "A" To "Z"
230           VI = Asc(sTemp)
240       Case "a" To "z"
250           VI = Asc(sTemp)
260       Case Else
270           VI = 0
280       End Select

290   Case 2:    'Y or N Only
300       sTemp = UCase(chr$(KeyAscii))
310       Select Case sTemp
          Case "Y", "N"
320           VI = Asc(sTemp)
330       Case Else
340           VI = 0
350       End Select

360   Case 3:    'AlphaNumeric Only...No Apostrophe
370       Select Case sTemp
          Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", ".", _
               " ", "/", ";", ":", "\", "-", ">", "<", "(", ")", "@", _
               "%", "!", """", "+", "^", "~", "`", "Ç", "´", "Ã", "Á", _
               "Â", "È", "É", "Ê", "Ì", "Í", "Î", "Ò", "Ó", "Ô", "Õ", _
               "Ù", "Ú", "Û", "Ü", "à", "á", "â", "ã", "ç", "è", "é", _
               "ê", "ì", "í", "î", "ò", "ó", "ô", "õ", "ö", "ù", "ú", _
               "û", "ü", "Æ", "æ", ",", "?", "=", "*", "#"
380           VI = Asc(sTemp)
390       Case "A" To "Z"
400           VI = Asc(sTemp)
410       Case "a" To "z"
420           VI = Asc(sTemp)
430       Case Else
440           VI = 0
450       End Select

460   Case 4:    'AlphaNumeric Only...With Apostophe
470       Select Case sTemp
          Case "1", "2", "3", "4", "5", "6", "7", "8", "9", ".", " ", "'"
480           VI = Asc(sTemp)
490       Case "A" To "Z"
500           VI = Asc(sTemp)
510       Case "a" To "z"
520           VI = Asc(sTemp)
530       Case Else
540           VI = 0
550       End Select

560   Case 5:    'Numbers Only
570       Select Case sTemp
          Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0"
580           VI = Asc(sTemp)
590       Case Else
600           VI = 0
610       End Select

620   Case 6:    'AlphaNumeric Only...With Apostrophe
630       Select Case sTemp
          Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", ".", _
               " ", "/", ";", ":", "\", "-", ">", "<", "(", ")", "@", _
               "%", "!", """", "+", "^", "~", "`", "Ç", "´", "Ã", "Á", _
               "Â", "È", "É", "Ê", "Ì", "Í", "Î", "Ò", "Ó", "Ô", "Õ", _
               "Ù", "Ú", "Û", "Ü", "à", "á", "â", "ã", "ç", "è", "é", _
               "ê", "ì", "í", "î", "ò", "ó", "ô", "õ", "ö", "ù", "ú", _
               "û", "ü", "Æ", "æ", ",", "?", "=", "*", "#", "'"
640           VI = Asc(sTemp)
650       Case "A" To "Z"
660           VI = Asc(sTemp)
670       Case "a" To "z"
680           VI = Asc(sTemp)
690       Case Else
700           VI = 0
710       End Select

720   End Select

730   If VI = 0 Then Beep

End Function

Public Function EntriesOK(ByVal SampleID As String, _
                          ByVal SurName As String, _
                          ByVal sex As String, _
                          ByVal Ward As String, _
                          ByVal GP As String) _
                          As Boolean

10        On Error GoTo EntriesOK_Error

20        EntriesOK = False

30        If Trim$(SampleID) = "" Then
40            iMsg "Must have Lab Number.", vbCritical
50            Exit Function
60        End If

70        If Trim$(sex) = "" Then
80            If iMsg("Sex not entered." & vbCrLf & "Do you want to enter sex now?", vbQuestion + vbYesNo) = vbYes Then
90                Exit Function
100           End If
110       End If

120       If Trim$(SurName) = "" Then
130           iMsg "Name not entered.", vbCritical
140           Exit Function
150       Else
160           If Trim$(Ward) = "" Then
170               iMsg "Must have Ward entry.", vbCritical
180               Exit Function
190           End If

200           If Trim$(Ward) = "GP" Then
210               If Trim$(GP) = "" Then
220                   iMsg "Must have Ward or GP entry.", vbCritical
230                   Exit Function
240               End If
250           End If
260       End If

270       EntriesOK = True

280       Exit Function

EntriesOK_Error:

          Dim strES As String
          Dim intEL As Integer

290       intEL = Erl
300       strES = Err.Description
310       LogError "basLibrary", "EntriesOK", intEL, strES


End Function
