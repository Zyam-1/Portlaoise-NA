Attribute VB_Name = "basHistology"
Option Explicit
Public gPaperSize As String

Public Enum HistoCyto
    Histology = 0
    Cytology = 1
End Enum

Public Function Swap_Year(ByVal Hyear As String) As String

10        On Error GoTo Swap_Year_Error

20        Swap_Year = Right(Hyear, 1) & Mid(Hyear, 3, 1) & Mid(Hyear, 2, 1) & Left(Hyear, 1)

30        Exit Function

Swap_Year_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "basHistology", "Swap_Year", intEL, strES


End Function

Sub PrintResultHisto(ByVal FullRunNumber As String)

          Dim tb As New Recordset
          Dim sql As String
          Dim n As Long
10        ReDim PL(1 To 1) As String
          Dim plCounter As Long
          Dim crPos As Long
          Dim HR As String
          Dim LinesAllowed As Long
          Dim TotalPages As Long
          Dim ThisPage As Long
          Dim TopLine As Long
          Dim BottomLine As Long
          Dim crlfFound As Boolean

20        On Error GoTo PrintResultHisto_Error

30        If gPaperSize = "A5" Then
40            LinesAllowed = 23
50        Else
60            LinesAllowed = 56
70        End If

80        sql = "SELECT * from historesults, demographics WHERE " & _
                "demographics.sampleid = '" & FullRunNumber & "' " & _
                "and demographics.sampleid = historesults.sampleid"
90        Set tb = New Recordset
100       RecOpenServer 0, tb, sql
110       If tb.EOF Then Exit Sub

120       HR = Trim(tb!historeport & "")
130       crlfFound = True
140       Do While crlfFound
150           HR = RTrim(HR)
160           crlfFound = False
170           If Right(HR, 1) = vbCr Or Right(HR, 1) = vbLf Then
180               HR = Left(HR, Len(HR) - 1)
190               crlfFound = True
200           End If
210       Loop

220       plCounter = 0
230       Do While Len(HR) > 0
240           crPos = InStr(HR, vbCr)
250           If crPos > 0 And crPos < 91 Then
260               plCounter = plCounter + 1
270               ReDim Preserve PL(1 To plCounter)
280               PL(plCounter) = Left(HR, crPos - 1)
290               HR = Mid(HR, crPos + 2)
300           Else
310               If Len(HR) > 91 Then
320                   For n = 91 To 1 Step -1
330                       If Mid(HR, n, 1) = " " Then
340                           Exit For
350                       End If
360                   Next
370                   plCounter = plCounter + 1
380                   ReDim Preserve PL(1 To plCounter)
390                   PL(plCounter) = Left(HR, n)
400                   HR = Mid(HR, n + 1)
410               Else
420                   plCounter = plCounter + 1
430                   ReDim Preserve PL(1 To plCounter)
440                   PL(plCounter) = HR
450                   Exit Do
460               End If
470           End If
480       Loop

490       TotalPages = Int((plCounter - 1) / LinesAllowed) + 1
500       If TotalPages = 0 Then TotalPages = 1
510       For ThisPage = 1 To TotalPages
520           PrintHistoHeading tb, FullRunNumber, ThisPage, TotalPages
530           TopLine = (ThisPage - 1) * LinesAllowed + 1
540           BottomLine = (ThisPage - 1) * LinesAllowed + LinesAllowed
550           If BottomLine > plCounter Then
560               BottomLine = plCounter
570           End If
580           For n = TopLine To BottomLine
590               If UCase(Left(PL(n), 24)) = "MICROSCOPIC EXAMINATION:" Or _
                     UCase(Left(PL(n), 29)) = "BONE MARROW ASPIRATE & BIOPSY" Or _
                     UCase(Left(PL(n), 18)) = "GROSS EXAMINATION:" Or _
                     UCase(Left(PL(n), 21)) = "SUPPLEMENTARY REPORT:" Or _
                     UCase(Left(PL(n), 14)) = "DR. K. CUNNANE" Or _
                     UCase(Left(PL(n), 17)) = "DR. KEVIN CUNNANE" Or _
                     UCase(Left(PL(n), 11)) = "DR GERARD C" Or _
                     UCase(Left(PL(n), 11)) = "PATHOLOGIST" Or _
                     UCase(Left(PL(n), 18)) = "DR. J. D. GILSENAN" Or _
                     UCase(Left(PL(n), 14)) = "FURTHER REPORT" Or _
                     UCase(Left(PL(n), 10)) = "APPEARANCE" Or _
                     UCase(Left(PL(n), 23)) = "MICROSCOPIC EXAMINATION" Or _
                     UCase(Left(PL(n), 10)) = "CONSULTANT" Or _
                     UCase(Left(PL(n), 7)) = "COMMENT" Or _
                     UCase(Left(PL(n), 21)) = "SUPPLEMENTARY REPORT" Then
600                   Printer.Font.Bold = True
610               Else
620                   Printer.Font.Bold = False
630               End If
640               Printer.Print Tab(3); PL(n)
650           Next
660           Printer.EndDoc
670       Next

680       Exit Sub

PrintResultHisto_Error:

          Dim strES As String
          Dim intEL As Integer

690       intEL = Erl
700       strES = Err.Description
710       LogError "basHistology", "PrintResultHisto", intEL, strES

End Sub

Public Sub PrintHistoHeading(ByVal tb As Recordset, _
                             ByVal FullRunNumber As String, _
                             ByVal ThisPage As Long, _
                             ByVal TotalPages As Long)

          Dim Hosp As String

          'Printer.ColorMode = vbPRCMColor
10        On Error GoTo PrintHistoHeading_Error

20        Printer.ScaleMode = vbTwips

30        If TotalPages = 1 Then
40            Printer.Print
50        Else
60            Printer.Font.Bold = True
70            Printer.Print "Page " & Format(ThisPage) & " of " & Format(TotalPages)
80        End If
90        Printer.FontName = "Courier New"
100       Printer.FontSize = 14
110       Printer.ForeColor = QBColor(4)

120       Printer.Font.Bold = True
130       Printer.Print Tab(17); HospName(0) & " Histopathology  Laboratory"
140       Printer.Font.Bold = False

150       Printer.FontName = "Courier New"
160       Printer.FontSize = 10
170       Printer.ForeColor = QBColor(0)

180       Printer.Font.Bold = False
190       Printer.Print Tab(72); "Lab #: ";
200       Printer.Font.Bold = True
210       Printer.Print FullRunNumber

220       Printer.Font.Bold = False
230       Printer.Print Tab(3); " Name: ";
240       Printer.Font.Bold = True
250       Printer.Print tb!PatName & "";
260       Printer.Font.Bold = False
270       Printer.Print Tab(73); "Cons: ";
280       Printer.Font.Bold = True
290       Printer.Print tb!Clinician & ""

300       Printer.Font.Bold = False
310       Printer.Print Tab(3); " Addr: ";
320       Printer.Font.Bold = True
330       Printer.Print tb!Addr0 & " " & tb!Addr1 & "";
340       Printer.Font.Bold = False
350       Printer.Print Tab(73); "Hosp: ";
360       Printer.Font.Bold = True
370       Printer.Print Hosp

380       Printer.Font.Bold = False
390       Printer.Print Tab(3); "  DoB: ";
400       Printer.Font.Bold = True
410       If Not IsNull(tb!Dob) Then
420           If IsDate(tb!Dob) Then Printer.Print Format(tb!Dob, "dd/mm/yyyy");
430       End If
440       Printer.Font.Bold = False
450       Printer.Print Tab(23); "Sex: ";
460       Printer.Font.Bold = True
470       Select Case Left(tb!sex & " ", 1)
          Case "M": Printer.Print "Male";
480       Case "F": Printer.Print "Female";
490       End Select
500       Printer.Font.Bold = False
510       Printer.Print Tab(75); "GP: ";
520       Printer.Font.Bold = True
530       Printer.Print tb!GP & ""

540       Printer.Font.Bold = False
550       Printer.Print Tab(3); "Chart: ";
560       Printer.Font.Bold = True
570       Printer.Print tb!Chart & "";
580       Printer.Font.Bold = False
590       Printer.Print Tab(22); "Ward: ";
600       Printer.Font.Bold = True
610       Printer.Print tb!Ward & "";
620       Printer.Font.Bold = False
630       Printer.Print Tab(68); "Date Recd: ";
640       Printer.Font.Bold = True
650       Printer.Print tb!Rundate & " " & tb!TimeTaken

660       Printer.Font.Bold = False
670       Printer.Print Tab(3); "Nature of Specimen [A]: ";
680       Printer.Font.Bold = True
690       Printer.Print tb!NatureOfSpecimen & "";
700       If tb!natureofspecimenB & "" <> "" Then
710           Printer.Font.Bold = False
720           Printer.Print Tab(55); "[B]: ";
730           Printer.Font.Bold = True
740           Printer.Print tb!natureofspecimenB & ""
750       Else
760           Printer.Print
770       End If

780       If tb!natureofspecimenC & "" <> "" Then
790           Printer.Font.Bold = False
800           Printer.Print Tab(22); "[C]: ";
810           Printer.Font.Bold = True
820           Printer.Print tb!natureofspecimenC & "";
830           Printer.Font.Bold = False
840           Printer.Print Tab(55); "[D]: ";
850           Printer.Font.Bold = True
860           Printer.Print tb!natureofspecimenD & ""
870       Else
880           Printer.Print
890       End If

900       Printer.Print Tab(3); String$(90, "-")

910       Exit Sub

PrintHistoHeading_Error:

          Dim strES As String
          Dim intEL As Integer



920       intEL = Erl
930       strES = Err.Description
940       LogError "basHistology", "PrintHistoHeading", intEL, strES

End Sub

Public Function AreHistoResultsPresent(ByVal SampleID As String, ByVal Year As String) As Long

          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo AreHistoResultsPresent_Error

20        AreHistoResultsPresent = 0

30        If Val(SampleID) <> 0 And Trim$(Year) <> "" Then

40            SampleID = SampleID + SysOptHistoOffset(0) + (Val(Swap_Year(Year)) * 1000)

50            sql = "SELECT count(*) as tot from Historesults WHERE " & _
                    "SampleID = '" & SampleID & "' and hyear = '" & Year & "'"
60            Set tb = New Recordset
70            Set tb = Cnxn(0).Execute(sql)

80            AreHistoResultsPresent = Sgn(tb!Tot)

90        End If

100       Exit Function

AreHistoResultsPresent_Error:

          Dim strES As String
          Dim intEL As Integer

110       intEL = Erl
120       strES = Err.Description
130       LogError "basHistology", "AreHistoResultsPresent", intEL, strES

End Function


Public Function IsHistoValid(SampleID As String, HC As HistoCyto) As Boolean

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo IsHistoValid_Error

20        sql = "Select COALESCE(@HistoCyto,0) As IsValid From Demographics Where SampleID = '@SampleID'"
30        sql = Replace(sql, "@HistoCyto", IIf(HC = Histology, "HistoValiD", "CytoValid"))
40        sql = Replace(sql, "@SampleID", SampleID)

50        Set tb = New Recordset
60        RecOpenClient 0, tb, sql
70        If tb.EOF Then
80            IsHistoValid = False
90        Else
100           IsHistoValid = tb!IsValid
110       End If
120       Exit Function

IsHistoValid_Error:

          Dim strES As String
          Dim intEL As Integer

130       intEL = Erl
140       strES = Err.Description
150       LogError "basHistology", "IsHistoValid", intEL, strES, sql

End Function
