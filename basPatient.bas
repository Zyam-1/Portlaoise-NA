Attribute VB_Name = "basPatient"
Option Explicit

Public Sub LoadPatientFromAandE(ByVal f As Form, ByVal NewRecord As Boolean)

          Dim tbPatIF As Recordset
          Dim tbDemog As Recordset
          Dim sql As String
          Dim RooH As Boolean
          Dim x As Long

10        On Error GoTo LoadPatientFromAandE_Error

20        If Trim(f.txtAandE) = "" Then Exit Sub

30        f.bViewBB.Enabled = False

40        sql = "SELECT * from PatientIFs WHERE " & _
                "AandE = '" & f.txtAandE & "'"
50        Set tbPatIF = New Recordset
60        RecOpenServer 0, tbPatIF, sql

70        sql = "SELECT * from demographics WHERE " & _
                "AandE = '" & f.txtAandE & "' " & _
                "order by rundate desc"
80        Set tbDemog = New Recordset
90        RecOpenServer 0, tbDemog, sql

100       If tbPatIF.EOF And tbDemog.EOF Then
110           f.txtName = ""
120           f.taddress(0) = ""
130           f.taddress(1) = ""
140           f.txtSex = ""
150           f.txtDoB = ""
160           f.txtAge = ""
170           If SysOptDemo(0) = True Then
180               f.cmbWard = "GP"
190               f.cmbClinician = ""
200               f.cmbGP = ""
210           End If
220           f.txtDemographicComment = ""
230           f.tSampleTime.Mask = ""
240           f.tSampleTime.Text = ""
250           f.tSampleTime.Mask = "##:##"
260       ElseIf tbDemog.EOF Then
270           With tbPatIF
280               f.txtChart = !Chart & ""
290               f.txtAandE = !AandE & ""
300               f.txtName = initial2upper(!PatName & "")
310               f.txtSex = !sex
320               f.txtDoB = !Dob
330               f.txtAge = CalcAge(f.txtDoB, f.dtSampleDate)
340               If SysOptDemo(0) = True Then
350                   f.cmbWard = !Ward & ""
360                   If Trim$(f.cmbWard) = "" Then
370                       f.cmbWard = "Accident & Emergency"
380                   End If
390                   If ClinName(!Clinician & "") <> "" Then
400                       f.cmbClinician = initial2upper(ClinName(!Clinician & ""))
410                   Else
420                       f.cmbClinician = initial2upper(!Clinician)
430                   End If
440               End If
450               f.taddress(0) = initial2upper(!Address0 & "")
460               f.taddress(1) = initial2upper(!Address1 & "")
470           End With
480       ElseIf tbPatIF.EOF Then
490           If NewRecord Then
500               RooH = IsRoutine()
510               f.cRooH(0) = RooH
520               f.cRooH(1) = Not RooH
530           Else
540               If tbDemog!SampleID = f.txtSampleID Then
550                   f.cRooH(0) = tbDemog!RooH
560                   f.cRooH(1) = Not tbDemog!RooH
570               End If
580           End If
590           f.txtName = initial2upper(tbDemog!PatName & "")
600           f.taddress(0) = initial2upper(tbDemog!Addr0 & "")
610           f.taddress(1) = initial2upper(tbDemog!Addr1 & "")
620           Select Case tbDemog!sex & ""
              Case "M": f.txtSex = "Male"
630           Case "F": f.txtSex = "Female"
640           End Select
650           f.txtChart = tbDemog!Chart & ""
660           f.txtAandE = tbDemog!AandE & ""
670           f.txtAge = tbDemog!Age & ""
680           f.txtDoB = Format$(tbDemog!Dob, "dd/mm/yyyy")
690           If IsDate(f.txtDoB) Then
700               f.txtAge = CalcAge(f.txtDoB, f.dtSampleDate)
710           End If
720           If SysOptDemo(0) = True Then
730               f.cmbWard = initial2upper(tbDemog!Ward & "")
740               If Trim$(tbDemog!Ward & "") = "" Then
750                   f.cmbWard = "Accident & Emergency"
760               Else
770                   f.cmbWard = initial2upper(tbDemog!Ward & "")
780               End If
790               f.cmbClinician = initial2upper(tbDemog!Clinician & "")
800               f.cmbGP = initial2upper(tbDemog!GP & "")
810           End If
820       Else
830           If tbDemog!DateTimeDemographics & "" <> "" And tbPatIF!DateTimeAmended & "" <> "" Then x = DateDiff("h", tbDemog!DateTimeDemographics, tbPatIF!DateTimeAmended)
840           If x < 0 Or IsNull(x) Then
850               If NewRecord Then
860                   RooH = IsRoutine()
870                   f.cRooH(0) = RooH
880                   f.cRooH(1) = Not RooH
890               Else
900                   If tbDemog!SampleID = f.txtSampleID Then
910                       f.cRooH(0) = tbDemog!RooH
920                       f.cRooH(1) = Not tbDemog!RooH
930                   End If
940               End If
950               f.txtName = initial2upper(tbDemog!PatName & "")
960               If f.Name = "frmEditAll" Then
970                   f.taddress(0) = initial2upper(tbDemog!Addr0 & "")
980                   f.taddress(1) = initial2upper(tbDemog!Addr1 & "")
990               Else
1000                  f.txtAddress(0) = initial2upper(tbDemog!Addr0 & "")
1010                  f.txtAddress(1) = initial2upper(tbDemog!Addr1 & "")
1020              End If
1030              Select Case tbDemog!sex & ""
                  Case "M": f.txtSex = "Male"
1040              Case "F": f.txtSex = "Female"
1050              End Select
1060              f.txtChart = tbDemog!Chart & ""
1070              f.txtAandE = tbDemog!AandE & ""
1080              f.txtAge = tbDemog!Age & ""
1090              f.txtDoB = Format$(tbDemog!Dob, "dd/mm/yyyy")
1100              If IsDate(f.txtDoB) Then
1110                  f.txtAge = CalcAge(f.txtDoB, f.dtSampleDate)
1120              End If
1130              If SysOptDemo(0) = True Then
1140                  If Trim$(tbDemog!Ward & "") = "" Then
1150                      f.cmbWard = "Accident & Emergency"
1160                  Else
1170                      f.cmbWard = initial2upper(tbDemog!Ward & "")
1180                  End If
1190                  f.cmbClinician = initial2upper(tbDemog!Clinician & "")
1200                  f.cmbGP = initial2upper(tbDemog!GP & "")
1210              End If
1220          Else
1230              With tbPatIF
1240                  f.txtChart = !Chart & ""
1250                  f.txtAandE = !AandE & ""
1260                  f.txtName = initial2upper(!PatName & "")
1270                  f.txtSex = !sex
1280                  f.txtDoB = !Dob
1290                  f.txtAge = CalcAge(f.txtDoB, f.dtSampleDate)
1300                  f.cmbWard = !Ward & ""
1310                  If SysOptDemo(0) = True Then
1320                      If Trim$(f.cmbWard) = "" Then
1330                          f.cmbWard = "Accident & Emergency"
1340                      End If
1350                      If ClinName(!Clinician & "") <> "" Then
1360                          f.cmbClinician = initial2upper(ClinName(!Clinician & ""))
1370                      Else
1380                          f.cmbClinician = initial2upper(!Clinician)
1390                      End If
1400                  End If
1410                  f.taddress(0) = initial2upper(!Address0 & "")
1420                  f.taddress(1) = initial2upper(!Address1 & "")
1430              End With
1440          End If
1450      End If

1460      If SysOptBloodBank(0) Then
1470          If Trim$(f.txtChart) <> "" Then
1480              sql = "SELECT * from PatientDetails WHERE " & _
                        "PatNum = '" & f.txtChart & "' " & _
                        "and Name = '" & AddTicks(f.txtName) & "'"
1490              Set tbDemog = New Recordset
1500              RecOpenClientBB tbDemog, sql
1510              f.bViewBB.Enabled = Not tbDemog.EOF
1520          End If
1530      End If

1540      Exit Sub

LoadPatientFromAandE_Error:

          Dim strES As String
          Dim intEL As Integer



1550      intEL = Erl
1560      strES = Err.Description
1570      LogError "basPatient", "LoadPatientFromAandE", intEL, strES, sql

End Sub

Public Sub LoadPatientFromChart(ByVal f As Form, ByVal NewRecord As Boolean)

      Dim tb As Recordset
      Dim tbPatIF As Recordset
      Dim tbDemog As Recordset
      Dim sql As String
      Dim RooH As Boolean
      Dim strPatientEntity As String
      Dim x As Long
      Dim CurrentHospital As String
      Dim strName As String
      Dim strSex As String
      Dim EnableiPMSChart As Boolean

10    On Error GoTo LoadPatientFromChart_Error

20    f.bViewBB.Enabled = False
30    EnableiPMSChart = GetOptionSetting("EnableiPMSChart", "0")

40    If EnableiPMSChart = True Then
50        sql = "SELECT Chart, PatName, Sex, DoB, Clinician, Ward, '' GP, Address0 Addr0, Address1 Addr1 FROM PatientIFs WHERE Chart = @chart "
60    Else
70        sql = "IF EXISTS (SELECT * FROM PatientIFs WHERE Chart = @chart) AND EXISTS (SELECT * FROM Demographics WHERE Chart = @chart) " & _
                "  BEGIN " & _
                "    IF DATEDIFF(minute,(SELECT DateTimeAmended FROM patientifs WHERE Chart = @chart),(SELECT TOP 1 COALESCE(RecordDateTime, '1/1/1900') FROM Demographics WHERE Chart = @chart ORDER BY RecordDateTime desc)) < 0 " & _
                "      SELECT Chart, PatName, Sex, DoB, Clinician, Ward, '' GP, Address0 Addr0, Address1 Addr1 FROM PatientIFs WHERE Chart = @chart " & _
                "    ELSE " & _
                "      SELECT TOP 1 Chart, PatName, Sex, DoB, Clinician, Ward, GP, Addr0 , Addr1 FROM Demographics WHERE Chart = @chart ORDER BY RunDate desc " & _
                "    END " & _
                "ELSE " & _
                "  IF EXISTS (SELECT * FROM PatientIFs WHERE Chart = @chart) " & _
                "    BEGIN " & _
                "      SELECT Chart, PatName, Sex, DoB, Clinician, Ward, '' GP, Address0 Addr0, Address1 Addr1 FROM PatientIFs WHERE Chart = @chart " & _
                "    END " & _
                "  ELSE " & _
                "    BEGIN " & _
                "      IF EXISTS (SELECT * FROM Demographics WHERE Chart = @chart) " & _
                "        SELECT TOP 1 Chart, PatName, Sex, DoB, Clinician, Ward, GP, Addr0 , Addr1 FROM Demographics WHERE Chart = @chart ORDER BY RunDate desc " & _
                "      ELSE " & _
                "        SELECT '' Chart, '' PatName, '' Sex, '' DoB, '' Clinician, '' Ward, '' GP, '' Addr0 , '' Addr1 " & _
                "      END"
80    End If
90    sql = Replace(sql, "@chart", "'" & f.txtChart & "'")

100   Set tb = New Recordset
110   Set tb = Cnxn(0).Execute(sql)

120   With tb
130       If Not .EOF Then

              '410       f.txtChart = !Chart & ""
              '420       f.txtAandE = !AandE & ""
140           f.txtName = initial2upper(!PatName & "")
150           Select Case Left$(UCase(!sex & " "), 1)
              Case "M": f.txtSex = "Male"
160           Case "F": f.txtSex = "Female"
170           End Select
180           strName = f.txtName
190           strSex = f.txtSex
200           NameLostFocus strName, strSex
210           If Left$(strSex, 1) <> Left$(!sex, 1) Then
220               f.txtSex = ""
230           End If
240           If Trim(!Dob & "") > "" Then f.txtDoB = !Dob
250           If SysOptDemo(0) = True Then
260               f.txtAge = CalcAge(f.txtDoB, f.dtSampleDate)
270               If Not EnableiPMSChart Then
280                   f.cmbWard = initial2upper(!Ward & "")
290                   f.cmbClinician = !Clinician & ""
300                   f.cmbGP = !GP & ""
310               End If
320           End If
330           f.taddress(0) = initial2upper(!Addr0 & "")
340           f.taddress(1) = initial2upper(!Addr1 & "")
350       End If
360   End With

'301          If Not IsNumeric(Left(f.txtChart, 1)) Then
'401            sql = "SELECT * from lists WHERE " & _
'                     "ListType = 'HO' " & _
'                     "and code = '" & Left(f.txtChart, 1) & "'"
'501            Set tbDemog = New Recordset
'601            RecOpenServer 0, tbDemog, sql
'701            If Not tbDemog.EOF Then
'801              CurrentHospital = tbDemog!Text
'901            End If
'1010         Else
'1101           strPatientEntity = ""
'1201           CurrentHospital = HospName(0)
'1301         End If
'
'1401         sql = "SELECT * from PatientIFs WHERE " & _
'                   "Chart = '" & f.txtChart & "' "
'1501         If strPatientEntity <> "" Then
'1601           sql = sql & "and Entity = '" & strPatientEntity & "'"
'1701         End If
'1801         Set tbPatIF = New Recordset
'1901         RecOpenServer 0, tbPatIF, sql
'
'2001         sql = "SELECT top 1 * from demographics WHERE " & _
'                   "Chart = '" & f.txtChart & "' " & _
'                   "and Hospital = '" & CurrentHospital & "' " & _
'                   "order by DateTimeDemographics desc"
'2101         Set tbDemog = New Recordset
'2201         RecOpenServer 0, tbDemog, sql
'
'2301         If tbPatIF.EOF And tbDemog.EOF Then
'2401           f.txtName = ""
'2501           f.taddress(0) = ""
'2601           f.taddress(1) = ""
'2701           f.txtSex = ""
'2801           f.txtDoB = ""
'2901           f.txtAge = ""
'3001           If SysOptDemo(0) = True Then
'3101             f.cmbWard = "GP"
'3201             f.cmbClinician = ""
'3301             f.cmbGP = ""
'3401           End If
'3501           f.txtDemographicComment = ""
'3601           f.tSampleTime.Mask = ""
'3701           f.tSampleTime.Text = ""
'3801           f.tSampleTime.Mask = "##:##"
'3901         ElseIf tbDemog.EOF Then
'4001           With tbPatIF
'4101             f.txtChart = !Chart & ""
'4201             f.txtAandE = !AandE & ""
'4301             f.txtName = initial2upper(!PatName & "")
'4401             Select Case UCase(!sex & "")
'                  Case "M": f.txtSex = "Male"
'4501               Case "F": f.txtSex = "Female"
'4601             End Select
'4701             strName = f.txtName
'4801             strSex = f.txtSex
'4901             NameLostFocus strName, strSex
'5001             If Left(strSex, 1) <> Left(!sex, 1) Then
'5101               f.txtSex = ""
'5201             End If
'5301             If Trim(!Dob & "") > "" Then f.txtDoB = !Dob
'5401             If SysOptDemo(0) = True Then
'5501               f.txtAge = CalcAge(f.txtDoB, f.dtSampleDate)
'5601               f.cmbWard = initial2upper(!Ward & "")
'5701               If ClinName(!Clinician & "") <> "" Then
'5801                 f.cmbClinician = initial2upper(ClinName(!Clinician & ""))
'5901               Else
'6001                 f.cmbClinician = initial2upper(!Clinician)
'6101               End If
'6201             End If
'6301             f.taddress(0) = initial2upper(!Address0 & "")
'6401             f.taddress(1) = initial2upper(!Address1 & "")
'6501           End With
'6601         ElseIf tbPatIF.EOF Then
'6701           If NewRecord Then
'6801             RooH = IsRoutine()
'6901             f.cRooH(0) = RooH
'7001             f.cRooH(1) = Not RooH
'7101           Else
'7201             If tbDemog!SampleID = f.txtSampleID Then
'7301               f.cRooH(0) = tbDemog!RooH
'7401               f.cRooH(1) = Not tbDemog!RooH
'7501             End If
'7601           End If
'7701           f.txtName = initial2upper(tbDemog!PatName & "")
'7801           f.taddress(0) = initial2upper(tbDemog!Addr0 & "")
'7901           f.taddress(1) = initial2upper(tbDemog!Addr1 & "")
'8001           Select Case UCase(tbDemog!sex & "")
'                Case "M": f.txtSex = "Male"
'8101             Case "F": f.txtSex = "Female"
'8201           End Select
'8301           strName = f.txtName
'8401           strSex = f.txtSex
'8501           NameLostFocus strName, strSex
'8601           If Left(strSex, 1) <> Left(tbDemog!sex, 1) Then
'8701             f.txtSex = ""
'8801           End If
'8901           f.txtChart = tbDemog!Chart & ""
'9001           f.txtAandE = tbDemog!AandE & ""
'9101           f.txtAge = tbDemog!Age & ""
'9201           f.txtDoB = Format$(tbDemog!Dob, "dd/mm/yyyy")
'9301           If IsDate(f.txtDoB) Then
'9401             f.txtAge = CalcAge(f.txtDoB, f.dtSampleDate)
'9501           End If
'9601           If SysOptDemo(0) = True Then
'9701             f.cmbWard = initial2upper(tbDemog!Ward & "")
'9801             f.cmbClinician = initial2upper(tbDemog!Clinician & "")
'9901             f.cmbGP = initial2upper(tbDemog!GP & "")
'1000          End If
'1011        Else
'1021          If IsNull(tbDemog!DateTimeDemographics) Or IsNull(tbPatIF!DateTimeAmended) Then
'1031            x = 0
'1041          Else
'1051            x = DateDiff("h", tbDemog!DateTimeDemographics, tbPatIF!DateTimeAmended)
'1061          End If
'1071          If x < 1 Then
'1081            If NewRecord Then
'1091              RooH = IsRoutine()
'1102              f.cRooH(0) = RooH
'1111              f.cRooH(1) = Not RooH
'1121            Else
'1131              If tbDemog!SampleID = f.txtSampleID Then
'1141                f.cRooH(0) = tbDemog!RooH
'1151                f.cRooH(1) = Not tbDemog!RooH
'1161              End If
'1171            End If
'1181            f.txtName = initial2upper(tbDemog!PatName & "")
'1191            f.taddress(0) = initial2upper(tbDemog!Addr0 & "")
'1202            f.taddress(1) = initial2upper(tbDemog!Addr1 & "")
'1211            Select Case UCase(tbDemog!sex & "")
'                  Case "M": f.txtSex = "Male"
'1221              Case "F": f.txtSex = "Female"
'1231            End Select
'1241            strName = f.txtName
'1251            strSex = f.txtSex
'1261            NameLostFocus strName, strSex
'1271            If Left(strSex, 1) <> Left(tbDemog!sex, 1) Then
'1281              f.txtSex = ""
'1291            End If
'1302            f.txtChart = tbDemog!Chart & ""
'1311            f.txtAge = tbDemog!Age & ""
'1321            f.txtDoB = Format$(tbDemog!Dob, "dd/mm/yyyy")
'1331            If IsDate(f.txtDoB) Then
'1341              f.txtAge = CalcAge(f.txtDoB, f.dtSampleDate)
'1351            End If
'1361            If SysOptDemo(0) = True Then
'1371              f.cmbWard = initial2upper(tbDemog!Ward & "")
'1381              f.cmbClinician = initial2upper(tbDemog!Clinician & "")
'1391              f.cmbGP = initial2upper(tbDemog!GP & "")
'1402            End If
'1411          Else
'1421            With tbPatIF
'1431              f.txtChart = !Chart & ""
'1441              f.txtAandE = !AandE & ""
'1451              f.txtName = initial2upper(!PatName & "")
'1461              f.txtSex = !sex
'1471              f.txtDoB = !Dob
'1481              f.txtAge = CalcAge(f.txtDoB, f.dtSampleDate)
'1491              If SysOptDemo(0) = True Then
'1502                  f.cmbWard = initial2upper(!Ward & "")
'1511                  If ClinName(!Clinician & "") <> "" Then
'1521                    f.cmbClinician = initial2upper(ClinName(!Clinician & ""))
'1531                  Else
'1541                    f.cmbClinician = initial2upper(!Clinician)
'1551                  End If
'1561              End If
'1571              f.taddress(0) = initial2upper(!Address0 & "")
'1581              f.taddress(1) = initial2upper(!Address1 & "")
'1591            End With
'1602          End If
'1610        End If
'      '
'370   If SysOptBloodBank(0) Then
'380       If Trim$(f.txtChart) <> "" Then
'390           sql = "SELECT * from PatientDetails WHERE " & _
'                    "PatNum = '" & f.txtChart & "' " & _
'                    "and Name = '" & AddTicks(f.txtName) & "'"
'400           Set tbDemog = New Recordset
'410           RecOpenClientBB tbDemog, sql
'420           f.bViewBB.Enabled = Not tbDemog.EOF
'430       End If
'440   End If

450   Exit Sub

LoadPatientFromChart_Error:

      Dim strES As String
      Dim intEL As Integer

460   intEL = Erl
470   strES = Err.Description
480   LogError "basPatient", "LoadPatientFromChart", intEL, strES, sql

End Sub

