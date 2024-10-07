Attribute VB_Name = "basOptions"
Option Explicit

Public SysOptMicroScreen() As Long
Public SysOptImmCodeForProt() As String
Public SysOptImmCodeForAlb() As String
Public SysOptImmCodeForPara1() As String
Public SysOptImmCodeForPara2() As String
Public SysOptImmCodeForPara3() As String
Public SysOptImmCodeForIGG() As String
Public SysOptImmCodeForIGA() As String
Public SysOptImmCodeForIGM() As String
Public SysOptImmCodeForB2M() As String
Public SysOptImmCodeForUPara() As String
Public SysOptUCVal() As Long
Public SysOptPhone() As Boolean
Public SysOptRTFView() As Boolean
Public SysOptMicroAna() As Boolean
Public SysOptShowIQ200() As Boolean

Public SysOptImmCodeForProtP() As String
Public SysOptImmCodeForAlbP() As String
Public SysOptImmCodeForPara1P() As String
Public SysOptImmCodeForPara2P() As String
Public SysOptImmCodeForPara3P() As String
Public SysOptImmCodeForIGGP() As String
Public SysOptImmCodeForIGAP() As String
Public SysOptImmCodeForIGMP() As String
Public SysOptImmCodeForUParaP() As String
Public sysOptAllowCopyDemographics() As Boolean

Public Type udtOptionList
    Description As String
    Value As String
    DefinedAs As String    'Boolean/String/Single/Long/Integer etc
    UserName As String
    OptionType As String
    OptionCategory As String
End Type

Public SysOptBioMask() As Boolean

Public SysOptHistView() As Boolean

Public SysOptBioST() As String
Public SysOptBioValFore() As Boolean    'move foreward after bio validation

Public SysOptAllowDemoPrint() As Boolean    'Allows demographics to be printed
Public SysOptWardDate() As Long    'Minus Rundate Value
Public SysOptFaxCom() As Boolean    'Composite Fax
'departments
Public SysOptDeptHaem() As Boolean    'Haematology In Use
Public SysOptDeptBio() As Boolean  'Biochemistry in Use
Public SysOptDeptCoag() As Boolean    'Coagulation In Use
Public SysOptDeptMicro() As Boolean    'Microbiology in use
Public SysOptDeptImm() As Boolean  'Immunology in Use
Public SysOptDeptEnd() As Boolean  'Endocriology in use
Public SysOptDeptBga() As Boolean  'Blood Gas in Use
Public SysOptDeptExt() As Boolean  'Externals in Use
Public SysOptDeptSemen() As Boolean    'Semen anaysis Un Use
Public SysOptDeptCyto() As Boolean  'Cytology In Use
Public SysOptDeptHisto() As Boolean    'Histology in Use
Public SysOptDeptMedibridge() As Boolean  'Medibroge in Use
Public SysOptDbConn(0) As String   'Databas Connection
Public SysOptPrintAll() As String  'password for print and validation screen
Public SysOptOptPass() As String    'Optioins Password
Public SysOptHospital() As Boolean  'Hospital Name
Public SysOptBadRes() As Boolean  'option for Bad Res
Public SysOptUrgent() As Boolean    'Urgent Request Look Up
Public SysOptViewTrans() As Boolean    'Allow BB View
Public SysOptClearHaem() As Boolean  'Allows Clear Haem Results to Run
Public SysOptNoSeeRF() As Boolean  ' No See RF

'added by myles
Public SysOptDemoVal() As Boolean    'validate demographics
Public SysOptCommVal() As Boolean    'Lock comments on validation
Public SysOptView() As Boolean
Public SysOptExp() As Boolean
Public SysOptHaemN1() As String    'haematology analyser name 1
Public SysOptHaemN2() As String    'haematology analyser name 2
Public SysOptBioN1() As String    'biochemistry analyser name 1
Public SysOptBioN2() As String    'biochemistry analyser name 2
Public SysOptChange() As Boolean
Public SysOptDemo() As Boolean  'Allow full Demographic Copy
Public SysSetFoc() As String  'Set Focus Event in Feditall
Public SysOptNil10() As String    'Set Urines Nil or 10

'Bio Codes
Public SysOptBioCodeForGent() As String    'gent
Public SysOptBioCodeForPGent() As String    'pgent
Public SysOptBioCodeForTGent() As String    'tgent
Public SysOptBioCodeForCreat() As String    'creatinine
Public SysOptBioCodeForUCreat() As String    'urinay creat
Public SysOptBioCodeForUProt() As String    ' urinary prot
Public SysOptBioCodeForAlb() As String    'albumin
Public SysOptBioCodeForGlob() As String    'glob
Public SysOptBioCodeForTProt() As String    'total protein
Public SysOptBioCodeFor24UProt() As String  '24 urinary protein
Public SysOptBioCodeFor24Vol() As String    '24 Volumne
Public SysOptBioCodeForGlucose() As String  'BioGlucose Code
Public SysOptBioCodeForGlucose1() As String  'BioGlucose 1Hr Code
Public SysOptBioCodeForGlucose2() As String  'BioGlucose 2 Hr Code
Public SysOptBioCodeForGlucose3() As String  'BioGlucose 3 Hr Code
Public SysOptBioCodeForFastGlucose() As String  'Bio fast Glucose Code
Public SysOptBioCodeForChol() As String  'Cholestrol Code
Public SysOptBioCodeForHDL() As String   'HDL Code
Public SysOptBioCodeForTrig() As String  'Triglyceride code
Public SysOptBioCodeForCholHDLRatio() As String  'CholHdl Ratio
Public SysOptBioCodeForHbA1c() As String  'HBA1c Code
Public SysOptCheckCholHDLRatio() As Boolean
Public SysOptBioCodeForCreatClear() As String  'Creatine Clear Code
Public SysOptBioCodeForBad() As String    'Bad Result
Public SysOptBioCodeForGlucoseP() As String  'BioGlucose Code Plasma
Public SysOptBioCodeForGlucose1P() As String  'BioGlucose 1Hr Code Plasma
Public SysOptBioCodeForGlucose2P() As String  'BioGlucose 2 Hr Code Plasma
Public SysOptBioCodeForGlucose3P() As String  'BioGlucose 3 Hr Code Plasma
Public SysOptBioCodeForFastGlucoseP() As String  'Bio fast Glucose Code Plasma
Public SysOptBioCodeForCholP() As String  'Cholestrol Code Plasma
Public SysOptBioCodeForTrigP() As String  'Triglyceride code Plasma

'Urines
Public SysOptBioCodeForUNa() As String    'Sodium
Public SysOptBioCodeForUUrea() As String  'Urea
Public SysOptBioCodeForUK() As String    'Pota
Public SysOptBioCodeForUChol() As String    'Chol
Public SysOptBioCodeForUCA() As String    'Calcium
Public SysOptBioCodeForUPhos() As String    'Phos
Public SysOptBioCodeForUMag() As String    'Mag

Public SysOptBioCodeFor24UCreat() As String    'urinay creat
Public SysOptBioCodeFor24UNa() As String    'Sodium
Public SysOptBioCodeFor24UUrea() As String  'Urea
Public SysOptBioCodeFor24UK() As String    'Pota
Public SysOptBioCodeFor24UChol() As String    'Chol
Public SysOptBioCodeFor24UCA() As String    'Calcium
Public SysOptBioCodeFor24UPhos() As String    'Phos
Public SysOptBioCodeFor24UMag() As String    'Mag

Public SysOptEBad() As String    'Extern Bad TestNumber
Public SysOptCBad() As String    'Coag Bad TestNumber

Public SysOptHivCode() As String

Public SysOptWBCDC() As Boolean  '

Public SysOptFullFaeces() As Boolean

Public SysOptUrgentRef() As Double  'Refresh rate of Urgent
Public SysOptNumLen() As Double  'Sample Id Length
Public SysOptNoCumShow() As Boolean    'Show or Not Cumulative

'phone Numbers
Public SysOptHaemPhone() As String  'haematology
Public SysOptBioPhone() As String   'Biochemistry
Public SysOptCoagPhone() As String  'Coagulation
Public SysOptBloodPhone() As String    'blood Trans
Public SysOptImmPhone() As String   'Immunology
Public SysOptEndPhone() As String   'Endocrinology
Public SysOptRealImm() As Boolean   'New Imm Print

'Analyser names
Public SysOptHaemAn1() As String
Public SysOptHaemAn2() As String

'offsets
Public SysOptSemenOffset() As Double    '100,000,000,000 was 10,000,000
Public SysOptMicroOffset() As Double    '200,000,000,000 was 20,000,000
Public SysOptHistoOffset() As Long    '30,000,000
Public SysOptCytoOffset() As Long    '40,000,000

Public SysOptPNE() As Boolean  'password change
Public SysOptViewHFlag() As Boolean

Public SysOptDipStick() As Boolean
Public SysOptMicroSpecific() As Boolean
Public SysOptUseFullID() As Boolean
Public SysOptDefaultABs() As Long

Public SysOptShortFaeces() As Boolean

Public SysOptBloodBank() As Boolean  'Allow Blood Bank Look up
Public SysOptRemote() As Boolean  'Remote Stations

Public SysOptDisablePractices() As Boolean
Public SysOptDisableWardOrdering() As Boolean   'Allow Ward Ordering

Public SysOptDontShowPrevCoag() As Boolean  'show prev coag on Edit screen
Public SysOptAllowWardFreeText() As Boolean
Public SysOptAllowClinicianFreeText() As Boolean
Public SysOptAllowGPFreeText() As Boolean

Public SysOptAlwaysRequestFBC() As Boolean
Public SysOptBioSamp() As Boolean

Public SysOptHistoSamps() As Long    'Number of samples 0 - 5

Public SysOptDoAssGlucose() As Boolean
Public SysOptAlphaOrderTechnicians() As Boolean    'Technecions order by name/number
Public SysOptLongOrShortBioNames() As String  'Show long/short name
Public SysOptSampleTime() As Boolean

'user options
Public SysOptToolTip() As Boolean
Public SysOptAutoRef() As Boolean    'Turn on/Off auto refresh
Public SysOptDefaultTab() As String  'Default Tab
Public SysOptFont() As String    'Font Type
Public SysOptLowFore() As String  'Text Color of Low
Public SysOptLowBack() As String
Public SysOptPlasFore() As String
Public SysOptPlasBack() As String
Public SysOptHighFore() As String
Public SysOptHighBack() As String
Public SysOptGpClin() As Boolean    'Allow UPDATE of Gp/Clin/Ward
Public SysOptExtDefault() As Boolean

Public SysOptMaxSampleUrineBatchPrinting() As Integer

Public Function GetOptionSetting(ByVal Description As String, _
                                 ByVal Default As String) As String

          Dim tb As Recordset
          Dim sql As String
          Dim RetVal As String

10        On Error GoTo GetOptionSetting_Error

20        sql = "SELECT Contents FROM Options WHERE " & _
                "Description = '" & Description & "'"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If tb.EOF Then
60            RetVal = Default
70        ElseIf Trim$(tb!Contents & "") = "" Then
80            RetVal = Default
90        Else
100           RetVal = tb!Contents
110       End If

120       GetOptionSetting = RetVal

130       Exit Function

GetOptionSetting_Error:

          Dim strES As String
          Dim intEL As Integer

140       intEL = Erl
150       strES = Err.Description
160       LogError "basOptions", "GetOptionSetting", intEL, strES, sql

End Function

Public Sub SaveOptionSetting(ByVal Description As String, _
                             ByVal Contents As String)

          Dim sql As String

10        On Error GoTo SaveOptionSetting_Error

20        sql = "IF EXISTS (SELECT * FROM Options WHERE " & _
                "           Description = '" & Description & "') " & _
                "  UPDATE Options " & _
                "  SET Contents = '" & Contents & "' " & _
                "  WHERE Description = '" & Description & "' " & _
                "ELSE " & _
                "  INSERT INTO Options " & _
                "  (Description, Contents) VALUES " & _
                "  ('" & Description & "', " & _
                "   '" & Contents & "')"
30        Cnxn(0).Execute sql

40        Exit Sub

SaveOptionSetting_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "basOptions", "SaveOptionSetting", intEL, strES, sql

End Sub

Public Sub SaveOptionSettingEx(ByVal Description As String, _
                               ByVal Contents As String, _
                               Optional Details As String = "", _
                               Optional Category As String = "", _
                               Optional OptionType As String = "", _
                               Optional ListOrder As Integer = 0)



          Dim sql As String

10        On Error GoTo SaveOptionSetting_Error

          'Created on 27/03/2012 16:40:31
          'Autogenerated by SQL Scripting

20        sql = "If Exists(Select 1 From Options " & _
                "Where Description = '@Description0' ) " & _
                "Begin " & _
                "Update Options Set " & _
                "Contents = '@Contents1', " & _
                "Username = '@Username2', " & _
                "Listorder = @Listorder3, " & _
                "OptType = '@OptType4', " & _
                "Details = '@Details7', " & _
                "optCategory = '@optCategory8' " & _
                "Where Description = '@Description0'  " & _
                "End  " & _
                "Else " & _
                "Begin  " & _
                "Insert Into Options (Description, Contents, Username, Listorder, OptType, Details, optCategory) Values " & _
                "('@Description0', '@Contents1', '@Username2', @Listorder3, '@OptType4', '@Details7', '@optCategory8') " & _
                "End"

30        sql = Replace(sql, "@Description0", Description)
40        sql = Replace(sql, "@Contents1", Contents)
50        sql = Replace(sql, "@Username2", UserName)
60        sql = Replace(sql, "@Listorder3", ListOrder)
70        sql = Replace(sql, "@OptType4", OptionType)
80        sql = Replace(sql, "@Details7", Details)
90        sql = Replace(sql, "@optCategory8", Category)

100       Cnxn(0).Execute sql

110       Exit Sub

SaveOptionSetting_Error:

          Dim strES As String
          Dim intEL As Integer

120       intEL = Erl
130       strES = Err.Description
140       LogError "basOptions", "SaveOptionSetting", intEL, strES, sql

End Sub

Public Sub LoadOptions()

          Dim tb As New Recordset
          Dim sql As String
          Dim n As Long

10        On Error GoTo LoadOptions_Error

20        ReDimOptions

30        For n = 0 To intOtherHospitalsInGroup
40            sysOptAllowCopyDemographics(n) = IIf(GetOptionSetting("ALLOWCOPYDEMOGRAPHICS", "1") = "1", True, False)
50            SysOptMicroAna(n) = IIf(GetOptionSetting("MicroANA", "1") = "1", True, False)
60            SysOptBioMask(n) = IIf(GetOptionSetting("BioMask", "1") = "1", True, False)
70            SysOptHistView(n) = IIf(GetOptionSetting("HistView", "1") = "1", True, False)
80            SysOptMicroScreen(n) = IIf(GetOptionSetting("MicroScreen", "1") = "1", True, False)
90            SysOptMicroSpecific(n) = IIf(GetOptionSetting("MicroSpecific", "1") = "1", True, False)
100           SysOptBioValFore(n) = IIf(GetOptionSetting("BioValFore", "1") = "1", True, False)
110           SysOptBioST(n) = GetOptionSetting("BioST", "")
120           SysOptRTFView(n) = IIf(GetOptionSetting("RTFView", "1") = "1", True, False)
130           SysOptPhone(n) = IIf(GetOptionSetting("PhoneLog", "1") = "1", True, False)


140           SysOptUrgentRef(n) = 0.063
150           SysOptWardDate(n) = 7300
160           SysOptUCVal(n) = 1
              'SysOptBioST(n) = "S"
              'SysOptMicroScreen(n) = 0
170           sql = "SELECT * from Options " & _
                    "order by ListOrder"

180           Set tb = New Recordset
190           RecOpenServer n, tb, sql
200           Do While Not tb.EOF
210               Select Case UCase$(Trim$(tb!Description & ""))
                  Case "URCREAVAL": SysOptUCVal(n) = Trim(tb!Contents & "")
220               Case "IMMPROT": SysOptImmCodeForProt(n) = Trim(tb!Contents & "")
230               Case "IMMUPARA": SysOptImmCodeForUPara(n) = Trim(tb!Contents & "")
240               Case "IMMALB": SysOptImmCodeForAlb(n) = Trim(tb!Contents & "")
250               Case "IMMPARA1": SysOptImmCodeForPara1(n) = Trim(tb!Contents & "")
260               Case "IMMPARA2": SysOptImmCodeForPara2(n) = Trim(tb!Contents & "")
270               Case "IMMPARA3": SysOptImmCodeForPara3(n) = Trim(tb!Contents & "")
280               Case "IMMIGG": SysOptImmCodeForIGG(n) = Trim(tb!Contents & "")
290               Case "IMMIGA": SysOptImmCodeForIGA(n) = Trim(tb!Contents & "")
300               Case "IMMIGM": SysOptImmCodeForIGM(n) = Trim(tb!Contents & "")

310               Case "IMMPROTP": SysOptImmCodeForProtP(n) = Trim(tb!Contents & "")
320               Case "IMMUPARAP": SysOptImmCodeForUParaP(n) = Trim(tb!Contents & "")
330               Case "IMMALBP": SysOptImmCodeForAlbP(n) = Trim(tb!Contents & "")
340               Case "IMMPARA1P": SysOptImmCodeForPara1P(n) = Trim(tb!Contents & "")
350               Case "IMMPARA2P": SysOptImmCodeForPara2P(n) = Trim(tb!Contents & "")
360               Case "IMMPARA3P": SysOptImmCodeForPara3P(n) = Trim(tb!Contents & "")
370               Case "IMMIGGP": SysOptImmCodeForIGGP(n) = Trim(tb!Contents & "")
380               Case "IMMIGAP": SysOptImmCodeForIGAP(n) = Trim(tb!Contents & "")
390               Case "IMMIGMP": SysOptImmCodeForIGMP(n) = Trim(tb!Contents & "")
400               Case "IMB2M": SysOptImmCodeForB2M(n) = Trim(tb!Contents & "")

410               Case "ALLOWDEMOPRINT": SysOptAllowDemoPrint(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
420               Case "FAXCOM": SysOptFaxCom(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
430               Case "WARDVAL": SysOptWardDate(n) = Trim(tb!Contents & "")
440               Case "HISTOSAMPS": SysOptHistoSamps(n) = (Trim(tb!Contents & ""))
450               Case "BIOSAMP": SysOptBioSamp(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
460               Case "NOSEERF": SysOptNoSeeRF(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
470               Case "BIOBAD": SysOptBioCodeForBad(n) = Trim$(tb!Contents & "")
480               Case "NIL10": SysOptNil10(n) = Trim(tb!Contents & "")
490               Case "FULLFAECES": SysOptFullFaeces(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
500               Case "VIEWHFLAG": SysOptViewHFlag(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
510               Case "EBAD": SysOptEBad(n) = Trim(tb!Contents & "")
520               Case "COAGBAD": SysOptCBad(n) = Trim(tb!Contents & "")

530               Case "ALLOWWARDFREETEXT": SysOptAllowWardFreeText(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
540               Case "ALLOWCLINICIANFREETEXT": SysOptAllowClinicianFreeText(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
550               Case "ALLOWGPFREETEXT": SysOptAllowGPFreeText(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
560               Case "ALPHAORDERTECHNICIANS": SysOptAlphaOrderTechnicians(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
570               Case "ALWAYSREQUESTFBC": SysOptAlwaysRequestFBC(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
580               Case "AUTOVALURGENT": SysOptUrgentRef(n) = IIf(Val(Trim$(tb!Contents & "")) = 0, 0.042, Val(tb!Contents))
590               Case "BADRES": SysOptBadRes(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
600               Case "BIOCODEFORGLUCOSE": SysOptBioCodeForGlucose(n) = Trim$(tb!Contents & "")
610               Case "BIOCODEFORGLUCOSEP": SysOptBioCodeForGlucoseP(n) = Trim$(tb!Contents & "")
620               Case "BIOCODEFORGLUCOSE1": SysOptBioCodeForGlucose1(n) = Trim$(tb!Contents & "")
630               Case "BIOCODEFORGLUCOSE2": SysOptBioCodeForGlucose2(n) = Trim$(tb!Contents & "")
640               Case "BIOCODEFORGLUCOSE3": SysOptBioCodeForGlucose3(n) = Trim$(tb!Contents & "")
650               Case "BIOCODEFORFASTGLUCOSE": SysOptBioCodeForFastGlucose(n) = Trim$(tb!Contents & "")
660               Case "BIOCODEFORGLUCOSE1P": SysOptBioCodeForGlucose1P(n) = Trim$(tb!Contents & "")
670               Case "BIOCODEFORGLUCOSE2P": SysOptBioCodeForGlucose2P(n) = Trim$(tb!Contents & "")
680               Case "BIOCODEFORGLUCOSE3P": SysOptBioCodeForGlucose3P(n) = Trim$(tb!Contents & "")
690               Case "BIOCODEFORFASTGLUCOSEP": SysOptBioCodeForFastGlucoseP(n) = Trim$(tb!Contents & "")
700               Case "BIOCODEFORGENT": SysOptBioCodeForGent(n) = Trim$(tb!Contents & "")
710               Case "BIOCODEFORPGENT": SysOptBioCodeForPGent(n) = Trim$(tb!Contents & "")
720               Case "BIOCODEFORTGENT": SysOptBioCodeForTGent(n) = Trim$(tb!Contents & "")
730               Case "BIOCODEFORCHOL": SysOptBioCodeForChol(n) = Trim$(tb!Contents & "")
740               Case "BIOCODEFORCHOLP": SysOptBioCodeForCholP(n) = Trim$(tb!Contents & "")
750               Case "BIOCODEFORHDL": SysOptBioCodeForHDL(n) = Trim$(tb!Contents & "")
760               Case "BIOCODEFORCHOLHDLRATIO": SysOptBioCodeForCholHDLRatio(n) = Trim$(tb!Contents & "")
770               Case "BIOCODEFORTRIG": SysOptBioCodeForTrig(n) = Trim$(tb!Contents & "")
780               Case "BIOCODEFORTRIGP": SysOptBioCodeForTrigP(n) = Trim$(tb!Contents & "")
790               Case "BIOCODEFORHBA1C": SysOptBioCodeForHbA1c(n) = Trim$(tb!Contents & "")
800               Case "BIOCODEFORCREAT": SysOptBioCodeForCreat(n) = Trim$(tb!Contents & "")
810               Case "BIOCODEFORUCREAT": SysOptBioCodeForUCreat(n) = Trim$(tb!Contents & "")
820               Case "BIOCODEFORUPROT": SysOptBioCodeForUProt(n) = Trim$(tb!Contents & "")
830               Case "BIOCODEFORALB": SysOptBioCodeForAlb(n) = Trim$(tb!Contents & "")
840               Case "BIOCODEFORGLOB": SysOptBioCodeForGlob(n) = Trim$(tb!Contents & "")
850               Case "BIOCODEFORTPROT": SysOptBioCodeForTProt(n) = Trim$(tb!Contents & "")
860               Case "BIOCODEFOR24VOL": SysOptBioCodeFor24Vol(n) = Trim$(tb!Contents & "")
870               Case "BIOCODEFOR24UPROT": SysOptBioCodeFor24UProt(n) = Trim$(tb!Contents & "")
880               Case "BIOCODEFORCREATCLEAR": SysOptBioCodeForCreatClear(n) = Trim$(tb!Contents & "")
890               Case "BIOCODEFORUNA": SysOptBioCodeForUNa(n) = Trim$(tb!Contents & "")
900               Case "BIOCODEFORUK": SysOptBioCodeForUK(n) = Trim$(tb!Contents & "")
910               Case "BIOCODEFORUMAG": SysOptBioCodeForUMag(n) = Trim$(tb!Contents & "")
920               Case "BIOCODEFORUPHOS": SysOptBioCodeForUPhos(n) = Trim$(tb!Contents & "")
930               Case "BIOCODEFORUCA": SysOptBioCodeForUCA(n) = Trim$(tb!Contents & "")
940               Case "BIOCODEFORUCHOL": SysOptBioCodeForUChol(n) = Trim$(tb!Contents & "")
950               Case "BIOCODEFORUUREA": SysOptBioCodeForUUrea(n) = Trim$(tb!Contents & "")

960               Case "BIOCODEFOR24UCREAT": SysOptBioCodeFor24UCreat(n) = Trim$(tb!Contents & "")
970               Case "BIOCODEFOR24UNA": SysOptBioCodeFor24UNa(n) = Trim$(tb!Contents & "")
980               Case "BIOCODEFOR24UK": SysOptBioCodeFor24UK(n) = Trim$(tb!Contents & "")
990               Case "BIOCODEFOR24UMAG": SysOptBioCodeFor24UMag(n) = Trim$(tb!Contents & "")
1000              Case "BIOCODEFOR24UPHOS": SysOptBioCodeFor24UPhos(n) = Trim$(tb!Contents & "")
1010              Case "BIOCODEFOR24UCA": SysOptBioCodeFor24UCA(n) = Trim$(tb!Contents & "")
1020              Case "BIOCODEFOR24UCHOL": SysOptBioCodeFor24UChol(n) = Trim$(tb!Contents & "")
1030              Case "BIOCODEFOR24UUREA": SysOptBioCodeFor24UUrea(n) = Trim$(tb!Contents & "")

1040              Case "BIOPHONE": SysOptBioPhone(n) = Trim$(tb!Contents & "")
1050              Case "COAGPHONE": SysOptCoagPhone(n) = Trim$(tb!Contents & "")
1060              Case "CHANGE": SysOptChange(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
1070              Case "COMMVAL": SysOptCommVal(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
1080              Case "CUMSHOW": SysOptNoCumShow(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
1090              Case "CYTOOFFSET": SysOptCytoOffset(n) = Val(Trim$(tb!Contents & ""))
1100              Case "DEFAULTABS": SysOptDefaultABs(n) = Val(tb!Contents & "")
1110              Case "DEMOSHOW": SysOptDemo(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
1120              Case "DEMOVAL": SysOptDemoVal(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
1130              Case "DEPTBGA": SysOptDeptBga(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
1140              Case "DEPTBIO": SysOptDeptBio(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
1150              Case "DEPTCOAG": SysOptDeptCoag(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
1160              Case "DEPTCYTO": SysOptDeptCyto(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
1170              Case "DEPTEND": SysOptDeptEnd(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
1180              Case "DEPTEXT": SysOptDeptExt(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
1190              Case "DEPTHAEM": SysOptDeptHaem(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
1200              Case "DEPTHISTO": SysOptDeptHisto(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
1210              Case "DEPTIMM": SysOptDeptImm(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
1220              Case "DEPTMICRO": SysOptDeptMicro(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
1230              Case "DEPTSEMEN": SysOptDeptSemen(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
1240              Case "DEPTMEDIBRIDGE": SysOptDeptMedibridge(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
1250              Case "DISABLEPRACTICES": SysOptDisablePractices(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
1260              Case "DISABLEWARDORDERING": SysOptDisableWardOrdering(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
1270              Case "DIPSTICK": SysOptDipStick(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
1280              Case "DONTCOAG": SysOptDontShowPrevCoag(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
1290              Case "SAMPLETIME": SysOptSampleTime(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
1300              Case "EXP": SysOptExp(n) = Trim$(tb!Contents & "")
1310              Case "HAEMAN1": SysOptHaemAn1(n) = Trim$(tb!Contents & "")
1320              Case "HAEMAN2": SysOptHaemAn2(n) = Trim$(tb!Contents & "")
1330              Case "HAEMPHONE": SysOptHaemPhone(n) = Trim$(tb!Contents & "")
1340              Case "HISTOOFFSET": SysOptHistoOffset(n) = Val(Trim$(tb!Contents & ""))
1350              Case "HIVCODE": SysOptHivCode(n) = Trim$(tb!Contents & "")
1360              Case "HOSPITAL": SysOptHospital(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
1370              Case "IMMPHONE": SysOptImmPhone(n) = Trim$(tb!Contents & "")
1380              Case "MICROOFFSET": SysOptMicroOffset(n) = Val(Trim$(tb!Contents & ""))
1390              Case "NUMLEN": SysOptNumLen(n) = Trim$(tb!Contents & "")
1400              Case "OPTPASS": SysOptOptPass(n) = Trim$(tb!Contents & "")
1410              Case "PNE": SysOptPNE(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
1420              Case "PRINTALL": SysOptPrintAll(n) = Trim$(tb!Contents & "")
1430              Case "SEMENOFFSET": SysOptSemenOffset(n) = Val(Trim$(tb!Contents & ""))
1440              Case "SHORTFAECES": SysOptShortFaeces(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
1450              Case "URGENT": SysOptUrgent(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
1460              Case "USEFULLID": SysOptUseFullID(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
1470              Case "VIEWBLOOD": SysOptViewTrans(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
1480              Case "CHECKCHOLHDLRATIO": SysOptCheckCholHDLRatio(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
1490              Case "REALIMM": SysOptRealImm(0) = IIf(Trim$(tb!Contents & "") = "1", True, False)

1500              Case "DOASSGLUCOSE": SysOptDoAssGlucose(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)

1510              Case "LONGORSHORTBIONAMES": SysOptLongOrShortBioNames(n) = Trim$(tb!Contents & "")
1520              Case "HAEMN1": SysOptHaemN1(n) = Trim$(tb!Contents & "")
1530              Case "HAEMN2": SysOptHaemN2(n) = Trim$(tb!Contents & "")
1540              Case "BION1": SysOptBioN1(n) = Trim$(tb!Contents & "")
1550              Case "BION2": SysOptBioN2(n) = Trim$(tb!Contents & "")
1560              Case "VIEW": SysOptView(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
1570              Case "WBCDC": SysOptWBCDC(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
1580              Case "MAXSAMPLEURINEBATCHPRINTING": SysOptMaxSampleUrineBatchPrinting(n) = Val(tb!Contents & "")
1590              Case "SHOWIQ200": SysOptShowIQ200(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
1600              End Select
1610              tb.MoveNext
1620          Loop
1630      Next

1640      Exit Sub

LoadOptions_Error:

          Dim strES As String
          Dim intEL As Integer



1650      intEL = Erl
1660      strES = Err.Description
1670      LogError "basOptions", "LoadOptions", intEL, strES, sql

End Sub

Public Sub LoadFormOptions(ByRef Opts() As udtOptionList)

          Dim tb As Recordset
          Dim sql As String
          Dim n As Integer

10        On Error GoTo LoadFormOptions_Error

20        For n = 0 To UBound(Opts)
30            sql = "SELECT * FROM Options WHERE " & _
                    "Description = '" & Opts(n).Description & "'"
40            Set tb = New Recordset
50            RecOpenClient 0, tb, sql
60            If Not tb.EOF Then
70                Opts(n).Value = Trim$(tb!Contents & "")
80            End If
90        Next

100       Exit Sub

LoadFormOptions_Error:

          Dim strES As String
          Dim intEL As Integer

110       intEL = Erl
120       strES = Err.Description
130       LogError "basOptions", "LoadFormOptions", intEL, strES, sql

End Sub

Private Sub ReDimOptions()
10        ReDim sysOptAllowCopyDemographics(0 To intOtherHospitalsInGroup) As Boolean
20        ReDim SysOptMicroAna(0 To intOtherHospitalsInGroup) As Boolean
30        ReDim SysOptBioMask(0 To intOtherHospitalsInGroup) As Boolean
40        ReDim SysOptHistView(0 To intOtherHospitalsInGroup) As Boolean
50        ReDim SysOptMicroScreen(0 To intOtherHospitalsInGroup) As Long
60        ReDim SysOptMicroSpecific(0 To intOtherHospitalsInGroup) As Boolean
70        ReDim SysOptBioValFore(0 To intOtherHospitalsInGroup) As Boolean
80        ReDim SysOptBioST(0 To intOtherHospitalsInGroup) As String
90        ReDim SysOptPhone(0 To intOtherHospitalsInGroup) As Boolean
100       ReDim SysOptUCVal(0 To intOtherHospitalsInGroup) As Long
110       ReDim SysOptImmCodeForProt(0 To intOtherHospitalsInGroup) As String
120       ReDim SysOptImmCodeForAlb(0 To intOtherHospitalsInGroup) As String
130       ReDim SysOptImmCodeForPara1(0 To intOtherHospitalsInGroup) As String
140       ReDim SysOptImmCodeForPara2(0 To intOtherHospitalsInGroup) As String
150       ReDim SysOptImmCodeForPara3(0 To intOtherHospitalsInGroup) As String
160       ReDim SysOptImmCodeForIGG(0 To intOtherHospitalsInGroup) As String
170       ReDim SysOptImmCodeForIGA(0 To intOtherHospitalsInGroup) As String
180       ReDim SysOptImmCodeForIGM(0 To intOtherHospitalsInGroup) As String
190       ReDim SysOptImmCodeForB2M(0 To intOtherHospitalsInGroup) As String
200       ReDim SysOptImmCodeForUPara(0 To intOtherHospitalsInGroup) As String
210       ReDim SysOptImmCodeForProtP(0 To intOtherHospitalsInGroup) As String
220       ReDim SysOptImmCodeForAlbP(0 To intOtherHospitalsInGroup) As String
230       ReDim SysOptImmCodeForPara1P(0 To intOtherHospitalsInGroup) As String
240       ReDim SysOptImmCodeForPara2P(0 To intOtherHospitalsInGroup) As String
250       ReDim SysOptImmCodeForPara3P(0 To intOtherHospitalsInGroup) As String
260       ReDim SysOptImmCodeForIGGP(0 To intOtherHospitalsInGroup) As String
270       ReDim SysOptImmCodeForIGAP(0 To intOtherHospitalsInGroup) As String
280       ReDim SysOptImmCodeForIGMP(0 To intOtherHospitalsInGroup) As String
290       ReDim SysOptImmCodeForUParaP(0 To intOtherHospitalsInGroup) As String

300       ReDim SysOptRTFView(0 To intOtherHospitalsInGroup) As Boolean
310       ReDim SysOptAllowDemoPrint(0 To intOtherHospitalsInGroup) As Boolean
320       ReDim SysOptFaxCom(0 To intOtherHospitalsInGroup) As Boolean
330       ReDim SysOptHistoSamps(0 To intOtherHospitalsInGroup) As Long
340       ReDim SysOptBioSamp(0 To intOtherHospitalsInGroup) As Boolean
350       ReDim SysOptWardDate(0 To intOtherHospitalsInGroup) As Long    'Minus Rundate Value
360       ReDim SysOptNoSeeRF(0 To intOtherHospitalsInGroup) As Boolean
370       ReDim SysOptBioCodeForBad(0 To intOtherHospitalsInGroup) As String
380       ReDim SysOptNil10(0 To intOtherHospitalsInGroup) As String
390       ReDim SysOptFullFaeces(0 To intOtherHospitalsInGroup) As Boolean
400       ReDim SysOptViewHFlag(0 To intOtherHospitalsInGroup) As Boolean
410       ReDim SysOptRealImm(0 To intOtherHospitalsInGroup) As Boolean
420       ReDim SysOptEBad(0 To intOtherHospitalsInGroup) As String
430       ReDim SysOptCBad(0 To intOtherHospitalsInGroup) As String
440       ReDim SysOptClearHaem(0 To intOtherHospitalsInGroup) As Boolean
450       ReDim SysOptViewTrans(0 To intOtherHospitalsInGroup) As Boolean
460       ReDim SysOptNoCumShow(0 To intOtherHospitalsInGroup) As Boolean
470       ReDim SysOptSampleTime(0 To intOtherHospitalsInGroup) As Boolean
480       ReDim SysOptHivCode(0 To intOtherHospitalsInGroup) As String
490       ReDim SysOptOptPass(0 To intOtherHospitalsInGroup) As String
500       ReDim SysOptUrgent(0 To intOtherHospitalsInGroup) As Boolean
510       ReDim SysOptBadRes(0 To intOtherHospitalsInGroup) As Boolean
520       ReDim SysOptGpClin(0 To intOtherHospitalsInGroup) As Boolean
530       ReDim SysOptDemoVal(0 To intOtherHospitalsInGroup) As Boolean
540       ReDim SysOptCommVal(0 To intOtherHospitalsInGroup) As Boolean
550       ReDim SysOptPrintAll(0 To intOtherHospitalsInGroup) As String
560       ReDim SysSetFoc(0 To intOtherHospitalsInGroup) As String
570       ReDim SysOptDemo(0 To intOtherHospitalsInGroup) As Boolean
580       ReDim SysOptDeptHaem(0 To intOtherHospitalsInGroup) As Boolean
590       ReDim SysOptDeptBio(0 To intOtherHospitalsInGroup) As Boolean
600       ReDim SysOptDeptCoag(0 To intOtherHospitalsInGroup) As Boolean
610       ReDim SysOptDeptMicro(0 To intOtherHospitalsInGroup) As Boolean
620       ReDim SysOptDeptImm(0 To intOtherHospitalsInGroup) As Boolean
630       ReDim SysOptDeptEnd(0 To intOtherHospitalsInGroup) As Boolean
640       ReDim SysOptDeptBga(0 To intOtherHospitalsInGroup) As Boolean
650       ReDim SysOptDeptExt(0 To intOtherHospitalsInGroup) As Boolean
660       ReDim SysOptDeptSemen(0 To intOtherHospitalsInGroup) As Boolean
670       ReDim SysOptDeptCyto(0 To intOtherHospitalsInGroup) As Boolean
680       ReDim SysOptDeptHisto(0 To intOtherHospitalsInGroup) As Boolean
690       ReDim SysOptDeptMedibridge(0 To intOtherHospitalsInGroup) As Boolean
700       ReDim SysOptHospital(0 To intOtherHospitalsInGroup) As Boolean
710       ReDim SysOptDontShowPrevCoag(0 To intOtherHospitalsInGroup) As Boolean

720       ReDim SysOptUrgentRef(0 To intOtherHospitalsInGroup) As Double
730       ReDim SysOptNumLen(0 To intOtherHospitalsInGroup) As Double  'Sample Number Length

          'added by myles
740       ReDim SysOptView(0 To intOtherHospitalsInGroup) As Boolean
750       ReDim SysOptExp(0 To intOtherHospitalsInGroup) As Boolean
760       ReDim SysOptHaemN1(0 To intOtherHospitalsInGroup) As String
770       ReDim SysOptHaemN2(0 To intOtherHospitalsInGroup) As String
780       ReDim SysOptBioN1(0 To intOtherHospitalsInGroup) As String
790       ReDim SysOptBioN2(0 To intOtherHospitalsInGroup) As String
800       ReDim SysOptChange(0 To intOtherHospitalsInGroup) As Boolean

810       ReDim SysOptBioCodeForGent(0 To intOtherHospitalsInGroup) As String
820       ReDim SysOptBioCodeForPGent(0 To intOtherHospitalsInGroup) As String
830       ReDim SysOptBioCodeForTGent(0 To intOtherHospitalsInGroup) As String

840       ReDim SysOptBioCodeForCreat(0 To intOtherHospitalsInGroup) As String
850       ReDim SysOptBioCodeForUCreat(0 To intOtherHospitalsInGroup) As String
860       ReDim SysOptBioCodeForUProt(0 To intOtherHospitalsInGroup) As String
870       ReDim SysOptBioCodeForAlb(0 To intOtherHospitalsInGroup) As String
880       ReDim SysOptBioCodeForGlob(0 To intOtherHospitalsInGroup) As String
890       ReDim SysOptBioCodeForTProt(0 To intOtherHospitalsInGroup) As String
900       ReDim SysOptBioCodeFor24UProt(0 To intOtherHospitalsInGroup) As String
910       ReDim SysOptBioCodeFor24Vol(0 To intOtherHospitalsInGroup) As String

920       ReDim SysOptBioCodeForUCA(0 To intOtherHospitalsInGroup) As String
930       ReDim SysOptBioCodeForUChol(0 To intOtherHospitalsInGroup) As String
940       ReDim SysOptBioCodeForUMag(0 To intOtherHospitalsInGroup) As String
950       ReDim SysOptBioCodeForUK(0 To intOtherHospitalsInGroup) As String
960       ReDim SysOptBioCodeForUNa(0 To intOtherHospitalsInGroup) As String
970       ReDim SysOptBioCodeForUPhos(0 To intOtherHospitalsInGroup) As String
980       ReDim SysOptBioCodeForUUrea(0 To intOtherHospitalsInGroup) As String


990       ReDim SysOptBioCodeFor24UCA(0 To intOtherHospitalsInGroup) As String
1000      ReDim SysOptBioCodeFor24UChol(0 To intOtherHospitalsInGroup) As String
1010      ReDim SysOptBioCodeFor24UMag(0 To intOtherHospitalsInGroup) As String
1020      ReDim SysOptBioCodeFor24UK(0 To intOtherHospitalsInGroup) As String
1030      ReDim SysOptBioCodeFor24UNa(0 To intOtherHospitalsInGroup) As String
1040      ReDim SysOptBioCodeFor24UPhos(0 To intOtherHospitalsInGroup) As String
1050      ReDim SysOptBioCodeFor24UUrea(0 To intOtherHospitalsInGroup) As String
1060      ReDim SysOptBioCodeFor24UCreat(0 To intOtherHospitalsInGroup) As String

1070      ReDim SysOptHaemPhone(0 To intOtherHospitalsInGroup) As String
1080      ReDim SysOptBioPhone(0 To intOtherHospitalsInGroup) As String
1090      ReDim SysOptCoagPhone(0 To intOtherHospitalsInGroup) As String
1100      ReDim SysOptBloodPhone(0 To intOtherHospitalsInGroup) As String
1110      ReDim SysOptImmPhone(0 To intOtherHospitalsInGroup) As String

1120      ReDim SysOptHaemAn1(0 To intOtherHospitalsInGroup) As String
1130      ReDim SysOptHaemAn2(0 To intOtherHospitalsInGroup) As String

1140      ReDim SysOptSemenOffset(0 To intOtherHospitalsInGroup) As Double  '100,000,000,000
1150      ReDim SysOptMicroOffset(0 To intOtherHospitalsInGroup) As Double  '200,000,000,000
1160      ReDim SysOptHistoOffset(0 To intOtherHospitalsInGroup) As Long    '30,000,000
1170      ReDim SysOptCytoOffset(0 To intOtherHospitalsInGroup) As Long     '30,000,000

1180      ReDim SysOptPNE(0 To intOtherHospitalsInGroup) As Boolean

1190      ReDim SysOptBioCodeForGlucose(0 To intOtherHospitalsInGroup) As String
1200      ReDim SysOptBioCodeForGlucose1(0 To intOtherHospitalsInGroup) As String
1210      ReDim SysOptBioCodeForGlucose2(0 To intOtherHospitalsInGroup) As String
1220      ReDim SysOptBioCodeForGlucose3(0 To intOtherHospitalsInGroup) As String
1230      ReDim SysOptBioCodeForFastGlucose(0 To intOtherHospitalsInGroup) As String
1240      ReDim SysOptBioCodeForChol(0 To intOtherHospitalsInGroup) As String
1250      ReDim SysOptBioCodeForHDL(0 To intOtherHospitalsInGroup) As String
1260      ReDim SysOptBioCodeForTrig(0 To intOtherHospitalsInGroup) As String
1270      ReDim SysOptBioCodeForCholHDLRatio(0 To intOtherHospitalsInGroup) As String
1280      ReDim SysOptBioCodeForHbA1c(0 To intOtherHospitalsInGroup) As String
1290      ReDim SysOptBioCodeForGlucoseP(0 To intOtherHospitalsInGroup) As String
1300      ReDim SysOptBioCodeForGlucose1P(0 To intOtherHospitalsInGroup) As String
1310      ReDim SysOptBioCodeForGlucose2P(0 To intOtherHospitalsInGroup) As String
1320      ReDim SysOptBioCodeForGlucose3P(0 To intOtherHospitalsInGroup) As String
1330      ReDim SysOptBioCodeForFastGlucoseP(0 To intOtherHospitalsInGroup) As String
1340      ReDim SysOptBioCodeForCholP(0 To intOtherHospitalsInGroup) As String
1350      ReDim SysOptBioCodeForTrigP(0 To intOtherHospitalsInGroup) As String

1360      ReDim SysOptCheckCholHDLRatio(0 To intOtherHospitalsInGroup) As Boolean

1370      ReDim SysOptDipStick(0 To intOtherHospitalsInGroup) As Boolean
1380      ReDim SysOptUseFullID(0 To intOtherHospitalsInGroup) As Boolean
1390      ReDim SysOptDefaultABs(0 To intOtherHospitalsInGroup) As Long

1400      ReDim SysOptShortFaeces(0 To intOtherHospitalsInGroup) As Boolean

1410      ReDim SysOptBloodBank(0 To intOtherHospitalsInGroup) As Boolean
1420      ReDim SysOptRemote(0 To intOtherHospitalsInGroup) As Boolean

1430      ReDim SysOptDisablePractices(0 To intOtherHospitalsInGroup) As Boolean
1440      ReDim SysOptDisableWardOrdering(0 To intOtherHospitalsInGroup) As Boolean

1450      ReDim SysOptAllowWardFreeText(0 To intOtherHospitalsInGroup) As Boolean
1460      ReDim SysOptAllowClinicianFreeText(0 To intOtherHospitalsInGroup) As Boolean
1470      ReDim SysOptAllowGPFreeText(0 To intOtherHospitalsInGroup) As Boolean

1480      ReDim SysOptAlwaysRequestFBC(0 To intOtherHospitalsInGroup) As Boolean

1490      ReDim SysOptDoAssGlucose(0 To intOtherHospitalsInGroup) As Boolean

1500      ReDim SysOptLongOrShortBioNames(0 To intOtherHospitalsInGroup) As String

1510      ReDim SysOptWBCDC(0 To intOtherHospitalsInGroup)
1520      ReDim SysOptAlphaOrderTechnicians(0 To intOtherHospitalsInGroup)
1530      ReDim SysOptBioCodeForCreatClear(0 To intOtherHospitalsInGroup)

1540      ReDim SysOptMaxSampleUrineBatchPrinting(0 To intOtherHospitalsInGroup)
1550      ReDim SysOptShowIQ200(0 To intOtherHospitalsInGroup) As Boolean

End Sub


Public Sub LoadUserOpts()
          Dim sql As String
          Dim n As Long
          Dim tb As New Recordset


10        On Error GoTo LoadUserOpts_Error

20        ReDimUserOpts

30        On Error GoTo LoadUserOpts_Error

40        For n = 0 To intOtherHospitalsInGroup
50            SysOptPlasBack(n) = vbGreen
60            SysOptPlasFore(n) = vbWhite
70            SysOptHighBack(n) = vbRed
80            SysOptHighFore(n) = vbYellow
90            SysOptLowBack(n) = vbBlue
100           SysOptLowFore(n) = vbYellow
110           SysOptGpClin(n) = False
120           SysOptExtDefault(n) = False
130           SysOptToolTip(n) = True
140           sql = "SELECT * from Options WHERE username = '" & AddTicks(UserName) & "' " & _
                    "order by ListOrder"

150           Set tb = New Recordset
160           RecOpenServer n, tb, sql
170           Do While Not tb.EOF
180               Select Case UCase$(Trim$(tb!Description & ""))
                  Case "TOOLTIP": SysOptToolTip(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
190               Case "SETFOC": SysSetFoc(n) = Trim(tb!Contents)
200               Case "AUTOREF": SysOptAutoRef(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
210               Case "DEFAULTTAB": SysOptDefaultTab(n) = Trim$(tb!Contents & "")
220               Case "FONT": SysOptFont(n) = Trim$(tb!Contents & "")
230               Case "LOWBACK": If Trim$(tb!Contents & "") <> "" Then SysOptLowBack(n) = Trim$(tb!Contents & "")
240               Case "LOWFORE": If Trim$(tb!Contents & "") <> "" Then SysOptLowFore(n) = Trim$(tb!Contents & "")
250               Case "PLASBACK": If Trim$(tb!Contents & "") <> "" Then SysOptPlasBack(n) = Trim$(tb!Contents & "")
260               Case "PLASFORE": If Trim$(tb!Contents & "") <> "" Then SysOptPlasFore(n) = Trim$(tb!Contents & "")
270               Case "HIGHBACK": If Trim$(tb!Contents & "") <> "" Then SysOptHighBack(n) = Trim$(tb!Contents & "")
280               Case "HIGHFORE": If Trim$(tb!Contents & "") <> "" Then SysOptHighFore(n) = Trim$(tb!Contents & "")
290               Case "GPCLIN": SysOptGpClin(n) = Trim(tb!Contents & "")
300               Case "EXTDEFAULT": SysOptExtDefault(n) = Trim(tb!Contents & "")
310               Case "CLEARHAEM": SysOptClearHaem(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
320               End Select
330               tb.MoveNext
340           Loop
350       Next






360       Exit Sub

LoadUserOpts_Error:

          Dim strES As String
          Dim intEL As Integer



370       intEL = Erl
380       strES = Err.Description
390       LogError "basOptions", "LoadUserOpts", intEL, strES, sql


End Sub

Private Sub ReDimUserOpts()

10        ReDim SysOptAutoRef(0 To intOtherHospitalsInGroup) As Boolean
20        ReDim SysOptDefaultTab(0 To intOtherHospitalsInGroup) As String
30        ReDim SysOptLowFore(0 To intOtherHospitalsInGroup) As String
40        ReDim SysOptLowBack(0 To intOtherHospitalsInGroup) As String
50        ReDim SysOptPlasFore(0 To intOtherHospitalsInGroup) As String
60        ReDim SysOptPlasBack(0 To intOtherHospitalsInGroup) As String
70        ReDim SysOptHighFore(0 To intOtherHospitalsInGroup) As String
80        ReDim SysOptHighBack(0 To intOtherHospitalsInGroup) As String
90        ReDim SysOptFont(0 To intOtherHospitalsInGroup) As String
100       ReDim SysOptExtDefault(0 To intOtherHospitalsInGroup) As Boolean
110       ReDim SysOptToolTip(0 To intOtherHospitalsInGroup) As Boolean

End Sub



