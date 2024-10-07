Attribute VB_Name = "modNoConstant"
Option Explicit

Private Declare Function GetPrivateProfileString Lib "Kernel32" Alias _
                                                 "GetPrivateProfileStringA" _
                                                 (ByVal lpSectionName As String, ByVal lpKeyName As String, _
                                                  ByVal lpDefault As String, ByVal lpbuffurnedString As String, _
                                                  ByVal nBuffSize As Long, ByVal lpFileName As String) As Long

Private Declare Function GetPrivateProfileSectionNames Lib "Kernel32.dll" Alias _
                                                       "GetPrivateProfileSectionNamesA" _
                                                       (ByVal lpszReturnBuffer As String, _
                                                        ByVal nSize As Long, _
                                                        ByVal lpFileName As String) As Long

Private Function GetPass(ByRef UID As String) As String

          Dim P As String
          Dim A As String
          Dim n As Integer

10        A = ""
20        For n = 97 To 122
30            A = A & chr$(n)
40        Next
50        For n = 65 To 90
60            A = A & chr$(n)
70        Next

80        A = A & "!£$%^&*()<>-_+={}[]:@~||;'#,./?"
90        For n = 48 To 57
100           A = A & chr$(n)
110       Next

          '    abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ!£$%^&*()<>-_+={}[]:@~||;'#,./?0123456789"
          '    123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123
          '             1         2         3         4         5         6         7         8         9

          'p = ""
          'UID = "sa"

          'LabUser
120       UID = Mid$(A, 38, 1) & Mid$(A, 1, 1) & Mid$(A, 2, 1) & Mid$(A, 47, 1) & _
                Mid$(A, 19, 1) & Mid$(A, 5, 1) & Mid$(A, 18, 1)

          'DfySiywtgtw$1>)=
130       P = Mid$(A, 30, 1) & Mid$(A, 6, 1) & Mid$(A, 25, 1) & Mid$(A, 45, 1) & _
              Mid$(A, 9, 1) & Mid$(A, 25, 1) & Mid$(A, 23, 1) & Mid$(A, 20, 1) & _
              Mid$(A, 7, 1) & Mid$(A, 20, 1) & Mid$(A, 23, 1) & Mid$(A, 55, 1) & _
              Mid$(A, 85, 1) & Mid$(A, 63, 1) & Mid$(A, 61, 1) & Mid$(A, 67, 1)

140       GetPass = P

End Function

Public Sub ConnectToDatabase()

          Dim dbDSN As String
          Dim dbDSNbb As String
          Dim dbRemoteDSNbb As String
          Dim tb As Recordset
          Dim TempCnxn As Connection
          Dim dbConnectRemoteBB As String
          Dim Con As String
          Dim ConBB As String
          Dim i As Integer

10        On Error GoTo ConnectToDatabase_Error


20        HospName(0) = GetcurrentConnectInfo(Con, ConBB, "")

          'HospName(1) = GetcurrentConnectInfo2(Con, ConBB) ' Masood 03-04-2014

30        If IsIDE And HospName(0) = "" Then
40            MsgBox "INI Error"
50            End
60        ElseIf HospName(0) = "" Then
70            If GetConnectInfo("Active", Con, HospName(0)) Then
80                GetConnectInfo "BB", ConBB
90                GetConnectInfo "RemoteBB", dbConnectRemoteBB
100           Else
110               Set TempCnxn = New Connection
120               TempCnxn.Open "uid=sa;dsn=Constant;"
130               Set tb = New Recordset
140               With tb
150                   .CursorLocation = adUseServer
160                   .CursorType = adOpenDynamic
170                   .LockType = adLockOptimistic
180                   .ActiveConnection = TempCnxn
190                   .Source = "Select * from Constant where active = 1"
200                   .Open
210               End With

220               dbDSN = tb!DSN & ""
230               dbDSNbb = tb!DSNBB & ""
240               dbRemoteDSNbb = tb!RemoteDSNbb & ""

250               HospName(i) = Trim$(tb!Hosp & "")

260               Con = "uid=sa;pwd=;dsn=" & dbDSN & ";"
270               If dbDSNbb <> "" Then ConBB = "uid=sa;pwd=;dsn=" & dbDSNbb & ";"
280               If dbRemoteDSNbb <> "" Then dbConnectRemoteBB = "uid=sa;pwd=;dsn=" & dbRemoteDSNbb & ";"
290           End If

300       End If

310       Set Cnxn(0) = New Connection
320       Cnxn(0).Open Con



330       HospName(1) = GetcurrentConnectInfo(Con, ConBB, "1")
340       If HospName(1) <> "" Then
350           Set Cnxn(1) = New Connection
360           Cnxn(1).Open (Con)
370       End If

380       HospName(2) = GetcurrentConnectInfo(Con, ConBB, "2")
390       If HospName(2) <> "" Then
400           Set Cnxn(2) = New Connection
410           Cnxn(2).Open (Con)
420       End If

430       Exit Sub

ConnectToDatabase_Error:

          Dim strES As String
          Dim intEL As Integer


440       intEL = Erl
450       strES = Err.Description
460       LogError "modNoConstant", "ConnectToDatabase", intEL, strES

End Sub
Public Function GetConnectInfo(ByVal ConnectTo As String, _
                               ByRef ReturnConnectionString As String, _
                               Optional ByRef HospName As Variant) As Boolean

      'ConnectTo = "Active"
      '            "BB"
      '            "Active" & n - HospitalGroup
      '            "BB" & n - HospitalGroup

10        On Error GoTo GetConnectInfo_Error

20        GetConnectInfo = False

30        If Not IsMissing(HospName) Then
40            HospName = GetSetting("NetAcquire", "HospName", ConnectTo, "")
50            If Left$(UCase$(HospName), 5) = "LOCAL" Then
60                HospName = Mid$(HospName, 6)
70            End If
80        End If

90        ReturnConnectionString = GetSetting("NetAcquire", "Cnxn", ConnectTo, "")


100       If Trim$(ReturnConnectionString) <> "" Then

110           ReturnConnectionString = Obfuscate(ReturnConnectionString)

120           GetConnectInfo = True

130       End If

140       Exit Function

GetConnectInfo_Error:

          Dim strES As String
          Dim intEL As Integer


150       intEL = Erl
160       strES = Err.Description
170       LogError "modNoConstant", "GetConnectInfo", intEL, strES


End Function


Public Function GetcurrentConnectInfo(ByRef Con As String, ByRef ConBB As String, Optional ConIndex As String = "") As String

          'Returns Hospital Name

          Dim HospitalNames() As String
          Dim n As Long
          Dim HospitalName As String
          Dim retHospitalName As String
          Dim ServerName As String
          Dim NetAcquireDB As String
          Dim TransfusionDB As String
          Dim UID As String
          Dim PWD As String
          Dim CurrentPath As String

10        On Error GoTo GetcurrentConnectInfo_Error

20        CurrentPath = App.Path & "\NetAcquire.INI"


30        HospitalNames = GetINISectionNames(CurrentPath, Val(ConIndex) * 8)
40        If UBound(HospitalNames) < Val(ConIndex) Then
50            Exit Function
60        End If
70        HospitalName = HospitalNames(Val(ConIndex))
80        If Left$(UCase$(HospitalName), 5) = "LOCAL" Then
90            retHospitalName = Mid$(HospitalName, 6)
100       Else
110           retHospitalName = HospitalName
120       End If
           retHospitalName = "Portlaoise"


140       NetAcquireDB = ProfileGetItem(HospitalName, "D" & ConIndex, "", CurrentPath)
150       TransfusionDB = ProfileGetItem(HospitalName, "T" & ConIndex, "", CurrentPath)

160       'PWD = GetPass(UID)
          'PWD = "DfySiywtgtw$1>)="
'170       Con = "DRIVER={SQL Server};" & _
'              "Server=" & Obfuscate(ServerName) & ";" & _
'              "Database=" & Obfuscate(NetAcquireDB) & ";" & _
'              "uid=" & UID & ";" & _
'              "pwd=" & PWD & ";"
'180       Con = "DRIVER={SQL Server};" & _
'              "Server=" & "WIN-LS4D35ITV6L" & ";" & _
'              "Database=" & "PortLive" & ";" & _
'              "uid=" & "usman" & ";" & _
'              "pwd=" & "usman123" & ";"
190       Con = "Provider=SQLOLEDB;" & _
              "Data Source=" & "DESKTOP-3OMS1N5\SQLEXPRESS" & ";" & _
              "Initial Catalog=" & "PortLive" & ";" & _
              "Integrated Security=SSPI;"
          'MsgBox Con
200       If TransfusionDB <> "" Then
210           ConBB = "DRIVER={SQL Server};" & _
                  "Server=" & Obfuscate(ServerName) & ";" & _
                  "Database=" & Obfuscate(TransfusionDB) & ";" & _
                  "uid=" & UID & ";" & _
                  "pwd=" & PWD & ";"
220       End If
          'MsgBox ConBB

230       If ConIndex = "" Then
240           frmMain.mnuHospital(0).Caption = retHospitalName
250           frmMain.mnuHospital(0).Tag = Con
260       End If
270       GetcurrentConnectInfo = retHospitalName

280       Exit Function

GetcurrentConnectInfo_Error:

          Dim strES As String
          Dim intEL As Integer

290       intEL = Erl
300       strES = Err.Description
310       MsgBox "GetCurrentConnectInfo Error Line " & intEL
          'Resume
320       LogError "modNoConstant", "GetcurrentConnectInfo", intEL, strES

End Function


'Public Function GetcurrentConnectInfo(ByRef Con As String, ByRef ConBB As String, ConIndex As String) As String
'
'      'Returns Hospital Name
'
'      Dim HospitalNames() As String
'      Dim n As Long
'      Dim HospitalName As String
'      Dim retHospitalName As String
'      Dim ServerName() As String
'      Dim NetAcquireDB As String
'      Dim TransfusionDB As String
'      Dim UID As String
'      Dim PWD As String
'      Dim CurrentPath As String
'
'10    On Error GoTo GetcurrentConnectInfo_Error
''20    If IsIDE Then
''30      CurrentPath = "C:\ClientCode\NetAcquire.INI"
''40    Else
'50        CurrentPath = App.Path & "\NetAcquire.INI"
''60    End If
'
'70    HospitalNames = GetINISectionNames(CurrentPath, n)
'80    HospitalName = HospitalNames(0)
'90    If Left$(UCase$(HospitalName), 5) = "LOCAL" Then
'100     retHospitalName = Mid$(HospitalName, 6)
'110   Else
'120     retHospitalName = HospitalName
'130   End If
'
'140   ServerName = ProfileGetItem(HospitalName, "N" & ConIndex, "", CurrentPath)
'150   NetAcquireDB = ProfileGetItem(HospitalName, "D" & ConIndex, "", CurrentPath)
'160   TransfusionDB = ProfileGetItem(HospitalName, "T" & ConIndex, "", CurrentPath)
'
'170   PWD = GetPass(UID)
'
'180   Con = "DRIVER={SQL Server};" & _
 '            "Server=" & Obfuscate(ServerName) & ";" & _
 '            "Database=" & Obfuscate(NetAcquireDB) & ";" & _
 '            "uid=" & UID & ";" & _
 '            "pwd=" & PWD & ";"
'
'190   If TransfusionDB <> "" Then
'200     ConBB = "DRIVER={SQL Server};" & _
 '                "Server=" & Obfuscate(ServerName) & ";" & _
 '                "Database=" & Obfuscate(TransfusionDB) & ";" & _
 '                "uid=" & UID & ";" & _
 '                "pwd=" & PWD & ";"
'210   End If
'
'220   frmMain.mnuHospital(0).Caption = retHospitalName
'230   frmMain.mnuHospital(0).Tag = Con
'
'240   GetcurrentConnectInfo = retHospitalName
'
'250   Exit Function
'
'GetcurrentConnectInfo_Error:
'
'      Dim strES As String
'      Dim intEL As Integer
'
'260   intEL = Erl
'270   strES = Err.Description
'280   MsgBox "GetCurrentConnectInfo Error Line " & intEL
'290   LogError "modNoConstant", "GetcurrentConnectInfo", intEL, strES
'
'End Function


Private Function ProfileGetItem(ByRef sSection As String, _
                                ByRef sKeyName As String, _
                                ByRef sDefValue As String, _
                                ByRef sIniFile As String) As String

      'retrieves a value from an ini file
      'corresponding to the section and
      'key name passed.

          Dim dwSize As Integer
          Dim nBuffSize As Integer
          Dim buff As String
          Dim RetVal As String

          'Call the API with the parameters passed.
          'nBuffSize is the length of the string
          'in buff, including the terminating null.
          'If a default value was passed, and the
          'section or key name are not in the file,
          'that value is returned. If no default
          'value was passed (""), then dwSize
          'will = 0 if not found.
          '
          'pad a string large enough to hold the data
10        On Error GoTo ProfileGetItem_Error

20        buff = Space(2048)
30        nBuffSize = Len(buff)
40        dwSize = GetPrivateProfileString(sSection, sKeyName, sDefValue, buff, nBuffSize, sIniFile)

50        If dwSize > 0 Then
60            RetVal = Left$(buff, dwSize)
70        End If

80        ProfileGetItem = RetVal

90        Exit Function

ProfileGetItem_Error:

          Dim strES As String
          Dim intEL As Integer


100       intEL = Erl
110       strES = Err.Description
120       LogError "modNoConstant", "ProfileGetItem", intEL, strES


End Function

Private Function GetINISectionNames(ByRef inFile As String, ByRef outCount As Long) As String()

          Dim StrBuf As String
          Dim BufLen As Long
          Dim RetVal() As String
          Dim Count As Long

10        On Error GoTo GetINISectionNames_Error

20        BufLen = 16

30        Do
40            BufLen = BufLen * 2
50            StrBuf = Space$(BufLen)
60            Count = GetPrivateProfileSectionNames(StrBuf, BufLen, inFile)
70        Loop While Count = BufLen - 2

80        If (Count) Then
90            RetVal = Split(Left$(StrBuf, Count - 1), vbNullChar)
100           outCount = UBound(RetVal) + 1
110       End If

120       GetINISectionNames = RetVal

130       Exit Function

GetINISectionNames_Error:

          Dim strES As String
          Dim intEL As Integer


140       intEL = Erl
150       strES = Err.Description
160       LogError "modNoConstant", "GetINISectionNames", intEL, strES


End Function


Public Function Obfuscate(ByVal strData As String) As String

          Dim lngI As Long
          Dim lngJ As Long

10        On Error GoTo Obfuscate_Error

20        For lngI = 0 To Len(strData) \ 4
30            For lngJ = 1 To 4
40                Obfuscate = Obfuscate & Mid$(strData, (4 * lngI) + 5 - lngJ, 1)
50            Next
60        Next

70        Exit Function

Obfuscate_Error:

          Dim strES As String
          Dim intEL As Integer


80        intEL = Erl
90        strES = Err.Description
100       LogError "modNoConstant", "Obfuscate", intEL, strES


End Function

