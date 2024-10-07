Attribute VB_Name = "modCreateODBC"
Option Explicit

Declare Function SQLConfigDataSource Lib "odbccp32.dll" (ByVal hwndParent As _
Long, ByVal fRequest As Long, ByVal lpszDriver As String, ByVal _
lpszAttributes As String) As Long


Function PrepareDSN(ByVal strServerName As String, ByVal strDBName As String, _
ByVal strDSN As String) As Boolean

          Dim boolError As Boolean
          Dim strDSNString As String
10    On Error GoTo PrepareDSN_Error

20        PrepareDSN = False

30        strDSNString = Space(1000)
40        strDSNString = ""
50        strDSNString = strDSNString & "DSN=" & strDSN & Chr(0)
60        strDSNString = strDSNString & "DESCRIPTION=" & "DSN Created Dynamically On " & CStr(Now) & Chr(0)
70        strDSNString = strDSNString & "Server=" & strServerName & Chr(0)
80        strDSNString = strDSNString & "DATABASE=" & strDBName & Chr(0)
90        strDSNString = strDSNString & Chr(0)

100      If Not CBool(SQLConfigDataSource(0, _
                              4, _
                              "SQL Server", _
                              strDSNString)) Then
110           boolError = True
120           MsgBox ("Error in PrepareDSN::SQLConfigDataSource")
130      End If

140       If boolError Then
150           Exit Function
160       End If
170       PrepareDSN = True
180       Exit Function

190   Exit Function

PrepareDSN_Error:

      Dim strES As String
      Dim intEL As Integer


200   intEL = Erl
210   strES = Err.Description
220   LogError "modCreateODBC", "PrepareDSN", intEL, strES


End Function

Public Sub Check_Odbc()
      Dim strUser As String
      Dim strData As String
      Dim strDSN As String
      Dim strServer As String

10    On Error GoTo Check_Odbc_Error

20    If GetSetting("NetAcquire", "NetAcquire", "Constant") <> "Yes" Then
30         strServer = iBOX("Enter Sever for Constant", , "(local)")
40         strDSN = iBOX("Dsn Name for Constant", , "Constant")
50         strData = iBOX("Database Name for Constant", , "Constant")
60         strUser = iBOX("User for Constant", , "sa")
70         If PrepareDSN(strServer, strData, strDSN) = False Then
80            iMsg "ODBC Set up error! System shutting down"
90            Exit Sub
100        Else
110          SaveSetting "NetAcquire", "NetAcquire", "Constant", "Yes"
120        End If
130   End If
          

140   If GetSetting("NetAcquire", "NetAcquire", "Lab") <> "Yes" Then
150        strServer = iBOX("Enter Sever for Live", , "(local)")
160        strDSN = iBOX("Dsn Name for Live", , "LabLive")
170        strData = iBOX("Database Name for Live", , "Lablive")
180        strUser = iBOX("User for Live", , "sa")
190        If PrepareDSN(strServer, strData, strDSN) = False Then
200           iMsg "ODBC Set up error! System shutting down"
210           Exit Sub
220        Else
230          SaveSetting "NetAcquire", "NetAcquire", "Lab", "Yes"
240        End If
250   End If

260   Exit Sub

Check_Odbc_Error:

      Dim strES As String
      Dim intEL As Integer


270   intEL = Erl
280   strES = Err.Description
290   LogError "modCreateODBC", "Check_Odbc", intEL, strES

          
End Sub

