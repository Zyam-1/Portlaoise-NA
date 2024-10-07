Attribute VB_Name = "modRegistration"
Option Explicit


Public Function Get_Reg() As String
      Dim Hdd As String
      Dim Reg As String

10    On Error GoTo Get_Reg_Error

20    Hdd = GetHDDSerial

30    Reg = (Val(Hdd) * 50) / 6 + Val(Format(Now, "yyyymmdd"))

40    Get_Reg = Format(Val(Reg), "###")

50    Exit Function

Get_Reg_Error:

      Dim strES As String
      Dim intEL As Integer


60    intEL = Erl
70    strES = Err.Description
80    LogError "modRegistration", "Get_Reg", intEL, strES


End Function


Public Function Check_Reg() As Boolean

      Dim g As String
      Dim Reg As String


10    On Error GoTo Check_Reg_Error

20    If GetSetting("NetAcquire", "NetAcquire", "Inval") <> "Yes" Then
30        g = Get_Reg
40        Reg = Gen_Reg(g)
50        If iBOX("Enter Registration!" & "Key - " & g) = Reg Then
60            SaveSetting "NetAcquire", "NetAcquire", "Inval", "Yes"
70            Check_Reg = True
80        Else
90            Check_Reg = False
100       End If
110   Else
120       Check_Reg = True '
130   End If


140   Exit Function

Check_Reg_Error:

      Dim strES As String
      Dim intEL As Integer


150   intEL = Erl
160   strES = Err.Description
170   LogError "modRegistration", "Check_Reg", intEL, strES


End Function

Public Function Gen_Reg(ByVal g As Double) As String
      Dim Reg As String

10    On Error GoTo Gen_Reg_Error

20    Reg = (g * 34) / 7 + Val(Format(Now + 1, "yyyymmdd"))

30    Gen_Reg = Format(Val(Reg), "####")

40    Exit Function

Gen_Reg_Error:

      Dim strES As String
      Dim intEL As Integer


50    intEL = Erl
60    strES = Err.Description
70    LogError "modRegistration", "Gen_Reg", intEL, strES


End Function
