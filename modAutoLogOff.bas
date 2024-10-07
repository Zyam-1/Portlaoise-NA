Attribute VB_Name = "modAutoLogOff"
Option Explicit

Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, _
                                                      lpPoint As POINTAPI) As Long

Declare Function GetForegroundWindow Lib "user32" () As Long
Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Declare Function GetKeyboardState Lib "user32.dll" (pbKeyState As Byte) As Long

Public Function KB() As Boolean

          Dim keystat(0 To 255) As Byte    ' receives key status information for all keys
          Dim RetVal As Long    ' return value of function
          Dim n As Integer
          Static strSave As String
          Dim s As String

10        RetVal = GetKeyboardState(keystat(0))    ' In VB, the array is passed by referencing element #0.

20        s = ""
30        For n = 0 To 255
40            s = s & CStr(keystat(n))
50        Next

60        If s <> strSave Then
70            strSave = s
80            KB = True
90        Else
100           KB = False
110       End If

End Function

Function MouseX(Optional ByVal hWnd As Long) As Long
          Dim lpPoint As POINTAPI
10        GetCursorPos lpPoint
20        If hWnd Then ScreenToClient hWnd, lpPoint
30        MouseX = lpPoint.X
End Function

' Get mouse Y coordinates in pixels
'
' If a window handle is passed, the result is relative to the client area
' of that window, otherwise the result is relative to the screen

Function MouseY(Optional ByVal hWnd As Long) As Long
          Dim lpPoint As POINTAPI
10        GetCursorPos lpPoint
20        If hWnd Then ScreenToClient hWnd, lpPoint
30        MouseY = lpPoint.Y
End Function



Public Function TopMostWindow() As String

          Dim h As Long
          Dim c As Long
          Dim lpString As String * 255
          Dim T As String * 255
          Dim X As Long

10        h = GetForegroundWindow()
20        c = GetClassName(h, lpString, 255)
30        X = GetWindowText(h, T, 255)

40        TopMostWindow = Left$(T, X)

End Function


