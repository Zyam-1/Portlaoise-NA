Attribute VB_Name = "basBoxes"
Option Explicit
Public StrEvent As String

Public Function iMsg(Optional ByVal Message As String, _
                     Optional ByVal T As Long = 0, _
                     Optional ByVal Caption As String = "NetAcquire", _
                     Optional ByVal BckColour As Long = &HC0C000, _
                     Optional ByVal MsgFontSize As Long) _
                     As Long

          Dim SafeMsgBox As New fcdrMsgBox

10        On Error GoTo iMsg_Error

20        With SafeMsgBox
30            .MsgFontSize = MsgFontSize
40            .BackColor = BckColour
50            .DisplayButtons = T And &H7
60            .DefaultButton = T And &H300
70            .ShowIcon = T And &H70
80            .Message = Message
90            .Caption = Caption
100           .Show vbModal
110           iMsg = .RetVal
120       End With

130       Unload SafeMsgBox
140       Set SafeMsgBox = Nothing

150       Exit Function

iMsg_Error:

          Dim strES As String
          Dim intEL As Integer



160       intEL = Erl
170       strES = Err.Description
180       LogError "basBoxes", "iMsg", intEL, strES


End Function

Public Function iBOX(ByVal Prompt As String, _
                     Optional ByVal Title As String = "NetAcquire", _
                     Optional ByVal Default As String, _
                     Optional ByVal Pass As Boolean) As String

          Dim Box As New fcdrInputBox

10        On Error GoTo iBOX_Error

20        With Box
30            .PassWord = Pass
40            .Caption = Title
50            .lblPrompt = Prompt
60            .txtInput = Default
70            .Show vbModal
80            iBOX = .RetVal
90        End With

100       Unload Box
110       Set Box = Nothing

120       Exit Function

iBOX_Error:

          Dim strES As String
          Dim intEL As Integer



130       intEL = Erl
140       strES = Err.Description
150       LogError "basBoxes", "iBOX", intEL, strES


End Function


Public Function iTIME(ByVal Prompt As String, _
                      Optional ByVal Title As String = "NetAcquire", _
                      Optional ByVal Default As String = "__:__") As String

          Dim Box As New frmcdrInputTime

10        With Box
20            .Caption = Title
30            .lblPrompt = Prompt
40            .txtIP = Default
50            .Show vbModal
60            iTIME = .RetVal
70        End With

80        Unload Box
90        Set Box = Nothing

End Function

Public Function FlagMessage(ByVal strType As String, _
                            ByVal Historical As String, _
                            ByVal current As String, _
                            Optional ByVal SampleID As String = "") _
                            As Boolean
      'Returns True to reject

          Dim s As String
          Dim RetVal As Boolean

10        If Trim$(Historical) = "" Then Historical = "<Blank>"
20        If Trim$(current) = "" Then current = "<Blank>"

30        s = "Patients " & strType & " has changed!" & vbCrLf & _
              "Was '" & Historical & "'" & vbCrLf & _
              "Now '" & current & "'" & vbCrLf & _
              "To accept this change, Press 'OK'"

40        RetVal = iMsg(s, vbCritical + vbOKCancel, "Critical Warning") = vbCancel

50        If Not RetVal Then
60            StrEvent = Trim$(SampleID & " Name Change accepted. (") & Replace(s, vbCrLf, " ") & ")"
70        End If

80        FlagMessage = RetVal

End Function

