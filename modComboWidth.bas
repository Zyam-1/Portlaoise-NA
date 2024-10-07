Attribute VB_Name = "modComboWidth"
Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal _
                                                                        hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
                                                                        lParam As Any) As Long
Const CB_SETDROPPEDWIDTH = &H160


Public Sub SetComboDropDownWidth(ComboBox As ComboBox)

          Dim n As Integer
          Dim w As Long
          Dim Max As Long

10        Max = 0
20        For n = 0 To ComboBox.ListCount - 1
30            w = frmEditMicrobiologyNew.TextWidth(ComboBox.List(n))
40            If w > Max Then
50                Max = w
60            End If
70        Next

80        w = (Max + 400) / Screen.TwipsPerPixelX

90        SendMessage ComboBox.hWnd, CB_SETDROPPEDWIDTH, w, ByVal 0&

End Sub
