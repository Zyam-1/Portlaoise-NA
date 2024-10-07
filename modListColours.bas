Attribute VB_Name = "modListColours"
Option Explicit

Public Type ListColour
    ListText As String
    BackColour As Long
    ForeColour As Long
End Type

Public Sub CycleLabel(ByRef ListOfItems() As ListColour, _
                      ByRef lbl As Label)

          Dim n As Integer
          Dim X As Integer
          Dim Found As Boolean

10        On Error GoTo CycleLabel_Error

20        Found = False
30        X = 1
40        For n = 0 To UBound(ListOfItems)
50            If lbl = ListOfItems(n).ListText Then
60                Found = True
70                If n = UBound(ListOfItems) Then
80                    X = 0
90                Else
100                   X = n + 1
110               End If
120               Exit For
130           End If
140       Next

150       If Not Found Then
160           X = 0
170       End If

180       lbl.Caption = ListOfItems(X).ListText
190       lbl.BackColor = ListOfItems(X).BackColour
200       lbl.ForeColor = ListOfItems(X).ForeColour

210       Exit Sub

CycleLabel_Error:

          Dim strES As String
          Dim intEL As Integer

220       intEL = Erl
230       strES = Err.Description
240       LogError "modListColours", "CycleLabel", intEL, strES

End Sub


Public Sub CycleTextBox(ByRef ListOfItems() As ListColour, _
                        ByRef txt As TextBox)

          Dim n As Integer
          Dim X As Integer
          Dim Found As Boolean

10        On Error GoTo CycleTextBox_Error

20        Found = False
30        X = 1
40        For n = 0 To UBound(ListOfItems)
50            If txt.Text = ListOfItems(n).ListText Then
60                Found = True
70                If n = UBound(ListOfItems) Then
80                    X = 0
90                Else
100                   X = n + 1
110               End If
120               Exit For
130           End If
140       Next

150       If Not Found Then
160           X = 0
170       End If

180       txt.Text = ListOfItems(X).ListText
190       txt.BackColor = ListOfItems(X).BackColour
200       txt.ForeColor = ListOfItems(X).ForeColour

210       Exit Sub

CycleTextBox_Error:

          Dim strES As String
          Dim intEL As Integer

220       intEL = Erl
230       strES = Err.Description
240       LogError "modListColours", "CycleTextBox", intEL, strES

End Sub


Public Sub CycleGridCell(ByRef ListOfItems() As ListColour, _
                         ByVal Row As Integer, _
                         ByVal Col As Integer, _
                         ByRef g As MSFlexGrid)

          Dim n As Integer
          Dim X As Integer
          Dim Found As Boolean

10        On Error GoTo CycleGridCell_Error

20        Found = False
30        X = 1
40        For n = 0 To UBound(ListOfItems)
50            If g.TextMatrix(Row, Col) = ListOfItems(n).ListText Then
60                Found = True
70                If n = UBound(ListOfItems) Then
80                    X = 0
90                Else
100                   X = n + 1
110               End If
120               Exit For
130           End If
140       Next

150       If Not Found Then
160           X = 0
170       End If

180       g.TextMatrix(Row, Col) = ListOfItems(X).ListText
190       g.Row = Row
200       g.Col = Col
210       g.CellBackColor = ListOfItems(X).BackColour
220       g.CellForeColor = ListOfItems(X).ForeColour

230       Exit Sub

CycleGridCell_Error:

          Dim strES As String
          Dim intEL As Integer

240       intEL = Erl
250       strES = Err.Description
260       LogError "modListColours", "CycleGridCell", intEL, strES

End Sub



Public Sub LoadListGenericColour(ByRef lst() As ListColour, _
                                 ByVal ListType As String)

          Dim sql As String
          Dim tb As Recordset
          Dim c() As String

10        On Error GoTo LoadListGenericColour_Error

20        ReDim lst(0 To 0) As ListColour
30        lst(0).ListText = ""
40        lst(0).BackColour = vbButtonFace
50        lst(0).ForeColour = vbBlack

60        sql = "SELECT * FROM Lists WHERE " & _
                "ListType = '" & ListType & "' " & _
                "ORDER BY ListOrder"
70        Set tb = New Recordset
80        RecOpenServer 0, tb, sql
90        Do While Not tb.EOF
100           ReDim Preserve lst(0 To UBound(lst) + 1)
110           lst(UBound(lst)).ListText = tb!Text & ""
120           c = Split(tb!Default & "", "|")
130           If UBound(c) < 1 Then
140               lst(UBound(lst)).BackColour = vbButtonFace
150               lst(UBound(lst)).ForeColour = vbBlack
160           Else
170               lst(UBound(lst)).BackColour = Val(c(0))
180               lst(UBound(lst)).ForeColour = Val(c(1))
190           End If
200           tb.MoveNext
210       Loop

220       Exit Sub

LoadListGenericColour_Error:

          Dim strES As String
          Dim intEL As Integer

230       intEL = Erl
240       strES = Err.Description
250       LogError "modListColours", "LoadListGenericColour", intEL, strES, sql

End Sub


