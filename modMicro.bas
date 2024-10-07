Attribute VB_Name = "modMicro"
Option Explicit

Public Sub OrderOnObserva(ByVal SampleIDNoOffset)

          Dim sql As String

10        On Error GoTo OrderOnObserva_Error

20        If Val(SampleIDNoOffset) = 0 Then Exit Sub

30        sql = "IF EXISTS (SELECT * FROM BactOrders WHERE " & _
                "           SampleID = '" & SampleIDNoOffset & "') " & _
                "  UPDATE BactOrders " & _
                "  SET Programmed = 0, " & _
                "  DateTimeOfRecord = getdate() " & _
                "  WHERE SampleID = '" & SampleIDNoOffset & "' " & _
                "ELSE " & _
                "  INSERT INTO BactOrders " & _
                "  (SampleID, Analyser, TestRequested, Programmed, DateTimeOfRecord) " & _
                "  VALUES " & _
                "  ('" & SampleIDNoOffset & "', " & _
                "   'Observa', " & _
                "   '', " & _
                "   0, " & _
                "   getdate())"
40        Cnxn(0).Execute sql

50        Exit Sub

OrderOnObserva_Error:

          Dim strES As String
          Dim intEL As Integer

60        intEL = Erl
70        strES = Err.Description
80        LogError "modMicro", "OrderOnObserva", intEL, strES, sql

End Sub


Public Function GetMicroSite(ByVal SampleID As Double) As String

          Dim tb As Recordset
          Dim sql As String
          Dim RetVal As String

10        On Error GoTo GetMicroSite_Error

20        RetVal = ""

30        sql = "SELECT Site FROM MicroSiteDetails WHERE " & _
                "SampleID = " & SampleID & ""
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        If Not tb.EOF Then
70            RetVal = UCase$(Trim$(tb!Site & ""))
80        End If

90        GetMicroSite = RetVal

100       Exit Function

GetMicroSite_Error:

          Dim strES As String
          Dim intEL As Integer

110       intEL = Erl
120       strES = Err.Description
130       LogError "modPrintMicroRTF", "GetMicroSite", intEL, strES, sql

End Function


