VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BGAResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private mcolBGAResults As New Collection
Private mBGAResults As New BGAResults

Public Function Add(ByVal BG As BGAResult) As BGAResult

mcolBGAResults.Add BG

Set Add = BG

End Function


Public Property Get BGAResults() As BGAResults
  
Set BGAResults = mBGAResults

End Property

Public Sub Clear()

Dim n As Long

For n = 1 To mcolBGAResults.Count
  mcolBGAResults.Remove 1
Next

End Sub




Public Function Count() As Long
  
Count = mcolBGAResults.Count

End Function



Public Function Item(ByVal Index As Integer) As BGAResult
Attribute Item.VB_UserMemId = 0

Set Item = mcolBGAResults.Item(Index)

End Function


Public Function LoadResults(ByVal Tn As Long, ByVal SampleID As String) _
                            As BGAResult

Dim sn As New Recordset
Dim CR As BGAResult
Dim SQL As String



On Error GoTo LoadResults_Error

SQL = "SELECT * from BGAresults WHERE " & _
      "sampleid = '" & SampleID & "'"
Set sn = New Recordset
RecOpenServer Tn, sn, SQL

If Not sn.EOF Then
  Set CR = New BGAResult
  With CR
    .Rundate = Trim(sn!Rundate & "")
    .RunDateTime = Trim(sn!RunDateTime & "")
    .pH = Trim(sn!pH & "")
    .PCO2 = Trim(sn!PCO2 & "")
    .PO2 = Trim(sn!PO2 & "")
    .HCO3 = Trim(sn!HCO3 & "")
    .BE = Trim(sn!BE & "")
    .O2SAT = Trim(sn!O2SAT & "")
    .TotCO2 = Trim(sn!TotCO2 & "")
    If Not IsNull(sn!Valid) Then
      .Valid = IIf(sn!Valid, True, False)
    Else
      .Valid = False
    End If
    If Not IsNull(sn!Printed) Then
      .Printed = IIf(sn!Printed, True, False)
    Else
      .Printed = False
    End If
    .NOPAS = sn!NOPAS & ""
  End With
  Set LoadResults = CR
Else
  Set LoadResults = Nothing
End If




Exit Function

LoadResults_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " BGAResults /LoadResults ")
  Case 1:    End
  Case 2:    Exit Function
  Case 3:    Resume Next
End Select


End Function
Public Sub LogAsPrinted(ByVal SampleID As String)

Dim SQL As String


On Error GoTo LogAsPrinted_Error

SQL = "UPDATE BGAresults " & _
      "set Printed = 1 WHERE " & _
      "sampleid = '" & SampleID & "'"
Cnxn(0).Execute SQL

Exit Sub

LogAsPrinted_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " BGAResults /LogAsPrinted ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Public Sub SaveResults(ByVal Tn As Long, ByVal CR As BGAResult)

Dim tb As New Recordset
Dim SQL As String
  


On Error GoTo SaveResults_Error

SQL = "SELECT * from BGAResults WHERE " & _
      "sampleid = '" & Format$(Val(CR.SampleID)) & "'"
Set tb = New Recordset
RecOpenServer Tn, tb, SQL
tb.AddNew
tb!SampleID = Format$(Val(CR.SampleID))
tb!Rundate = Format(CR.Rundate, "dd/mmm/yyyy")
tb!RunDateTime = Format(CR.RunDateTime, "dd/mmm/yyyy hh:mm:ss")
tb!pH = CR.pH
tb!PCO2 = CR.PCO2
tb!PO2 = CR.PO2
tb!HCO3 = CR.HCO3
tb!BE = CR.BE
tb!O2SAT = CR.O2SAT
tb!TotCO2 = CR.TotCO2
tb!Valid = CR.Valid
tb!Printed = CR.Printed
tb!NOPAS = CR.NOPAS
tb.Update



Exit Sub

SaveResults_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " BGAResults /SaveResults ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  
Set NewEnum = mcolBGAResults.[_NewEnum]

End Function

Public Sub Remove(ByVal Index As Integer)

mcolBGAResults.Remove (Index)

End Sub





