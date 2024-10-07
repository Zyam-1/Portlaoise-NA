VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HaemTestDefinitions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'In a shared module put the following declaration
'Public colHaemTestDefinitions As New HaemTestDefinitions

Private mcolHaemTestDefinitions As New Collection

Private mHaemTestDefinitions As New HaemTestDefinitions



Public Sub Refresh()

Set mcolHaemTestDefinitions = Nothing

Load

End Sub


Public Property Get HaemTestDefinitions() As HaemTestDefinitions
  
Set HaemTestDefinitions = mHaemTestDefinitions

End Property


Public Function Add(ByVal HD As HaemTestDefinition) _
                    As HaemTestDefinition

Dim tb As New Recordset
Dim SQL As String



On Error GoTo Add_Error

With HD
  SQL = "SELECT * from HaemTestDefinitions WHERE " & _
        "AnalyteName = '" & .AnalyteName & "' " & _
        "and AgeFromDays = " & .AgeFromDays & " " & _
        "and AgeToDays = " & .AgeToDays & " " & _
        "And Hospital = '" & Hospname(0) & "'"
  Set tb = New Recordset
  RecOpenClient 0, tb, SQL
  If tb.EOF Then
    tb.AddNew
  End If
  tb!AnalyteName = .AnalyteName
  tb!MaleLow = .MaleLow
  tb!MaleHigh = .MaleHigh
  tb!FemaleLow = .FemaleLow
  tb!FemaleHigh = .FemaleHigh
  tb!LowControlMean = .LowControlMean
  tb!NormalControlMean = .NormalControlMean
  tb!HighControlMean = .HighControlMean
  tb!LowControl1SD = .LowControl1SD
  tb!NormalControl1SD = .NormalControl1SD
  tb!HighControl1SD = .HighControl1SD
  tb!DoRM = .DoRM
  tb!DoDelta = .DoDelta
  tb!DeltaValue = .DeltaValue
  tb!AgeFromDays = .AgeFromDays
  tb!AgeToDays = .AgeToDays
  tb!PlausibleHigh = .PlausibleHigh
  tb!PlausibleLow = .PlausibleLow
  tb!InUse = .InUse
  tb!Hospital = Hospname(0)
  tb!Printformat = .Printformat
  tb.Update

End With

Refresh

Set Add = HD



Exit Function

Add_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " HaemTestDefinitions /Add ")
  Case 1:    End
  Case 2:    Exit Function
  Case 3:    Resume Next
End Select


End Function

Public Function Item(ByVal AnalyteName As String, _
                     ByVal AgeFromDays As Long, _
                     ByVal AgeToDays As Long, _
                     Optional ByVal Hosp As String) _
                     As HaemTestDefinition
Attribute Item.VB_UserMemId = 0
  


On Error GoTo Item_Error

If Hosp = "" Then Hosp = Hospname(0)

Set Item = mcolHaemTestDefinitions.Item(AnalyteName & " " & CStr(AgeFromDays) & " " & CStr(AgeToDays) & " " & Hosp)

Exit Function

Item_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " HaemTestDefinitions /Item ")
  Case 1:    End
  Case 2:    Exit Function
  Case 3:    Resume Next
End Select


End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  
Set NewEnum = mcolHaemTestDefinitions.[_NewEnum]

End Function

Public Function Count() As Long
  
Count = mcolHaemTestDefinitions.Count

End Function


Private Sub Class_Initialize()

Load

End Sub
Private Sub Load()

Dim tb As New Recordset
Dim HD As HaemTestDefinition
Dim SQL As String


On Error GoTo Load_Error

SQL = "SELECT * from HaemTestDefinitions"

Set tb = New Recordset
RecOpenServer 0, tb, SQL
Do While Not tb.EOF
  Set HD = New HaemTestDefinition
  With HD
    .AnalyteName = tb!AnalyteName & ""
    .MaleLow = IIf(IsNull(tb!MaleLow), 0, tb!MaleLow)
    .MaleHigh = IIf(IsNull(tb!MaleHigh), 9999, tb!MaleHigh)
    .FemaleLow = IIf(IsNull(tb!FemaleLow), 0, tb!FemaleLow)
    .FemaleHigh = IIf(IsNull(tb!FemaleHigh), 9999, tb!FemaleHigh)
    .LowControlMean = IIf(IsNull(tb!LowControlMean), 0, tb!LowControlMean)
    .NormalControlMean = IIf(IsNull(tb!NormalControlMean), 0, tb!NormalControlMean)
    .HighControlMean = IIf(IsNull(tb!HighControlMean), 0, tb!HighControlMean)
    .LowControl1SD = IIf(IsNull(tb!LowControl1SD), 9999, tb!LowControl1SD)
    .NormalControl1SD = IIf(IsNull(tb!NormalControl1SD), 9999, tb!NormalControl1SD)
    .HighControl1SD = IIf(IsNull(tb!HighControl1SD), 9999, tb!HighControl1SD)
    .DoRM = tb!DoRM
    .DoDelta = tb!DoDelta
    .DeltaValue = IIf(IsNull(tb!DeltaValue), 9999, tb!DeltaValue)
    .PlausibleHigh = IIf(IsNull(tb!PlausibleHigh), 999, tb!PlausibleHigh)
    .PlausibleLow = IIf(IsNull(tb!PlausibleLow), 0, tb!PlausibleLow)
    .AgeFromDays = IIf(IsNull(tb!AgeFromDays), 0, tb!AgeFromDays)
    .AgeToDays = IIf(IsNull(tb!AgeToDays), MaxAgeToDays, tb!AgeToDays)
    .InUse = tb!InUse
    .Category = tb!Category & ""
    .Hospital = UCase(Trim$(tb!Hospital & ""))
    .Printformat = IIf(IsNull(tb!Printformat), 1, tb!Printformat)
    If IsNull(tb!Eod) Then .Eod = False Else .Eod = tb!Eod
    mcolHaemTestDefinitions.Add HD, .AnalyteName & " " & CStr(.AgeFromDays) & " " & CStr(.AgeToDays) & " " & .Category & " " & .Hospital
  End With
  tb.MoveNext
Loop

Exit Sub


Exit Sub

Load_Error:

Dim er As Long
Dim ers As String

er = Err.Number
ers = Err.Description

Select Case ToDo(er, ers, " HaemTestDefinitions /Load ")
  Case 1:    End
  Case 2:    Exit Sub
  Case 3:    Resume Next
End Select


End Sub




