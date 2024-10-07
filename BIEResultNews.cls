VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BIEResultNews"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private mcolBIEResultNews As New Collection
Private mBIEResultNews As New BIEResultNews

Public Function Add(br As BIEResultNew) As BIEResultNew

10    On Error GoTo Add_Error

20    mcolBIEResultNews.Add br

30    Set Add = br

40    Exit Function

Add_Error:

      Dim strES As String
      Dim intEL As Integer


50    intEL = Erl
60    strES = Err.Description
70    LogError "BIEResultNews", "Add", intEL, strES


End Function

Public Property Get BIEResultNews() As BIEResultNews
  
10    On Error GoTo BIEResultNews_Error

20    Set BIEResultNews = mBIEResultNews

30    Exit Property

BIEResultNews_Error:

      Dim strES As String
      Dim intEL As Integer


40    intEL = Erl
50    strES = Err.Description
60    LogError "BIEResultNews", "BIEResultNews", intEL, strES


End Property

Public Sub Clear()

      Dim n As Long

10    On Error GoTo Clear_Error

20    For n = 1 To mcolBIEResultNews.Count
30      mcolBIEResultNews.Remove 1
40    Next

50    Exit Sub

Clear_Error:

      Dim strES As String
      Dim intEL As Integer


60    intEL = Erl
70    strES = Err.Description
80    LogError "BIEResultNews", "Clear", intEL, strES


End Sub

Public Function Count() As Long
  
10    On Error GoTo Count_Error

20    Count = mcolBIEResultNews.Count

30    Exit Function

Count_Error:

      Dim strES As String
      Dim intEL As Integer


40    intEL = Erl
50    strES = Err.Description
60    LogError "BIEResultNews", "Count", intEL, strES


End Function

Public Function Item(ByVal x As Long) _
                     As BIEResultNew
  
10    On Error GoTo Item_Error

20    On Error Resume Next

30    Set Item = mcolBIEResultNews(x)

40    Exit Function

Item_Error:

      Dim strES As String
      Dim intEL As Integer


50    intEL = Erl
60    strES = Err.Description
70    LogError "BIEResultNews", "Item", intEL, strES


End Function


Public Function Load(ByVal Discipline As String, _
                     ByVal SampleID As String, _
                     ByVal ResultOrRepeat As String, _
                     ByVal v As Long, _
                     ByVal P As Long, _
                     ByVal Cn As Long, _
                     ByVal Cat As String, _
                     ByVal Rundate As String) _
                     As BIEResultNews
      'Discipline is either "Bio", "Imm" or "End"

      Dim BRs As New BIEResultNews
      Dim br As BIEResultNew
      Dim tb As New Recordset
      Dim sql As String
      Dim DaysOld As Long
      Dim TableName As String

10    On Error GoTo Load_Error

20    DaysOld = 9125

30    If SampleID = "" Then Exit Function

40    If UCase(Discipline) = "BIO" Or UCase(Discipline) = "BGA" Then Cat = ""

50    TableName = Discipline & ResultOrRepeat

60    sql = "select * from " & TableName & " where sampleid = '" & SampleID & "'"

70    If P = gNOTPRINTED And v = gNOTVALID Then
80      sql = sql & "and Printed = 0 and Valid = 0 "
90    ElseIf P = gNOTPRINTED And v = gVALID Then
100     sql = sql & "and Printed = 0 and Valid = 1 "
110   ElseIf P = gNOTPRINTED And v = gDONTCARE Then
120     sql = sql & "and Printed = 0 "
130   ElseIf P = gPRINTED And v = gNOTVALID Then
140     sql = sql & "and Printed = 1 and Valid = 0 "
150   ElseIf P = gPRINTED And v = gVALID Then
160     sql = sql & "and Printed = 1 and Valid = 1 "
170   ElseIf P = gPRINTED And v = gDONTCARE Then
180     sql = sql & "and Printed = 1 "
190   ElseIf P = gDONTCARE And v = gNOTVALID Then
200     sql = sql & "and Valid = 0 "
210   ElseIf P = gDONTCARE And v = gVALID Then
220     sql = sql & "and Valid = 1 "
230   End If
      'sql = sql & "Order by PrintPriority asc"
240   Set tb = New Recordset
250   RecOpenServer Cn, tb, sql '  RecOpenClient 0,tb, Sql
260   Do While Not tb.EOF
270     Set br = New BIEResultNew
280     With br
290       .SampleID = Trim(tb!SampleID & "")
300       .Code = Trim(tb!Code & "")
310       .Result = Trim(tb!Result & "")
320       .Operator = Trim(tb!Operator & "")
330       .Rundate = Format$(tb!Rundate, "dd/mm/yyyy")
340       .RunTime = Format$(tb!RunTime, "dd/mm/yyyy hh:mm")
350       .Units = Trim(tb!Units & "")
360       If Trim(tb!Printed & "") <> "" Then .Printed = IIf(tb!Printed, True, False)
370       If Trim(tb!Valid & "") <> "" Then .Valid = IIf(tb!Valid, True, False)
380       .Flags = Trim(tb!Flags & "")
390       If tb!DefIndex & "" = "" Then .DefIndex = 0 Else .DefIndex = IIf(tb!DefIndex, tb!DefIndex, 0)
400       BRs.Add br
410     End With
420     tb.MoveNext
430   Loop

440   If BRs.Count <> 0 Then
450     Set Load = BRs
460   Else
470     Set Load = Nothing
480   End If
490   Set br = Nothing
500   Set BRs = Nothing


510   Exit Function



520   Exit Function

Load_Error:

      Dim strES As String
      Dim intEL As Integer


530   intEL = Erl
540   strES = Err.Description
550   LogError "BIEResultNews", "Load", intEL, strES, sql


End Function
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  
10    On Error GoTo NewEnum_Error

20    Set NewEnum = mcolBIEResultNews.[_NewEnum]

30    Exit Function

NewEnum_Error:

      Dim strES As String
      Dim intEL As Integer


40    intEL = Erl
50    strES = Err.Description
60    LogError "BIEResultNews", "NewEnum", intEL, strES


End Function



Public Sub RemoveItem(ByVal x As Long)

10    On Error GoTo RemoveItem_Error

20    mcolBIEResultNews.Remove x

30    Exit Sub

RemoveItem_Error:

      Dim strES As String
      Dim intEL As Integer


40    intEL = Erl
50    strES = Err.Description
60    LogError "BIEResultNews", "RemoveItem", intEL, strES


End Sub


