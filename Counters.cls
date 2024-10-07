VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Counters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mcolCounters As New Collection

Private mCounters As New Counters


Public Function Add(ByVal NewC As Counter) _
       As Counter

10        On Error GoTo Add_Error

20        mcolCounters.Add NewC

30        Set Add = NewC

40        Exit Function

Add_Error:

          Dim strES As String
          Dim intEL As Integer


50        intEL = Erl
60        strES = Err.Description
70        LogError "Counters", "Add", intEL, strES


End Function


Public Property Get Counters() As Counters

10        On Error GoTo Counters_Error

20        Set Counters = mCounters

30        Exit Property

Counters_Error:

          Dim strES As String
          Dim intEL As Integer


40        intEL = Erl
50        strES = Err.Description
60        LogError "Counters", "Counters", intEL, strES


End Property

Public Function Item(ByVal Index As Integer) _
       As Counter
Attribute Item.VB_UserMemId = 0

10        On Error GoTo Item_Error

20        On Error Resume Next

30        Set Item = mcolCounters(Index)

40        Exit Function

Item_Error:

          Dim strES As String
          Dim intEL As Integer


50        intEL = Erl
60        strES = Err.Description
70        LogError "Counters", "Item", intEL, strES


End Function


Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"

10        On Error GoTo NewEnum_Error

20        Set NewEnum = mcolCounters.[_NewEnum]

30        Exit Function

NewEnum_Error:

          Dim strES As String
          Dim intEL As Integer


40        intEL = Erl
50        strES = Err.Description
60        LogError "Counters", "NewEnum", intEL, strES


End Function



Public Sub RemoveItem(ByVal Index As Integer)

10        On Error GoTo RemoveItem_Error

20        mcolCounters.Remove Index

30        Exit Sub

RemoveItem_Error:

          Dim strES As String
          Dim intEL As Integer


40        intEL = Erl
50        strES = Err.Description
60        LogError "Counters", "RemoveItem", intEL, strES


End Sub


