VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Counterbs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private mcolCounterbs As New Collection

Private mCounterbs As New Counterbs


Public Sub Clear()

          Dim n As Long

10        On Error GoTo Clear_Error

20        For n = 1 To mcolCounterbs.Count
30            mcolCounterbs.Remove 1
40        Next

50        Exit Sub

Clear_Error:

          Dim strES As String
          Dim intEL As Integer


60        intEL = Erl
70        strES = Err.Description
80        LogError "Counterbs", "Clear", intEL, strES


End Sub

Public Property Get Counterbs() As Counterbs

10        On Error GoTo Counterbs_Error

20        Set Counterbs = mCounterbs

30        Exit Property

Counterbs_Error:

          Dim strES As String
          Dim intEL As Integer


40        intEL = Erl
50        strES = Err.Description
60        LogError "Counterbs", "Counterbs", intEL, strES


End Property


Public Sub Add(ByVal Clinician As String, _
               ByVal Ward As String, _
               ByVal Which As String)

          Dim c As Counterb
          Dim Found As Boolean

10        On Error GoTo Add_Error

20        Clinician = Trim(UCase(Clinician))
30        Ward = Trim(UCase(Ward))

40        For Each c In mcolCounterbs
50            With c
60                If Trim(UCase(.Clinician)) = Clinician And Trim(UCase(.Ward)) = Ward Then
70                    Select Case Which
                      Case "E": .ESRCounter = .ESRCounter + 1
80                    Case "F": .FBCCounter = .FBCCounter + 1
90                    Case "S": .SampleCounter = .SampleCounter + 1
100                   Case "T": .TestCounter = .TestCounter + 1
110                   End Select
120                   Found = True
130                   Exit For
140               End If
150           End With
160       Next

170       If Not Found Then
180           Set c = New Counterb
190           c.Clinician = Clinician
200           c.Ward = Ward
210           Select Case Which
              Case "E": c.ESRCounter = 1
220           Case "F": c.FBCCounter = 1
230           Case "S": c.SampleCounter = 1
240           Case "T": c.TestCounter = 1
250           End Select
260           mcolCounterbs.Add c
270       End If

280       Set c = Nothing

290       Exit Sub

Add_Error:

          Dim strES As String
          Dim intEL As Integer


300       intEL = Erl
310       strES = Err.Description
320       LogError "Counterbs", "Add", intEL, strES


End Sub

Public Function Item(ByVal X As Long) As Counterb
Attribute Item.VB_UserMemId = 0

10        On Error GoTo Item_Error

20        Set Item = mcolCounterbs.Item(CLng(X))

30        Exit Function

Item_Error:

          Dim strES As String
          Dim intEL As Integer


40        intEL = Erl
50        strES = Err.Description
60        LogError "Counterbs", "Item", intEL, strES


End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"

10        On Error GoTo NewEnum_Error

20        Set NewEnum = mcolCounterbs.[_NewEnum]

30        Exit Function

NewEnum_Error:

          Dim strES As String
          Dim intEL As Integer


40        intEL = Erl
50        strES = Err.Description
60        LogError "Counterbs", "NewEnum", intEL, strES


End Function

Public Function Count() As Long

10        On Error GoTo Count_Error

20        Count = mcolCounterbs.Count

30        Exit Function

Count_Error:

          Dim strES As String
          Dim intEL As Integer


40        intEL = Erl
50        strES = Err.Description
60        LogError "Counterbs", "Count", intEL, strES


End Function



