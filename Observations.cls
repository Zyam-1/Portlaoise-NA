VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Observations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mcolObservations As New Collection
Private mObservations As New Observations

Public Sub Clear()

          Dim n As Long

10        For n = 1 To mcolObservations.Count
20            mcolObservations.Remove 1
30        Next

End Sub

Public Property Get Observations() As Observations

10        Set Observations = mObservations

End Property


Public Function Item(ByVal X As Long) _
       As Observation

10        Set Item = mcolObservations(X)

End Function



Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"

10        Set NewEnum = mcolObservations.[_NewEnum]

End Function




Public Sub Remove(ByVal Index As Integer)

10        mcolObservations.Remove (Index)

End Sub


Public Function Load(ByVal SampleID As Double, ParamArray Disc()) As Observations
      'Paramarray is one or a combination of :-
      'Demographic Biochemistry Haematology Coagulation Immunology BloodGas
      'Semen Microcs MicroIdent MicroGeneral MicroConsultant Film
      'Endocrinology Histology Cytology CSFFluid
      'ImmunologyA ImmunologyB ImmunologyC

          Dim Obs As New Observations
          Dim Ob As Observation
          Dim sql As String
          Dim tb As New Recordset
          Dim n As Integer
          Dim s As String

10        On Error GoTo Load_Error

20        sql = "SELECT * FROM Observations WHERE " & _
                "SampleID = " & SampleID & " AND ("
30        s = ""
40        For n = LBound(Disc) To UBound(Disc)
50            s = s & "Discipline = '" & Disc(n) & "' OR "
60        Next
70        s = Left$(s, Len(s) - 3) & ")"
80        sql = sql & s

90        Set tb = New Recordset
100       RecOpenClient 0, tb, sql

110       Do While Not tb.EOF
120           Set Ob = New Observation
130           With Ob
140               .SampleID = tb!SampleID & ""
150               .Discipline = tb!Discipline & ""
160               .Comment = RemoveLeadingCrLf(tb!Comment & "")
170               .DateTimeOfRecord = Format$(tb!DateTimeOfRecord, "dd/MM/yyyy HH:nn:ss")
180               .Username = tb!Username & ""
190               If .Comment <> "" Then
200                   Obs.Add Ob
210               End If
220           End With
230           tb.MoveNext
240       Loop

250       If Obs.Count <> 0 Then
260           Set Load = Obs
270       Else
280           Set Load = Nothing
290       End If

300       Exit Function

Load_Error:

          Dim strES As String
          Dim intEL As Integer

310       intEL = Erl
320       strES = Err.Description
330       LogError "Observations", "Load", intEL, strES, sql


End Function
Public Function Save(ByVal SampleID As String, ByVal OverWrite As Boolean, ParamArray Disc()) As Observations

      'OverWrite = True : Any existing comment is replaced.
      'OverWrite = False : New comment is concatenated

      'Paramarray is one or a combination of :-
      'Demographic Biochemistry Haematology Coagulation Immunology BloodGas
      'Semen Microcs MicroIdent MicroGeneral MicroConsultant Film
      'Endocrinology Histology Cytology CSFFluid
      'ImmunologyA ImmunologyB ImmunologyC
      'eg "Haematology","Haem Comment","Biochemistry","Bio Comment" etc

          Dim sql As String
          Dim n As Integer
          Dim Discipline As String
          Dim Comment As String

10        On Error GoTo Save_Error

20        For n = LBound(Disc) To UBound(Disc) Step 2
30            Discipline = Disc(n)
40            Comment = AddTicks(Disc(n + 1))
50            Comment = RemoveLeadingCrLf(Comment)
60            If Comment = "" Then
70                sql = "DELETE FROM Observations " & _
                        "WHERE SampleID = '" & SampleID & "' " & _
                        "AND Discipline = '" & Discipline & "'"
80            Else
90                sql = "IF EXISTS (SELECT * FROM Observations " & _
                        "WHERE SampleID = '" & SampleID & "' " & _
                        "AND Discipline = '" & Discipline & "') " & _
                        "  UPDATE Observations "
100               If OverWrite Then
110                   sql = sql & "  SET Comment = '" & Comment & "' "
120               Else
130                   sql = sql & "  SET Comment = Comment + ' " & Comment & "' "
140               End If
150               sql = sql & "  WHERE SampleID = '" & SampleID & "' " & _
                        "  AND Discipline = '" & Discipline & "' " & _
                        "ELSE " & _
                        "  INSERT INTO Observations " & _
                        "  (SampleID, Discipline, Comment, UserName ) " & _
                        "  VALUES " & _
                        "  ('" & SampleID & "', " & _
                        "   '" & Discipline & "', " & _
                        "   '" & Comment & "', " & _
                        "   '" & AddTicks(Username) & "')"
160           End If
170           Cnxn(0).Execute sql
180       Next

190       Exit Function

Save_Error:

          Dim strES As String
          Dim intEL As Integer

200       intEL = Erl
210       strES = Err.Description
220       LogError "Observations", "Save", intEL, strES, sql


End Function
Public Function Count() As Long

10        Count = mcolObservations.Count

20        Exit Function

End Function



Public Function Add(Ob As Observation) As Observation
Attribute Add.VB_MemberFlags = "40"

10        mcolObservations.Add Ob

20        Set Add = Ob

End Function


