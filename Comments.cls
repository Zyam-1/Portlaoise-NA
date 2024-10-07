VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Comments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Load(ByVal Tn As Long, ByVal SampleID As Double) As Comment

          Dim sql As String
          Dim tb As New Recordset
          Dim Cx As New Comment

10        On Error GoTo Load_Error

20        If SampleID = 0 Then Exit Function

30        sql = "SELECT * from Comments WHERE " & _
                "SampleID = " & SampleID
40        Set tb = New Recordset
50        RecOpenServer Tn, tb, sql

60        If tb.EOF Then
70            Set Load = Nothing
80        Else
90            Cx.SampleID = SampleID
100           Cx.Biochemistry = Trim$(tb!Biochemistry & "")
110           Cx.Haematology = Trim$(tb!Haematology & "")
120           Cx.Demographics = Trim$(tb!Demographic & "")
130           Cx.Coagulation = Trim$(tb!Coagulation & "")
140           Cx.Immunology = Trim$(tb!Immunology & "")
150           Cx.Endocrinology = Trim$(tb!Endocrinology & "")
160           Cx.BloodGas = Trim$(tb!BloodGas & "")
170           Cx.Semen = Trim$(tb!Semen & "")
180           Cx.MicroCS = tb!MicroCS & ""
190           Cx.MicroIdent = tb!MicroIdent & ""
200           Cx.MicroGeneral = tb!MicroGeneral & ""
210           Cx.MicroConsultant = tb!MicroConsultant & ""
220           Cx.Film = tb!Film & ""
230           Cx.Histology = tb!Histology & ""
240           Cx.Cytology = tb!Cytology & ""
250           Cx.CSFFluid = tb!CSFFluid & ""
260           Set Load = Cx
270       End If

280       Exit Function

Load_Error:

          Dim strES As String
          Dim intEL As Integer

290       intEL = Erl
300       strES = Err.Description
310       LogError "Comments", "Load", intEL, strES, sql

End Function


Public Sub Save(ByVal Cx As Comment)

          Dim sql As String
          Dim tb As New Recordset

10        On Error GoTo Save_Error

20        With Cx
30            If .SampleID = 0 Then Exit Sub

40            If Trim$(.Biochemistry & .Demographics & .Immunology & _
                       .Haematology & .Histology & .Cytology & .Coagulation & _
                       .Semen & .MicroIdent & .MicroConsultant & _
                       .MicroGeneral & .MicroCS & .BloodGas & .Endocrinology & .CSFFluid) = "" Then
50                sql = "DELETE from comments WHERE " & _
                        "sampleid = " & .SampleID
60                Cnxn(0).Execute sql
70            Else
80                sql = "SELECT * from Comments WHERE " & _
                        "SampleID = " & .SampleID
90                Set tb = New Recordset
100               RecOpenServer 0, tb, sql
110               If tb.EOF Then
120                   tb.AddNew
130                   tb!SampleID = .SampleID
140               End If
150               tb!Biochemistry = AddTicks(.Biochemistry)
160               tb!Demographic = AddTicks(.Demographics)
170               tb!Haematology = AddTicks(.Haematology)
180               tb!Coagulation = AddTicks(.Coagulation)
190               tb!Endocrinology = AddTicks(.Endocrinology)
200               tb!BloodGas = AddTicks(.BloodGas)
210               tb!Semen = AddTicks(.Semen)
220               tb!MicroIdent = AddTicks(.MicroIdent)
230               tb!MicroGeneral = AddTicks(.MicroGeneral)
240               tb!MicroConsultant = AddTicks(.MicroConsultant)
250               tb!MicroCS = AddTicks(.MicroCS)
260               tb!Immunology = AddTicks(.Immunology)
270               tb!Film = AddTicks(.Film)
280               tb!Cytology = AddTicks(.Cytology)
290               tb!Histology = AddTicks(.Histology)
300               tb!CSFFluid = AddTicks(.CSFFluid)
310               tb.Update
320           End If
330       End With

340       Exit Sub

Save_Error:

          Dim strES As String
          Dim intEL As Integer

350       intEL = Erl
360       strES = Err.Description
370       LogError "Comments", "Save", intEL, strES, sql

End Sub


