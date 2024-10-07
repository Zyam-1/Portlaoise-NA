VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmSystemFunctions 
   Caption         =   "System Functions"
   ClientHeight    =   5265
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9555
   LinkTopic       =   "Form1"
   ScaleHeight     =   5265
   ScaleWidth      =   9555
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkDisc 
      Caption         =   "Repeat"
      Height          =   255
      Index           =   5
      Left            =   2730
      TabIndex        =   8
      Top             =   1620
      Width           =   915
   End
   Begin VB.CheckBox chkDisc 
      Caption         =   "Imm"
      Height          =   255
      Index           =   4
      Left            =   2730
      TabIndex        =   7
      Top             =   1320
      Width           =   795
   End
   Begin VB.CheckBox chkDisc 
      Caption         =   "Repeat"
      Height          =   255
      Index           =   3
      Left            =   1695
      TabIndex        =   6
      Top             =   1620
      Width           =   915
   End
   Begin VB.CheckBox chkDisc 
      Caption         =   "End"
      Height          =   255
      Index           =   2
      Left            =   1695
      TabIndex        =   5
      Top             =   1320
      Width           =   795
   End
   Begin VB.CheckBox chkDisc 
      Caption         =   "Repeat"
      Height          =   255
      Index           =   1
      Left            =   660
      TabIndex        =   4
      Top             =   1620
      Width           =   915
   End
   Begin VB.CheckBox chkDisc 
      Caption         =   "Bio"
      Height          =   255
      Index           =   0
      Left            =   660
      TabIndex        =   3
      Top             =   1320
      Width           =   795
   End
   Begin MSComCtl2.DTPicker dtFrom 
      Height          =   315
      Left            =   480
      TabIndex        =   1
      Top             =   720
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   556
      _Version        =   393216
      Format          =   166461441
      CurrentDate     =   41898
   End
   Begin VB.CommandButton cmdUpdateDefIndexBio 
      Caption         =   "Update DefIndex"
      Height          =   555
      Left            =   300
      TabIndex        =   0
      Top             =   2220
      Width           =   3915
   End
   Begin MSComCtl2.DTPicker dtTo 
      Height          =   315
      Left            =   2460
      TabIndex        =   2
      Top             =   720
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   556
      _Version        =   393216
      Format          =   166461441
      CurrentDate     =   41898
   End
End
Attribute VB_Name = "frmSystemFunctions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Procedure : cmdUpdateDefIndexBio_Click
' Author    : Babar Shahzad
' Date      : 06/09/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmdUpdateDefIndexBio_Click()

On Error GoTo cmdUpdateDefIndexBio_Click_Error

Dim tb As Recordset
Dim tbDefIndex As Recordset
Dim tbTD As Recordset
Dim sql As String
Dim br As BIEResult
Dim BRs As New BIEResults

'get list of all bioresults joined with demographics
sql = "SELECT Distinct SampleID FROM Demographics WHERE  SampleID < " & SysOptSemenOffset(0) & " " & _
        "AND RunDate between '" & Format(dtFrom, "dd/MMM/yyyy") & "' and '" & Format(dtTo, "dd/MMM/yyyy") & "'"

'sql = "select distinct sampleid from BioResults where SampleId in ( " & _
'        "select distinct SampleID  from Demographics " & _
'        "where DateTimeDemographics between '01/jan/2014 00:00:01' and '31/dec/2014 23:59:59' " & _
'        "and SampleID < 100000000000) and DefIndex is  null"
Set tb = New Recordset
RecOpenServer 0, tb, sql


While Not tb.EOF
    cmdUpdateDefIndexBio.Caption = "Updating SampleID: " & tb!SampleID & " (Bio)"
    If chkDisc(0).Value = 1 Then
        Set BRs = LoadResults("Bio", tb!SampleID, "Results", gDONTCARE, gDONTCARE, 0, "Default", "")
    End If
    If chkDisc(1).Value = 1 Then
        Set BRs = LoadResults("Bio", tb!SampleID, "Repeats", gDONTCARE, gDONTCARE, 0, "Default", "")
    End If
        
    cmdUpdateDefIndexBio.Caption = "Updating SampleID: " & tb!SampleID & " (End)"
    If chkDisc(2).Value = 1 Then
        Set BRs = LoadResults("End", tb!SampleID, "Results", gDONTCARE, gDONTCARE, 0, "Default", "")
    End If
    If chkDisc(3).Value = 1 Then
        Set BRs = LoadResults("End", tb!SampleID, "Repeats", gDONTCARE, gDONTCARE, 0, "Default", "")
    End If

    cmdUpdateDefIndexBio.Caption = "Updating SampleID: " & tb!SampleID & " (Imm)"
    If chkDisc(4).Value = 1 Then
        Set BRs = LoadResults("Imm", tb!SampleID, "Results", gDONTCARE, gDONTCARE, 0, "Default", "")
    End If
    If chkDisc(5).Value = 1 Then
        Set BRs = LoadResults("Imm", tb!SampleID, "Repeats", gDONTCARE, gDONTCARE, 0, "Default", "")
    End If

    tb.MoveNext
    Me.Refresh
Wend

cmdUpdateDefIndexBio.Caption = "Update Def Index"

Exit Sub

cmdUpdateDefIndexBio_Click_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "frmSystemFunctions", "cmdUpdateDefIndexBio_Click", intEL, strES, sql

End Sub

Public Function LoadResults(ByVal Discipline As String, _
                     ByVal SampleID As String, _
                     ByVal ResultOrRepeat As String, _
                     ByVal v As Long, _
                     ByVal P As Long, _
                     ByVal Cn As Long, _
                     ByVal Cat As String, _
                     ByVal Rundate As String) _
                     As BIEResults
'Discipline is either "Bio", "Imm" or "End"

Dim BRs As New BIEResults
Dim br As BIEResult
Dim tb As New Recordset
Dim sql As String
Dim Dob As String
Dim DaysOld As Long
Dim SELECTNormalRange As String
Dim SELECTFlagRange As String
Dim TableName As String
Dim sex As String
Dim tbRange As Recordset
Dim tbNewIDX  As Recordset

On Error GoTo Load_Error

DaysOld = 9125

If SampleID = "" Then Exit Function

If UCase(Discipline) = "BIO" Or UCase(Discipline) = "BGA" Then Cat = ""

TableName = Discipline & ResultOrRepeat

sql = "SELECT DoB, Sex,rundate from Demographics WHERE " & _
      "SampleID = '" & Val(SampleID) & "'"
Set tb = Cnxn(Cn).Execute(sql)
If Not tb.EOF Then
    If IsDate(tb!Dob) Then
        Dob = Format$(tb!Dob, "dd/mmm/yyyy")
        DaysOld = DateDiff("d", Dob, tb!Rundate)
        If DaysOld = 0 Then DaysOld = 1
    End If
    sex = Left$(UCase$(Trim$(tb!sex & "")), 1)
    Select Case Left$(UCase$(Trim$(tb!sex & "")), 1)
    Case "M": SELECTNormalRange = " MaleLow as Low, MaleHigh as High, "
        SELECTFlagRange = " FlagMaleLow as FlagLow, FlagMaleHigh as FlagHigh, "
    Case "F": SELECTNormalRange = " FemaleLow as Low, FemaleHigh as High, "
        SELECTFlagRange = " FlagFemaleLow as FlagLow, FlagFemaleHigh as FlagHigh, "
    Case Else: SELECTNormalRange = " FemaleLow as Low, MaleHigh as High, "
        SELECTFlagRange = " FlagFemaleLow as FlagLow, FlagMaleHigh as FlagHigh, "
    End Select
Else
    SELECTNormalRange = " FemaleLow as Low, MaleHigh as High, "
    SELECTFlagRange = " FlagFemaleLow as FlagLow, FlagMaleHigh as FlagHigh, "
End If


sql = "SELECT COALESCE(R.DefIndex, 0) DefIndex, COALESCE(X.NormalLow, 0) Low, " & _
      "COALESCE(X.NormalHigh, 9999) High, COALESCE(X.FlagLow, 0) FlagLow, COALESCE(X.FlagHigh, 9999) FlagHigh, " & _
      "COALESCE(X.PlausibleLow, 0) PlausibleLow, COALESCE(X.PlausibleHigh, 9999) PlausibleHigh, " & _
      "LongName, ShortName, DoDelta, DeltaLimit, CheckTime, Printable, " & _
      "DP, PrintPriority, " & _
      "R.SampleID, R.Code, R.Result, " & _
      "COALESCE(R.Valid, 0) AS Valid, COALESCE(R.Printed, 0) Printed, " & _
      "R.RunTime, R.RunDate, R.Operator, R.Flags, R.Units, " & _
      "R.SampleType, R.Analyser, R.Faxed, R.Authorised, " & _
      "R.Comment AS Comment, R.PC "
If UCase$(Discipline) = "IMM" Then
    sql = sql & ", prnrr "
End If
sql = sql & "FROM " & TableName & " R JOIN " & Discipline & "TestDefinitions2 D ON R.Code = D.Code " & _
      "LEFT JOIN " & Discipline & "DefIndex X ON R.DefIndex = X.DefIndex " & _
      "WHERE " & _
      "SampleID = '" & SampleID & "' AND D.category = '" & Cat & "' " & _
      "AND R.Code = D.Code " & _
      "AND AgeFromDays <= " & DaysOld & " " & _
      "AND AgeToDays >= " & DaysOld & " "




sql = sql & "and R.SampleType = D.SampleType "
If P = gNOTPRINTED And v = gNOTVALID Then
    sql = sql & "and Printed = 0 and Valid = 0 "
ElseIf P = gNOTPRINTED And v = gVALID Then
    sql = sql & "and Printed = 0 and Valid = 1 "
ElseIf P = gNOTPRINTED And v = gDONTCARE Then
    sql = sql & "and Printed = 0 "
ElseIf P = gPRINTED And v = gNOTVALID Then
    sql = sql & "and Printed = 1 and Valid = 0 "
ElseIf P = gPRINTED And v = gVALID Then
    sql = sql & "and Printed = 1 and Valid = 1 "
ElseIf P = gPRINTED And v = gDONTCARE Then
    sql = sql & "and Printed = 1 "
ElseIf P = gDONTCARE And v = gNOTVALID Then
    sql = sql & "and Valid = 0 "
ElseIf P = gDONTCARE And v = gVALID Then
    sql = sql & "and Valid = 1 "
End If
sql = sql & "Order by PrintPriority asc"
Set tb = New Recordset
RecOpenServer Cn, tb, sql    '  RecOpenClient 0,tb, Sql
Do While Not tb.EOF
    Set br = New BIEResult
    With br
        .SampleID = Trim(tb!SampleID & "")
        .DefIndex = tb!DefIndex
        .Code = Trim(tb!Code & "")
        .Result = Trim(tb!Result & "")
        .Operator = Trim(tb!Operator & "")
        .Rundate = Format$(tb!Rundate, "dd/mm/yyyy")
        .RunTime = Format$(tb!RunTime, "dd/mm/yyyy hh:mm")
        .Units = Trim(tb!Units & "")
        If Trim(tb!Printed & "") <> "" Then .Printed = IIf(tb!Printed, True, False)
        If Trim(tb!Valid & "") <> "" Then .Valid = IIf(tb!Valid, True, False)
        .Flags = Trim(tb!Flags & "")
        .SampleType = Trim(tb!SampleType & "")
        .Low = IIf(IsNull(tb!Low), 0, tb!Low)
        .FlagLow = IIf(IsNull(tb!FlagLow), 0, tb!FlagLow)
        .PlausibleLow = IIf(IsNull(tb!PlausibleLow), 0, tb!PlausibleLow)
        .High = IIf(IsNull(tb!High), 9999, tb!High)
        .FlagHigh = IIf(IsNull(tb!FlagHigh), 9999, tb!FlagHigh)
        .PlausibleHigh = IIf(IsNull(tb!PlausibleHigh), 99999, tb!PlausibleHigh)
        .Printformat = IIf(IsNull(tb!DP), 0, tb!DP)
        .ShortName = Trim(tb!ShortName & "")
        .LongName = Trim(tb!LongName & "")
        If tb!DoDelta & "" <> "" Then .DoDelta = tb!DoDelta Else .DoDelta = False
        .DeltaLimit = IIf(IsNull(tb!DeltaLimit), 9999, tb!DeltaLimit)
        .Analyser = Trim(tb!Analyser & "")
        If Discipline = "Imm" Then
            .PrnRR = IIf(IsNull(tb!PrnRR), True, tb!PrnRR)
        End If
        .Comment = Trim(tb!Comment & "")
        .Pc = Trim(tb!Pc & "")
        If IsNull(tb!CheckTime) Then
            .CheckTime = 1
        Else
            .CheckTime = tb!CheckTime
        End If
        .Printable = tb!Printable
        '   .NormalLow = tb!NormalLow
        '   .NormalHigh = tb!NormalHigh
        '   .NormalUsed = tb!NormalUsed

        If .DefIndex = 0 Then
            If Dob <> "" And sex <> "" Then
                sql = "SELECT " & _
                      SELECTNormalRange & SELECTFlagRange & _
                      "COALESCE(PlausibleLow, 0) PlausibleLow, " & _
                      "COALESCE(PlausibleHigh, 9999) PlausibleHigh " & _
                      "FROM " & Discipline & "TestDefinitions2  " & _
                      "WHERE category = '" & Cat & "' " & _
                      "AND Code = '" & .Code & "' " & _
                      "AND AgeFromDays <= " & DaysOld & " " & _
                      "AND AgeToDays >= " & DaysOld & " "
                Set tbRange = New Recordset
                RecOpenServer 0, tbRange, sql
                If Not tbRange.EOF Then
                    .Low = tbRange!Low
                    .High = tbRange!High
                    .FlagLow = tbRange!FlagLow
                    .FlagHigh = tbRange!FlagHigh
                    .PlausibleLow = tbRange!PlausibleLow
                    .PlausibleHigh = tbRange!PlausibleHigh

                    sql = "SELECT * FROM " & Discipline & "DefIndex " & _
                          "WHERE NormalLow = '" & .Low & "' " & _
                          "AND NormalHigh = '" & .High & "' " & _
                          "AND FlagLow = '" & .FlagLow & "' " & _
                          "AND FlagHigh = '" & .FlagHigh & "' " & _
                          "AND PlausibleLow = '" & .PlausibleLow & "' " & _
                          "AND PlausibleHigh = '" & .PlausibleHigh & "' "

                    Set tbNewIDX = New Recordset
                    RecOpenClient 0, tbNewIDX, sql
                    If Not tbNewIDX.EOF Then
                        .DefIndex = tbNewIDX!DefIndex
                    Else

                        sql = "INSERT INTO " & Discipline & "DefIndex " & _
                              "( NormalLow, NormalHigh, FlagLow, FlagHigh, " & _
                              "  PlausibleLow, PlausibleHigh, AutoValLow, AutoValHigh) " & _
                              "VALUES ( " & _
                              .Low & ", " & .High & ", " & .FlagLow & ", " & .FlagHigh & ", " & _
                              .PlausibleLow & ", " & .PlausibleHigh & ", 0,9999) "
                        Cnxn(0).Execute sql

                        sql = "SELECT MAX(DefIndex) NewIndex FROM " & Discipline & "DefIndex"
                        Set tbNewIDX = New Recordset
                        RecOpenClient 0, tbNewIDX, sql
                        .DefIndex = tbNewIDX!NewIndex

                    End If

                    sql = "UPDATE " & TableName & " " & _
                          "SET DefIndex = '" & .DefIndex & "' " & _
                          "WHERE SampleID = '" & .SampleID & "' " & _
                          "AND Code = '" & .Code & "'"
                    Cnxn(0).Execute sql

                End If
            End If
        End If
        BRs.Add br
    End With
    tb.MoveNext
Loop

If BRs.Count <> 0 Then
    Set LoadResults = BRs
Else
    Set LoadResults = Nothing
End If
Set br = Nothing
Set BRs = Nothing


Exit Function



Exit Function

Load_Error:

Dim strES As String
Dim intEL As Integer


intEL = Erl
strES = Err.Description
LogError "BIEResults", "Load", intEL, strES, sql


End Function


Private Sub Form_Load()
dtFrom.Value = Now
dtTo.Value = Now
End Sub
