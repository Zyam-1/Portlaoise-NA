VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmBatECPS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetAcquire - Batch Culture Entry"
   ClientHeight    =   7650
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8580
   ClipControls    =   0   'False
   Icon            =   "frmBatchEPCS.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   8580
   Begin VB.CommandButton cmdSet 
      BackColor       =   &H0000FFFF&
      Caption         =   "E Coli 0157    &Set All Negative"
      Height          =   315
      Index           =   0
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   210
      Width           =   3315
   End
   Begin VB.CommandButton cmdSave 
      Cancel          =   -1  'True
      Caption         =   "Save"
      Height          =   900
      Left            =   7260
      Picture         =   "frmBatchEPCS.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5490
      Width           =   1100
   End
   Begin VB.CommandButton cmdSet 
      BackColor       =   &H00FFFF00&
      Caption         =   "Campylobacter Set &Negative "
      Height          =   315
      Index           =   1
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   210
      Width           =   3315
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Exit"
      Height          =   900
      Left            =   7260
      Picture         =   "frmBatchEPCS.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6525
      Width           =   1100
   End
   Begin MSFlexGridLib.MSFlexGrid grdCul 
      Height          =   6885
      Index           =   1
      Left            =   3720
      TabIndex        =   0
      Top             =   540
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   12144
      _Version        =   393216
      BackColor       =   -2147483644
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483643
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      ScrollBars      =   2
      FormatString    =   "<Sample ID   |Result                               "
   End
   Begin MSFlexGridLib.MSFlexGrid grdCul 
      Height          =   6885
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   540
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   12144
      _Version        =   393216
      BackColor       =   -2147483644
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483643
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      ScrollBars      =   2
      FormatString    =   "<Sample ID   |Result                               "
   End
End
Attribute VB_Name = "frmBatECPS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

10        Unload Me

End Sub

Private Sub cmdSave_Click()
          Dim sql As String
          Dim tb As Recordset
          Dim Num As Long


10        On Error GoTo cmdSave_Click_Error

20        For Num = 0 To grdCul(0).Rows - 1
30            If grdCul(0).TextMatrix(Num, 1) <> "" Then
40                sql = "SELECT * from faeces WHERE " & _
                        "SampleID = " & SysOptMicroOffset(0) + Val(grdCul(0).TextMatrix(Num, 0))
50                Set tb = New Recordset
60                RecOpenServer 0, tb, sql
70                If tb.EOF Then
80                    tb.AddNew
90                    tb!SampleID = SysOptMicroOffset(0) + Val(grdCul(0).TextMatrix(Num, 0))
100               End If
110               tb!pc0157 = Left(grdCul(0).TextMatrix(Num, 1), 1)
120               If Left(grdCul(0).TextMatrix(Num, 1), 1) = "N" Then
130                   tb!pc0157report = "E Coli 0157 Not Isolated"
140               End If
150               tb.Update
160           End If
170       Next

180       For Num = 0 To grdCul(1).Rows - 1
190           If grdCul(1).TextMatrix(Num, 1) <> "" Then
200               sql = "SELECT * from faeces WHERE " & _
                        "sampleid = " & SysOptMicroOffset(0) + Val(grdCul(1).TextMatrix(Num, 0))
210               Set tb = New Recordset
220               RecOpenServer 0, tb, sql
230               If tb.EOF Then
240                   tb.AddNew
250                   tb!SampleID = SysOptMicroOffset(0) + Val(grdCul(1).TextMatrix(Num, 0))
260               End If
270               tb!camp = Left(grdCul(1).TextMatrix(Num, 1), 1)
280               If Left(grdCul(1).TextMatrix(Num, 1), 1) = "N" Then
290                   tb!CampCulture = "Not Isolated"
300               End If
310               tb.Update
320           End If
330       Next

340       For Num = 0 To grdCul(2).Rows - 1
350           If grdCul(2).TextMatrix(Num, 1) <> "" Then
360               sql = "SELECT * from faeces WHERE " & _
                        "sampleid = '" & SysOptMicroOffset(0) + Val(grdCul(2).TextMatrix(Num, 0))
370               Set tb = New Recordset
380               RecOpenServer 0, tb, sql
390               If tb.EOF Then
400                   tb.AddNew
410                   tb!SampleID = SysOptMicroOffset(0) + Val(grdCul(2).TextMatrix(Num, 0))
420               End If
430               If grdCul(2).TextMatrix(Num, 1) = "Done" Then
440                   tb!pcdone = 1
450               Else
460                   tb!pcdone = 0
470               End If
480               tb!Pc = Left(grdCul(2).TextMatrix(Num, 2), 1)
490               tb.Update
500           End If
510       Next

520       For Num = 0 To grdCul(3).Rows - 1
530           If grdCul(3).TextMatrix(Num, 1) <> "" Then
540               sql = "SELECT * from faeces WHERE sampleid = " & SysOptMicroOffset(0) + Val(grdCul(3).TextMatrix(Num, 0))
550               Set tb = New Recordset
560               RecOpenServer 0, tb, sql
570               If tb.EOF Then
580                   tb.AddNew
590                   tb!SampleID = SysOptMicroOffset(0) + Val(grdCul(3).TextMatrix(Num, 0))
600               End If
610               If grdCul(3).TextMatrix(Num, 1) = "Done" Then
620                   tb!selenitedone = 1
630               Else
640                   tb!selenitedone = 0
650               End If
660               tb!selenite = Left(grdCul(3).TextMatrix(Num, 2), 1)
670               tb.Update
680           End If
690       Next

700       Unload Me

710       Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

720       intEL = Erl
730       strES = Err.Description
740       LogError "frmBatECPS", "cmdsave_Click", intEL, strES, sql

End Sub

Private Sub cmdSet_Click(Index As Integer)

          Dim Num As Long
          Dim NumR As Long

10        On Error GoTo cmdSet_Click_Error

20        If grdCul(Index).Rows = 1 And grdCul(Index).TextMatrix(0, 0) = "" Then Exit Sub

30        NumR = grdCul(Index).Rows - 1

40        For Num = 1 To NumR
50            grdCul(Index).Col = 1
60            grdCul(Index).Row = Num
70            grdCul(Index) = ""
80            grdCul_Click (Index)
90        Next

100       Exit Sub

cmdSet_Click_Error:

          Dim strES As String
          Dim intEL As Integer

110       intEL = Erl
120       strES = Err.Description
130       LogError "frmBatECPS", "cmdSet_Click", intEL, strES


End Sub

Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        load0157
30        LoadCamp

40        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmBatECPS", "Form_Load", intEL, strES

End Sub

Private Sub grdCul_Click(Index As Integer)


10        On Error GoTo grdCul_Click_Error

20        If grdCul(Index).Rows = 1 Then Exit Sub

30        If grdCul(Index).MouseCol = 0 Then Exit Sub

40        If Index < 2 Then
50            If Left(grdCul(Index), 3) = "Neg" Then
60                grdCul(Index) = "Pos"
70            ElseIf Left(grdCul(Index), 3) = "Pos" Then
80                grdCul(Index) = ""
90            ElseIf grdCul(Index) = "" Then
100               grdCul(Index) = "Neg"
110           End If
120       Else
130           If grdCul(Index).MouseCol = 2 Then
140               If Left(grdCul(Index), 3) = "Neg" Then
150                   grdCul(Index) = "Pos"
160               ElseIf Left(grdCul(Index), 3) = "Pos" Then
170                   grdCul(Index) = ""
180               ElseIf grdCul(Index) = "" Then
190                   grdCul(Index) = "Neg"
200               End If
210           Else
220               If grdCul(Index) = "Done" Then
230                   grdCul(Index) = ""
240               ElseIf grdCul(Index) = "" Then
250                   grdCul(Index) = "Done"
260               End If
270           End If
280       End If

290       Exit Sub

grdCul_Click_Error:

          Dim strES As String
          Dim intEL As Integer

300       intEL = Erl
310       strES = Err.Description
320       LogError "frmBatECPS", "grdCul_Click", intEL, strES


End Sub

Private Sub load0157()

          Dim s As String
          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo load0157_Error

20        With grdCul(0)
30            .Rows = 1
40            .ColWidth(0) = 1000
50            .ColWidth(1) = 1000
60            .Rows = 1
70        End With

80        sql = "Select R.SampleID, F.pc0157, COALESCE(L.Valid, 0) As Valid From FaecalRequests R " & _
                "Left Join (Select * From PrintValidLog Where Department = 'D') L On R.SampleID = L.SampleID " & _
                "Left Join Faeces F On R.SampleID = F.SampleID " & _
                "Where R.Coli0157 = 1 And COALESCE(L.Valid, 0) = 0 " & _
                "And COALESCE(F.pc0157report, '') = '' " & _
                "Order By R.SampleID"

          'sql = "SELECT R.Coli0157, R.SampleID " & _
           '      "FROM FaecalRequests R JOIN Faeces F ON " & _
           '      "R.SampleID = F.SampleID WHERE " & _
           '      "R.Coli0157 = 1 " & _
           '      "AND COALESCE(F.Valid, 0) = 0 " & _
           '      "ORDER BY R.SampleID"

90        Set tb = New Recordset
100       RecOpenServer 0, tb, sql

110       Do While Not tb.EOF
120           s = tb!SampleID - SysOptMicroOffset(0) & vbTab
130           If tb!pc0157 = "P" Then
140               s = s & "Positive"
150           ElseIf tb!pc0157 = "N" Then
160               s = s & "Negative"
170           Else
180               s = s & ""
190           End If
200           grdCul(0).AddItem s


210           tb.MoveNext
220       Loop


230       Exit Sub

load0157_Error:

          Dim strES As String
          Dim intEL As Integer

240       intEL = Erl
250       strES = Err.Description
260       LogError "frmBatECPS", "load0157", intEL, strES

End Sub

Private Sub LoadCamp()

          Dim s As String
          Dim addit As Long
          Dim sql As String
          Dim sn As Recordset
          Dim tb As Recordset

10        On Error GoTo LoadCamp_Error

20        With grdCul(1)
30            .Rows = 1
40            .ColWidth(0) = 1000
50            .ColWidth(1) = 1000
60        End With

70        sql = "SELECT R.SampleID " & _
                "FROM FaecalRequests R JOIN Faeces F ON " & _
                "R.SampleID = F.SampleID " & _
                "WHERE " & _
                "COALESCE(F.Valid, 0) = 0 " & _
                "AND R.Campylobacter = 1 " & _
                "ORDER BY R.SampleID"

80        Set sn = New Recordset
90        RecOpenServer 0, sn, sql

100       Do While Not sn.EOF
110           sql = "SELECT * from faeces WHERE sampleid = " & sn!SampleID
120           Set tb = New Recordset
130           RecOpenServer 0, tb, sql
140           addit = False
150           If tb.EOF Then
160               grdCul(1).AddItem sn!SampleID - SysOptMicroOffset(0)
170           Else
180               If Trim(tb!camp) & "" = "" Then
190                   addit = True
200               End If
210           End If
220           If addit Then
230               s = sn!SampleID - SysOptMicroOffset(0) & vbTab
240               If tb!camp = "P" Then
250                   s = s & "Positive"
260               ElseIf tb!camp = "N" Then
270                   s = s & "Negative"
280               Else
290                   s = s & ""
300               End If
310               grdCul(1).AddItem s
320           End If
330           sn.MoveNext
340       Loop

350       Exit Sub

LoadCamp_Error:

          Dim strES As String
          Dim intEL As Integer

360       intEL = Erl
370       strES = Err.Description
380       LogError "frmBatECPS", "LoadCamp", intEL, strES

End Sub

