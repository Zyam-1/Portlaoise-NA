VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmViewBioRepeat 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Biochemistry Repeats"
   ClientHeight    =   5685
   ClientLeft      =   4770
   ClientTop       =   2475
   ClientWidth     =   5175
   Icon            =   "frmViewBioRepeat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton bDelete 
      Caption         =   "&Delete All Repeats"
      Height          =   945
      Left            =   3645
      Picture         =   "frmViewBioRepeat.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton bTransfer 
      Caption         =   "Copy to Main File"
      Height          =   855
      Left            =   3645
      Picture         =   "frmViewBioRepeat.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   495
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton bCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   765
      Left            =   3645
      Picture         =   "frmViewBioRepeat.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4680
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   4965
      Left            =   120
      TabIndex        =   0
      Top             =   510
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   8758
      _Version        =   393216
      Cols            =   3
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      FormatString    =   "<Test                  |<Result  |<Units    "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tests in RED will be Transfered"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   210
      Width           =   2445
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Highlight Tests to be Transferred"
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   3645
      TabIndex        =   3
      Top             =   540
      Width           =   1215
   End
End
Attribute VB_Name = "frmViewBioRepeat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Activated As Boolean



Private Sub FillG()

          Dim s As String
          Dim bioReps As New BIEResults
          Dim BRs As BIEResults
          Dim br As BIEResult
          Dim n As Long

10        On Error GoTo FillG_Error

20        Set BRs = bioReps.Load("Bio", frmEditAll.txtSampleID, "Repeats", gDONTCARE, gDONTCARE, 0, "", frmEditAll.dtRunDate)

30        g.Rows = 2
40        g.AddItem ""
50        g.RemoveItem 1

60        If Not BRs Is Nothing Then
70            For Each br In BRs
80                For n = 1 To frmEditAll.gBio.Rows - 1
90                    If frmEditAll.gBio.TextMatrix(n, 0) = br.ShortName Then
100                       s = br.ShortName & vbTab & Trim(br.Result) & vbTab & br.Units
110                       g.AddItem s
120                   End If
130               Next
140           Next
150       End If

160       If g.Rows > 2 Then
170           g.RemoveItem 1
180       End If

190       Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer
          MsgBox (strES & " " & intEL)



200       intEL = Erl
210       strES = Err.Description
220       LogError "frmViewBioRepeat", "FillG", intEL, strES


End Sub

Private Sub bcancel_Click()

10        Unload Me

End Sub


Private Sub bDELETE_Click()

          Dim sql As String

10        On Error GoTo bDELETE_Click_Error

20        If iMsg("DELETE All Repeats?" & vbCrLf & _
                  "You will not be able to undo this process!" & vbCrLf & _
                  "Continue?", vbQuestion + vbYesNo) = vbYes Then

30            sql = "DELETE from BioRepeats WHERE " & _
                    "SampleID = '" & frmEditAll.txtSampleID & "'"

40            Cnxn(0).Execute sql


50            frmEditAll.LoadBiochemistry
60            Unload Me

70        End If

80        Exit Sub

bDELETE_Click_Error:

          Dim strES As String
          Dim intEL As Integer



90        intEL = Erl
100       strES = Err.Description
110       LogError "frmViewBioRepeat", "bDELETE_Click", intEL, strES, sql


End Sub

'Private Sub bTransfer_Click()
'
'          Dim Y As Long
'          Dim sqlFrom As String
'          Dim sqlTo As String
'          Dim fld As Field
'          Dim tbFrom As Recordset
'          Dim tbTo As Recordset
'          Dim Code As String
'          Dim sql As String
'
'10        On Error GoTo bTransfer_Click_Error
'
'20        g.Col = 0
'30        For Y = 1 To g.Rows - 1
'40            g.Row = Y
'50            If g.CellBackColor = vbRed Then
'60                Code = CodeForShortName(g)
'70                sqlFrom = "SELECT * from BioRepeats WHERE " & _
'                            "SampleID = " & frmEditAll.txtSampleID & " " & _
'                            "and Code = '" & Code & "' and result = '" & (g.TextMatrix(g.RowSel, 1)) & "'"
'80                sqlTo = "SELECT * from BioResults WHERE " & _
'                          "SampleID = " & frmEditAll.txtSampleID & " " & _
'                          "and Code = '" & Code & "'"
'
'90                Set tbFrom = New Recordset
'100               RecOpenClient 0, tbFrom, sqlFrom
'
'110               Set tbTo = New Recordset
'120               RecOpenClient 0, tbTo, sqlTo
'
'130               If tbTo.EOF Then
'140                   tbTo.AddNew
'                      '    Else
'                      '      Archive 0, tbFrom, "ArcBioRepeats"
'150               End If
'160               For Each fld In tbTo.Fields
'170                   If UCase(fld.Name) <> UCase("rowguid") Then
'180                       tbTo(fld.Name) = tbFrom(fld.Name)
'190                   End If
'200               Next
'210               tbTo.Update
'220               sql = "DELETE from biorepeats WHERE sampleid = '" & frmEditAll.txtSampleID & "' and code = '" & Code & "' and result = '" & g.TextMatrix(g.RowSel, 1) & "'"
'230               Cnxn(0).Execute sql
'240           End If
'250       Next
'
'260       frmEditAll.LoadBiochemistry
'270       FillG
'
'280       Exit Sub
'
'bTransfer_Click_Error:
'
'          Dim strES As String
'          Dim intEL As Integer
'
'
'
'290       intEL = Erl
'300       strES = Err.Description
'310       LogError "frmViewBioRepeat", "bTransfer_Click", intEL, strES, sql
'
'
'End Sub

Private Function BioCodeFor(TestName As String) As String

    Dim sql As String
    Dim tb As New Recordset
    Set tb = New Recordset
    Dim Code As String
    sql = "SELECT Top 1 Code from Biotestdefinitions WHERE ShortName = '" & TestName & "' AND inuse = 1"
    RecOpenClient 0, tb, sql
    
'    Do While Not tb.EOF
        Code = tb!Code
'    Loop
    BioCodeFor = Code
End Function



Private Sub bTransfer_Click()
          Dim n As Long
          Dim sql As String
          Dim tb As New Recordset
          Dim TestCode As String
          Dim Units As String


10        On Error GoTo bTransfer_Click_Error

20        For n = 1 To g.Rows - 1
30            g.Row = n
40            g.Col = 0
50            If g.CellBackColor = vbRed Then
60                TestCode = BioCodeFor(g.TextMatrix(n, 0))
70                Units = Trim(g.TextMatrix(n, 2))
                  'MsgBox (TestCode & " " & Units & " " & frmEditAll.txtSampleID)
80                sql = "SELECT * from Bioresults WHERE " & _
                        "SampleID = '" & frmEditAll.txtSampleID & "' " & _
                        "and Code = '" & TestCode & "' and units = '" & Units & "'"
90                Set tb = New Recordset
100               RecOpenClient 0, tb, sql
110               With tb
120                   If .EOF Then .AddNew
130                   !Code = TestCode
140                   !Printed = False
150                   !Valid = 0
160                   !Result = g.TextMatrix(g.Row, 1)
170                   !Rundate = Format$(Now, "dd/mmm/yyyy")
180                   !RunTime = Format$(Now, "dd/mmm/yyyy hh:mm")
190                   !SampleID = frmEditAll.txtSampleID
200                   !Units = Units
210                   .Update
220               End With

230               sql = "DELETE from bioRepeats WHERE " & _
                        "SampleID = '" & frmEditAll.txtSampleID & "' " & _
                        "and Code = '" & TestCode & "' and units = '" & Units & "'"
240               Set tb = New Recordset
250               RecOpenClient 0, tb, sql
260           End If
270       Next

280       FillG
290       bTransfer.Enabled = False

300       Exit Sub

bTransfer_Click_Error:

          Dim strES As String
          Dim intEL As Integer

310       intEL = Erl
320       strES = Err.Description
330       LogError "frmCoagRepeats", "bTransfer_Click_Click", intEL, strES
End Sub

Private Sub Form_Activate()

10        On Error GoTo Form_Activate_Error

20        If Activated Then Exit Sub

30        Activated = True

40        FillG

50        Exit Sub

Form_Activate_Error:

          Dim strES As String
          Dim intEL As Integer



60        intEL = Erl
70        strES = Err.Description
80        LogError "frmViewBioRepeat", "Form_Activate", intEL, strES


End Sub

Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        Activated = False

30        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmViewBioRepeat", "Form_Load", intEL, strES


End Sub

Private Sub Form_Unload(Cancel As Integer)

10        On Error GoTo Form_Unload_Error

20        Activated = False

30        Exit Sub

Form_Unload_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmViewBioRepeat", "Form_Unload", intEL, strES


End Sub

Private Sub g_Click()

          Dim Y As Long
          Dim n As Long
          '
10        On Error GoTo g_Click_Error

20        If g.MouseRow = 0 Then Exit Sub

30        bTransfer.Visible = False

40        For n = 1 To frmEditAll.gBio.Rows - 1
50            If frmEditAll.gBio.TextMatrix(n, 0) = g.TextMatrix(g.RowSel, 0) Then
60                If InStr(frmEditAll.gBio.TextMatrix(n, 5), "V") > 0 Then
70                    Exit Sub
80                End If
90            End If
100       Next

110       g.Col = 0
120       If g.CellBackColor = vbRed Then
130           g.CellBackColor = 0
140       Else
150           g.CellBackColor = vbRed
160           bTransfer.Visible = True
170           Exit Sub
180       End If

190       For Y = 1 To g.Rows - 1
200           g.Row = Y
210           If g.CellBackColor = vbRed Then
220               bTransfer.Visible = True
230               Exit For
240           End If
250       Next

260       Exit Sub

g_Click_Error:

          Dim strES As String
          Dim intEL As Integer

270       intEL = Erl
280       strES = Err.Description
290       LogError "frmViewBioRepeat", "g_Click", intEL, strES

End Sub


