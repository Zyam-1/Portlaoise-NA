VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmHaemImm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Add Haematology Result"
   ClientHeight    =   6795
   ClientLeft      =   465
   ClientTop       =   495
   ClientWidth     =   11820
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "frmHaemImm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "frmHaemImm.frx":030A
   ScaleHeight     =   6795
   ScaleWidth      =   11820
   Begin VB.CommandButton cmdSetPrinter 
      Height          =   840
      Left            =   10845
      Picture         =   "frmHaemImm.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Back Colour Normal:Automatic Printer Selection.Back Colour Red:-Forced"
      Top             =   2610
      Width           =   915
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   10845
      Picture         =   "frmHaemImm.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5580
      Width           =   885
   End
   Begin MSComCtl2.DTPicker dtDate 
      Height          =   315
      Left            =   990
      TabIndex        =   3
      Top             =   90
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      _Version        =   393216
      Format          =   59113473
      CurrentDate     =   37082
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Save  && &Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   10845
      Picture         =   "frmHaemImm.frx":0C28
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3570
      Width           =   885
   End
   Begin MSFlexGridLib.MSFlexGrid grdHaem 
      Height          =   6135
      Left            =   90
      TabIndex        =   1
      Top             =   450
      Width           =   10650
      _ExtentX        =   18785
      _ExtentY        =   10821
      _Version        =   393216
      Cols            =   15
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      GridLines       =   3
      GridLinesFixed  =   3
      FormatString    =   "<Specimen  |<Chart         |<Name                      |<Ward               |<Gp                 |<Clinician            "
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   10845
      Picture         =   "frmHaemImm.frx":0F32
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4590
      Width           =   885
   End
   Begin VB.TextBox txtInput 
      Height          =   330
      Left            =   5805
      TabIndex        =   9
      Top             =   2565
      Width           =   1860
   End
   Begin VB.Frame Frame1 
      Caption         =   "Show"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1395
      TabIndex        =   4
      Top             =   2340
      Width           =   3255
      Begin VB.OptionButton optView 
         Caption         =   "&All"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   135
         TabIndex        =   7
         Top             =   210
         Width           =   675
      End
      Begin VB.OptionButton optView 
         Caption         =   "&Incomplete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   930
         TabIndex        =   6
         Top             =   210
         Value           =   -1  'True
         Width           =   1245
      End
      Begin VB.OptionButton optView 
         Caption         =   "&Ordered"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   2190
         TabIndex        =   5
         Top             =   210
         Width           =   1035
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Rundate"
      Height          =   285
      Left            =   180
      TabIndex        =   11
      Top             =   135
      Width           =   1005
   End
   Begin VB.Label Label1 
      Height          =   555
      Left            =   -180
      TabIndex        =   10
      Top             =   -45
      Width           =   825
   End
End
Attribute VB_Name = "frmHaemImm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pPrintToPrinter As String
Private Activated As Boolean

Private Sub cmdCancel_Click()

10        Unload Me

End Sub

Private Sub cmdPrint_Click()

          Dim tb As New Recordset
          Dim sql As String
          Dim Num As Long
          Dim cNum As Long

10        On Error GoTo cmdPrint_Click_Error

20        If grdHaem.Rows = 2 And grdHaem.TextMatrix(1, 0) = "" Then Exit Sub

30        For Num = 1 To grdHaem.Rows - 1
40            For cNum = 6 To grdHaem.Cols - 1
50                If grdHaem.TextMatrix(Num, cNum) <> "" Then
60                    sql = "delete from immrequests where sampleid = " & grdHaem.TextMatrix(Num, 0) & " and code = '" & ICodeForShortName(grdHaem.TextMatrix(0, cNum)) & "'"
70                    Cnxn(0).Execute sql
80                    sql = "SELECT * from immresults WHERE " & _
                            "SampleID = '" & grdHaem.TextMatrix(Num, 0) & "' " & _
                            "and code = '" & ICodeForShortName(grdHaem.TextMatrix(0, cNum)) & "'"
90                    Set tb = New Recordset
100                   RecOpenServer 0, tb, sql
110                   If tb.EOF Then tb.AddNew
120                   tb!SampleID = grdHaem.TextMatrix(Num, 0)
130                   tb!Code = ICodeForShortName(grdHaem.TextMatrix(0, cNum))
140                   tb!Result = grdHaem.TextMatrix(Num, cNum)
150                   tb!SampleType = "S"
160                   tb!Valid = 1
170                   tb!Printed = 0
180                   tb!RunTime = Format(Now, "dd/MMM/yyyy hh:mm:ss")
190                   tb!Rundate = Format(Now, "dd/MMM/yyyy")
200                   tb!Units = GetImmUnit(ICodeForShortName(grdHaem.TextMatrix(0, cNum))) & ""
210                   tb!Operator = UserCode
220                   tb.Update
230                   sql = "SELECT * FROM PrintPending WHERE " & _
                            "Department = 'J' " & _
                            "AND SampleID = '" & grdHaem.TextMatrix(Num, 0) & "'"
240                   Set tb = New Recordset
250                   RecOpenServer 0, tb, sql
260                   If tb.EOF Then tb.AddNew
270                   tb!SampleID = grdHaem.TextMatrix(Num, 0)
280                   tb!Department = "J"
290                   tb!Initiator = Username
300                   tb!Ward = grdHaem.TextMatrix(Num, 3) & ""
310                   tb!Clinician = grdHaem.TextMatrix(Num, 5) & ""
320                   tb!GP = grdHaem.TextMatrix(Num, 4) & ""
330                   tb!UsePrinter = pPrintToPrinter
340                   tb!pTime = Now
350                   tb.Update
360               End If
370           Next
380       Next

390       FillGrid



400       Exit Sub

cmdPrint_Click_Error:

          Dim strES As String
          Dim intEL As Integer



410       intEL = Erl
420       strES = Err.Description
430       LogError "frmHaemImm", "cmdPrint_Click", intEL, strES, sql

End Sub

Private Sub cmdSave_Click()

          Dim tb As New Recordset
          Dim sql As String
          Dim Num As Long
          Dim cNum As Long

10        On Error GoTo cmdSave_Click_Error

20        If grdHaem.Rows = 2 And grdHaem.TextMatrix(1, 0) = "" Then Exit Sub

30        For Num = 1 To grdHaem.Rows - 1
40            For cNum = 6 To grdHaem.Cols - 1
50                If grdHaem.TextMatrix(Num, cNum) <> "" Then
60                    sql = "DELETE FROM ImmRequests WHERE " & _
                            "SampleID = '" & grdHaem.TextMatrix(Num, 0) & "' " & _
                            "AND Code = '" & ICodeForShortName(grdHaem.TextMatrix(0, cNum)) & "'"
70                    Cnxn(0).Execute sql
80                    sql = "SELECT * from immresults WHERE " & _
                            "SampleID = '" & grdHaem.TextMatrix(Num, 0) & "' " & _
                            "and code = '" & ICodeForShortName(grdHaem.TextMatrix(0, cNum)) & "'"
90                    Set tb = New Recordset
100                   RecOpenServer 0, tb, sql
110                   If tb.EOF Then
120                       tb.AddNew
130                   End If
140                   tb!SampleID = grdHaem.TextMatrix(Num, 0)
150                   tb!Code = ICodeForShortName(grdHaem.TextMatrix(0, cNum))
160                   tb!Result = grdHaem.TextMatrix(Num, cNum)
170                   tb!SampleType = "S"
180                   tb!Valid = 1
190                   tb!Printed = 0
200                   tb!RunTime = Format(Now, "dd/MMM/yyyy hh:mm:ss")
210                   tb!Rundate = Format(Now, "dd/MMM/yyyy")
220                   tb!Units = GetImmUnit(ICodeForShortName(grdHaem.TextMatrix(0, cNum)) & "")
230                   tb!Operator = UserCode
240                   tb.Update
250               End If
260           Next
270       Next

280       FillGrid



290       cmdSave.Enabled = False    '




300       Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer



310       intEL = Erl
320       strES = Err.Description
330       LogError "frmHaemImm", "cmdsave_Click", intEL, strES, sql


End Sub

Private Sub dtDate_CloseUp()

10        On Error GoTo dtDate_CloseUp_Error

20        FillGrid

30        Exit Sub

dtDate_CloseUp_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmHaemImm", "dtDate_CloseUp", intEL, strES


End Sub

Private Sub FillGrid()

          Dim sn As New Recordset
          Dim sql As String
          Dim Str As String
          Dim intN As Integer

10        On Error GoTo FillGrid_Error

20        ClearFGrid grdHaem

30        grdHaem.Visible = False

40        sql = "SELECT D.*, I.* FROM " & _
                "Demographics AS D, ImmRequests AS I WHERE " & _
                "D.RunDate = '" & Format$(dtDate, "dd/mmm/yyyy") & "' " & _
                "AND D.SampleID = I.SampleID " & _
                "AND ( "

50        For intN = 6 To 14
60            If grdHaem.ColWidth(intN) > 0 Then
70                sql = sql & " I.code = '" & ICodeForShortName(grdHaem.TextMatrix(0, intN)) & "' OR "
80            End If
90        Next
100       sql = Left(sql, Len(sql) - 4)

110       sql = sql & " ) ORDER BY D.SampleID"

120       Set sn = New Recordset
130       RecOpenServer 0, sn, sql
140       Do While Not sn.EOF
150           If sn!SampleID <> grdHaem.TextMatrix(grdHaem.Rows - 1, 0) Then
160               Str = Trim$(sn!SampleID) & vbTab & _
                        Trim$(sn!Chart & "") & vbTab & _
                        sn!PatName & vbTab & _
                        Trim(sn!Ward) & vbTab & _
                        Trim(sn!GP & "") & vbTab & _
                        Trim(sn!Clinician & "") & vbTab
170               grdHaem.AddItem Str
180           End If

190           For intN = 6 To 14
200               If sn!Code = ICodeForShortName(grdHaem.TextMatrix(0, intN)) Then
210                   grdHaem.Row = grdHaem.Rows - 1
220                   grdHaem.Col = intN
230                   grdHaem.CellBackColor = vbYellow
240               End If
250           Next

260           sn.MoveNext
270       Loop

280       FixG grdHaem



290       Exit Sub

FillGrid_Error:

          Dim strES As String
          Dim intEL As Integer



300       intEL = Erl
310       strES = Err.Description
320       LogError "frmHaemImm", "FillGrid", intEL, strES

End Sub

Private Sub Form_Activate()

10        On Error GoTo Form_Activate_Error

20        If Not Activated Then
30            Activated = True

40            FillG
50            FillGrid
60        End If

70        Set_Font Me

80        Exit Sub

Form_Activate_Error:

          Dim strES As String
          Dim intEL As Integer



90        intEL = Erl
100       strES = Err.Description
110       LogError "frmHaemImm", "Form_Activate", intEL, strES


End Sub

Private Sub cmdSetPrinter_Click()

10        On Error GoTo cmdSetPrinter_Click_Error

20        Set frmForcePrinter.f = frmHaemImm
30        frmForcePrinter.Show 1

40        If pPrintToPrinter = "Automatic SELECTion" Then
50            pPrintToPrinter = ""
60        End If

70        If pPrintToPrinter <> "" Then
80            cmdSetPrinter.BackColor = vbRed
90            cmdSetPrinter.ToolTipText = "Print Forced to " & pPrintToPrinter
100       Else
110           cmdSetPrinter.BackColor = vbButtonFace
120           pPrintToPrinter = ""
130           cmdSetPrinter.ToolTipText = "Printer SELECTed Automatically"
140       End If

150       Exit Sub

cmdSetPrinter_Click_Error:

          Dim strES As String
          Dim intEL As Integer



160       intEL = Erl
170       strES = Err.Description
180       LogError "frmHaemImm", "cmdSetPrinter_Click", intEL, strES


End Sub

Public Property Let PrintToPrinter(ByVal strNewValue As String)

10        On Error GoTo PrintToPrinter_Error

20        pPrintToPrinter = strNewValue

30        Exit Property

PrintToPrinter_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmHaemImm", "PrintToPrinter", intEL, strES


End Property

Public Property Get PrintToPrinter() As String

10        On Error GoTo PrintToPrinter_Error

20        PrintToPrinter = pPrintToPrinter

30        Exit Property

PrintToPrinter_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmHaemImm", "PrintToPrinter", intEL, strES


End Property

Private Sub FillG()

          Dim sql As String
          Dim tb As Recordset
          Dim intN As Integer

10        On Error GoTo FillG_Error

20        For intN = 6 To 14
30            grdHaem.ColWidth(intN) = 0
40        Next

50        intN = 6
60        sql = "select distinct(shortname) from immtestdefinitions where haem = 1 and inuse = 1"
70        Set tb = New Recordset
80        RecOpenServer 0, tb, sql
90        Do While Not tb.EOF
100           grdHaem.TextMatrix(0, intN) = tb!ShortName
110           grdHaem.ColWidth(intN) = 1000
120           intN = intN + 1
130           tb.MoveNext
140       Loop


150       Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer



160       intEL = Erl
170       strES = Err.Description
180       LogError "frmHaemImm", "FillG", intEL, strES, sql


End Sub
Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        Activated = False

30        dtDate = Format$(Now, "dd/mm/yyyy")

40        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer



50        intEL = Erl
60        strES = Err.Description
70        LogError "frmHaemImm", "Form_Load", intEL, strES


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
60        LogError "frmHaemImm", "Form_Unload", intEL, strES


End Sub

Private Sub grdHaem_Click()

10        On Error GoTo grdHaem_Click_Error

20        If grdHaem.MouseRow = 0 Or grdHaem.Col < 6 Then Exit Sub

30        If grdHaem.TextMatrix(grdHaem.Row, 0) = "" Then Exit Sub

40        If grdHaem.CellBackColor <> vbYellow Then Exit Sub

50        If grdHaem.Col > 5 Then
60            txtInput.Text = grdHaem.TextMatrix(grdHaem.Row, grdHaem.Col)
70            txtInput.SetFocus
80            Exit Sub
90        End If

100       Exit Sub

grdHaem_Click_Error:

          Dim strES As String
          Dim intEL As Integer



110       intEL = Erl
120       strES = Err.Description
130       LogError "frmHaemImm", "grdHaem_Click", intEL, strES

End Sub

Private Sub optView_Click(Index As Integer)

10        On Error GoTo optView_Click_Error

20        FillGrid

30        Exit Sub

optView_Click_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmHaemImm", "optView_Click", intEL, strES


End Sub

Private Sub txtInput_Change()


10        On Error GoTo txtInput_Change_Error

20        txtInput.SelStart = Len(txtInput)

30        grdHaem.TextMatrix(grdHaem.RowSel, grdHaem.ColSel) = Trim(txtInput)

40        If grdHaem.TextMatrix(grdHaem.RowSel, grdHaem.ColSel) = "" Then
50            grdHaem.TextMatrix(grdHaem.RowSel, grdHaem.ColSel) = ""
60        End If

70        cmdSave.Enabled = True

80        Exit Sub

txtInput_Change_Error:

          Dim strES As String
          Dim intEL As Integer



90        intEL = Erl
100       strES = Err.Description
110       LogError "frmHaemImm", "txtInput_Change", intEL, strES


End Sub



Private Sub txtInput_KeyPress(KeyAscii As Integer)
          Dim intN As Integer

10        On Error GoTo txtInput_KeyPress_Error

20        intN = KeyAscii

30        If KeyAscii <> 8 Then
40            KeyAscii = 0
50            If intN = 80 Or intN = 112 Then
60                txtInput = "Positive"
70            ElseIf intN = 78 Or intN = 110 Then
80                txtInput = "Negative"
90            Else
100               txtInput = txtInput & Chr(intN)
110           End If
120       End If

130       Exit Sub

txtInput_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer



140       intEL = Erl
150       strES = Err.Description
160       LogError "frmHaemImm", "txtInput_KeyPress", intEL, strES


End Sub

