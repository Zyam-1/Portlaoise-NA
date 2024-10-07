VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCoagRepeats 
   Caption         =   "NetAcquire - Coagulation Repeats"
   ClientHeight    =   2940
   ClientLeft      =   5595
   ClientTop       =   3540
   ClientWidth     =   4545
   Icon            =   "frmCoagRepeats.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   4545
   Begin VB.CommandButton bDelete 
      Caption         =   "&Delete All Repeats"
      Height          =   930
      Left            =   3150
      Picture         =   "frmCoagRepeats.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   990
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Exit"
      Height          =   750
      Left            =   3150
      Picture         =   "frmCoagRepeats.frx":040C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2115
      Width           =   1245
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy to &Result"
      Enabled         =   0   'False
      Height          =   840
      Left            =   3150
      Picture         =   "frmCoagRepeats.frx":0716
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   60
      Width           =   1245
   End
   Begin MSFlexGridLib.MSFlexGrid grdRep 
      Height          =   2805
      Left            =   60
      TabIndex        =   3
      Top             =   60
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   4948
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      FormatString    =   "<Parameter            |<Result    |<Units     "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmCoagRepeats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mSampleID As String
Private mEditForm As Form



Private Sub cmdCancel_Click()

10        On Error GoTo cmdCancel_Click_Error

20        mEditForm.LoadCoagulation

30        Unload Me

40        Exit Sub

cmdCancel_Click_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmCoagRepeats", "cmdCancel_Click", intEL, strES

End Sub

Private Sub cmdCopy_Click()

          Dim n As Long
          Dim sql As String
          Dim tb As New Recordset
          Dim TestCode As String
          Dim Units As String


10        On Error GoTo cmdCopy_Click_Error

20        For n = 1 To grdRep.Rows - 1
30            grdRep.Row = n
40            grdRep.Col = 0
50            If grdRep.CellBackColor = vbYellow Then
60                TestCode = CoagCodeFor(grdRep.TextMatrix(n, 0))
70                Units = Trim(grdRep.TextMatrix(n, 2))
80                sql = "SELECT * from CoagResults WHERE " & _
                        "SampleID = '" & mSampleID & "' " & _
                        "and Code = '" & TestCode & "' and units = '" & Units & "'"
90                Set tb = New Recordset
100               RecOpenClient 0, tb, sql
110               With tb
120                   If .EOF Then .AddNew
130                   !Code = TestCode
140                   !Printed = False
150                   !Valid = 0
160                   !Result = grdRep.TextMatrix(grdRep.Row, 1)
170                   !Rundate = Format$(Now, "dd/mmm/yyyy")
180                   !RunTime = Format$(Now, "dd/mmm/yyyy hh:mm")
190                   !SampleID = mSampleID
200                   !Units = Units
210                   .Update
220               End With

230               sql = "DELETE from CoagRepeats WHERE " & _
                        "SampleID = '" & mSampleID & "' " & _
                        "and Code = '" & TestCode & "' and units = '" & Units & "'"
240               Set tb = New Recordset
250               RecOpenClient 0, tb, sql
260           End If
270       Next

280       FillG

290       cmdCopy.Enabled = False





300       Exit Sub

cmdCopy_Click_Error:

          Dim strES As String
          Dim intEL As Integer

310       intEL = Erl
320       strES = Err.Description
330       LogError "frmCoagRepeats", "cmdCopy_Click", intEL, strES


End Sub



Private Sub bDELETE_Click()

          Dim sql As String

10        On Error GoTo bDELETE_Click_Error

20        sql = "DELETE from CoagRepeats WHERE " & _
                "SampleID = '" & mSampleID & "'"
30        Cnxn(0).Execute sql

40        mEditForm.LoadCoagulation

50        Unload Me

60        Exit Sub

bDELETE_Click_Error:

          Dim strES As String
          Dim intEL As Integer

70        intEL = Erl
80        strES = Err.Description
90        LogError "frmCoagRepeats", "bDELETE_Click", intEL, strES

End Sub

Public Property Let EditForm(ByVal f As Form)

10        On Error GoTo EditForm_Error

20        Set mEditForm = f

30        Exit Property

EditForm_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmCoagRepeats", "EditForm", intEL, strES


End Property

Private Sub FillG()

          Dim cRR As New CoagResults
          Dim CR As CoagResult
          Dim s As String
          Dim sql As String
          Dim tb As New Recordset



10        On Error GoTo FillG_Error

20        Set cRR = cRR.LoadRepeats(mSampleID, gDONTCARE, gDONTCARE, SysOptExp(0))

30        ClearFGrid grdRep

40        For Each CR In cRR
50            sql = "SELECT * from coagtestdefinitions WHERE code = '" & CR.Code & "'"
60            Set tb = New Recordset
70            RecOpenServer 0, tb, sql
80            If Not tb.EOF Then
90                If tb!InUse Then
100                   s = CoagNameFor(CR.Code) & vbTab & CR.Result & vbTab & CR.Units
110                   grdRep.AddItem s
120               End If
130           End If
140       Next

150       FixG grdRep




160       Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

170       intEL = Erl
180       strES = Err.Description
190       LogError "frmCoagRepeats", "FillG", intEL, strES


End Sub

Private Sub Form_Activate()

10        On Error GoTo Form_Activate_Error

20        FillG

30        Exit Sub

Form_Activate_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmCoagRepeats", "Form_Activate", intEL, strES


End Sub

Private Sub grdRep_Click()

10        On Error GoTo grdRep_Click_Error

20        If grdRep.MouseRow = 0 Then Exit Sub

30        grdRep.Col = 0
40        If grdRep.CellBackColor = vbYellow Then
50            grdRep.CellBackColor = 0
60            grdRep.Col = 1
70            grdRep.CellBackColor = 0
80            grdRep.Col = 2
90            grdRep.CellBackColor = 0
100       Else
110           grdRep.CellBackColor = vbYellow
120           grdRep.Col = 1
130           grdRep.CellBackColor = vbYellow
140           grdRep.Col = 2
150           grdRep.CellBackColor = vbYellow
160       End If

170       cmdCopy.Enabled = True

180       Exit Sub

grdRep_Click_Error:

          Dim strES As String
          Dim intEL As Integer

190       intEL = Erl
200       strES = Err.Description
210       LogError "frmCoagRepeats", "grdRep_Click", intEL, strES


End Sub

Public Property Let SampleID(ByVal sNewValue As String)

10        On Error GoTo SampleID_Error

20        mSampleID = sNewValue

30        Exit Property

SampleID_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmCoagRepeats", "SampleID", intEL, strES


End Property
