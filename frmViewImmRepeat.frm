VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmViewImmRepeat 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Immunology Repeats"
   ClientHeight    =   5685
   ClientLeft      =   4770
   ClientTop       =   2475
   ClientWidth     =   5175
   Icon            =   "frmViewImmRepeat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton bCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   765
      Left            =   3690
      Picture         =   "frmViewImmRepeat.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4620
      Width           =   1215
   End
   Begin VB.CommandButton bTransfer 
      Caption         =   "Copy to Main File"
      Height          =   855
      Left            =   3690
      Picture         =   "frmViewImmRepeat.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   450
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton bDelete 
      Caption         =   "&Delete All Repeats"
      Height          =   945
      Left            =   3690
      Picture         =   "frmViewImmRepeat.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1740
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Highlight Tests to be Transferred"
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   3690
      TabIndex        =   5
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tests in RED will be Transfered"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   210
      Width           =   2445
   End
End
Attribute VB_Name = "frmViewImmRepeat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Activated As Boolean



Private Sub FillG()

          Dim s As String
          Dim sngValue As Single
          Dim strValue As String
          Dim IR As BIEResult
          Dim IRs As New BIEResults
          Dim Cat As String

10        On Error GoTo FillG_Error

20        If frmEditAll.cCat(1) = "" Then Cat = "Default" Else Cat = frmEditAll.cCat(1)

30        Set IRs = IRs.Load("Imm", frmEditAll.txtSampleID, "Repeats", gDONTCARE, gDONTCARE, 0, Cat, frmEditAll.dtRunDate)

40        g.Rows = 2
50        g.AddItem ""
60        g.RemoveItem 1

70        For Each IR In IRs
80            If IsNumeric(IR.Result) Then
90                sngValue = Val(IR.Result)
100               Select Case IR.Printformat
                  Case 0: strValue = Format$(sngValue, "0")
110               Case 1: strValue = Format$(sngValue, "0.0")
120               Case 2: strValue = Format$(sngValue, "0.00")
130               Case 3: strValue = Format$(sngValue, "0.000")
140               Case Else: strValue = Format$(sngValue, "0.000")
150               End Select
160           Else
170               strValue = IR.Result
180           End If
190           s = IR.ShortName & vbTab & _
                  strValue & vbTab & _
                  IR.Units
200           g.AddItem s
210       Next
220       If g.Rows > 2 Then
230           g.RemoveItem 1
240       End If

250       Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer



260       intEL = Erl
270       strES = Err.Description
280       LogError "frmViewImmRepeat", "FillG", intEL, strES


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

30            sql = "DELETE from ImmRepeats WHERE " & _
                    "SampleID = '" & frmEditAll.txtSampleID & "'"

40            Cnxn(0).Execute sql

50            frmEditAll.LoadImmunology
60            Unload Me

70        End If

80        Exit Sub

bDELETE_Click_Error:

          Dim strES As String
          Dim intEL As Integer



90        intEL = Erl
100       strES = Err.Description
110       LogError "frmViewImmRepeat", "bDELETE_Click", intEL, strES, sql


End Sub

Private Sub bTransfer_Click()

          Dim Y As Long
          Dim sqlFrom As String
          Dim sqlTo As String
          Dim fld As Field
          Dim tbFrom As Recordset
          Dim tbTo As Recordset
          Dim Code As String


10        On Error GoTo bTransfer_Click_Error

20        g.Col = 0
30        For Y = 1 To g.Rows - 1
40            g.Row = Y
50            If g.CellBackColor = vbRed Then
60                Code = ICodeForShortName(g)
70                sqlFrom = "SELECT * from immRepeats WHERE " & _
                            "SampleID = '" & frmEditAll.txtSampleID & "' " & _
                            "and Code = '" & Code & "'"
80                sqlTo = "SELECT * from immResults WHERE " & _
                          "SampleID = '" & frmEditAll.txtSampleID & "' " & _
                          "and Code = '" & Code & "'"

90                Set tbFrom = New Recordset
100               RecOpenClient 0, tbFrom, sqlFrom

110               Set tbTo = New Recordset
120               RecOpenClient 0, tbTo, sqlTo

130               If tbTo.EOF Then
140                   tbTo.AddNew
150               End If
160               For Each fld In tbTo.Fields
170                   tbTo(fld.Name) = tbFrom(fld.Name)
180               Next
190               tbTo.Update
200               sqlFrom = "delete from immRepeats WHERE " & _
                            "SampleID = '" & frmEditAll.txtSampleID & "' " & _
                            "and Code = '" & Code & "'"
210               Cnxn(0).Execute sqlFrom
220           End If
230       Next

240       frmEditAll.LoadImmunology
250       Unload Me

260       Exit Sub

bTransfer_Click_Error:

          Dim strES As String
          Dim intEL As Integer



270       intEL = Erl
280       strES = Err.Description
290       LogError "frmViewImmRepeat", "bTransfer_Click", intEL, strES, sqlFrom


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
80        LogError "frmViewImmRepeat", "Form_Activate", intEL, strES


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
60        LogError "frmViewImmRepeat", "Form_Load", intEL, strES


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
60        LogError "frmViewImmRepeat", "Form_Unload", intEL, strES


End Sub

Private Sub g_Click()

          Dim Y As Long

10        On Error GoTo g_Click_Error

20        If g.MouseRow = 0 Then Exit Sub

30        bTransfer.Visible = False

40        g.Col = 0
50        If g.CellBackColor = vbRed Then
60            g.CellBackColor = 0
70        Else
80            g.CellBackColor = vbRed
90            bTransfer.Visible = True
100           Exit Sub
110       End If

120       For Y = 1 To g.Rows - 1
130           g.Row = Y
140           If g.CellBackColor = vbRed Then
150               bTransfer.Visible = True
160               Exit For
170           End If
180       Next

190       Exit Sub

g_Click_Error:

          Dim strES As String
          Dim intEL As Integer



200       intEL = Erl
210       strES = Err.Description
220       LogError "frmViewImmRepeat", "g_Click", intEL, strES


End Sub


