VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmBatchHiv 
   Caption         =   "Netacquire - Batch Hiv"
   ClientHeight    =   6330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5340
   Icon            =   "frmBatchHiv.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   5340
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSCommand cmdSave 
      Height          =   870
      Left            =   675
      TabIndex        =   6
      Top             =   5220
      Width           =   1365
      _Version        =   65536
      _ExtentX        =   2408
      _ExtentY        =   1535
      _StockProps     =   78
      Caption         =   "Save"
      Enabled         =   0   'False
      Picture         =   "frmBatchHiv.frx":030A
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "Set"
      Height          =   285
      Left            =   4230
      TabIndex        =   5
      Top             =   180
      Width           =   735
   End
   Begin VB.TextBox txtTo 
      Height          =   285
      Left            =   2970
      TabIndex        =   3
      Top             =   180
      Width           =   1095
   End
   Begin VB.TextBox txtFrom 
      Height          =   285
      Left            =   1530
      TabIndex        =   1
      Top             =   180
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid grdHiv 
      Height          =   4515
      Left            =   495
      TabIndex        =   4
      Top             =   675
      Width           =   4170
      _ExtentX        =   7355
      _ExtentY        =   7964
      _Version        =   393216
      Cols            =   4
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
      FormatString    =   "<Specimen  |<Chart         |<Name                      |<Hiv     "
   End
   Begin Threed.SSCommand cmdCancel 
      Height          =   870
      Left            =   3150
      TabIndex        =   7
      Top             =   5220
      Width           =   1365
      _Version        =   65536
      _ExtentX        =   2408
      _ExtentY        =   1535
      _StockProps     =   78
      Caption         =   "Exit"
      Picture         =   "frmBatchHiv.frx":0624
   End
   Begin VB.Label Label2 
      Caption         =   "To"
      Height          =   285
      Left            =   2700
      TabIndex        =   2
      Top             =   180
      Width           =   555
   End
   Begin VB.Label Label1 
      Caption         =   "Number From"
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Top             =   180
      Width           =   1095
   End
End
Attribute VB_Name = "frmBatchHiv"
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
          Dim n As Long
          Dim s As String
          Dim PosFound As Boolean


10        On Error GoTo cmdSave_Click_Error

20        PosFound = False

30        s = "The following numbers were positive : " & vbCrLf

40        For n = 1 To grdHiv.Rows - 1
50            If grdHiv.TextMatrix(n, 3) = "Pos" Then
60                PosFound = True
70                s = s & grdHiv.TextMatrix(n, 0) & vbCrLf
80            End If
90        Next

100       If PosFound = True Then
110           If iMsg(s, vbYesNo) = vbNo Then
120               Exit Sub
130           End If
140       End If

150       For n = 1 To grdHiv.Rows - 1
160           sql = "SELECT * from immresults WHERE sampleid = '" & grdHiv.TextMatrix(n, 0) & "' and code = '" & SysOptHivCode(0) & "'"
170           Set tb = New Recordset
180           RecOpenServer 0, tb, sql
190           If tb.EOF Then tb.AddNew
200           tb!SampleID = grdHiv.TextMatrix(n, 0)
210           tb!Rundate = Format(Now, "dd/MMM/yyyy")
220           tb!Code = SysOptHivCode(0)
230           tb!Result = grdHiv.TextMatrix(n, 3)
240           tb!Valid = 1
250           tb!Printed = 0
260           tb!SampleType = "S"
270           tb.Update
280       Next

290       ClearFGrid grdHiv

300       FixG grdHiv

310       txtFrom = ""
320       txtTo = ""
330       cmdSave.Enabled = False

340       Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

350       intEL = Erl
360       strES = Err.Description
370       LogError "frmBatchHiv", "cmdsave_Click", intEL, strES, sql


End Sub

Private Sub cmdSet_Click()
          Dim n As Long
          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo cmdSet_Click_Error

20        If Not IsNumeric(txtFrom) Or Not IsNumeric(txtTo) Then
30            iMsg "Number must be Numeric"
40            Exit Sub
50        End If

60        If Val(txtFrom) >= Val(txtTo) Then
70            iMsg "From must be less than to !"
80            Exit Sub
90        End If

100       ClearFGrid grdHiv

110       For n = Val(txtFrom) To Val(txtTo)
120           sql = "SELECT * FROM Demographics WHERE " & _
                    "SampleID = '" & n & "'"
130           Set tb = New Recordset
140           RecOpenServer 0, tb, sql
150           If Not tb.EOF Then
160               grdHiv.AddItem n & vbTab & Trim(tb!Chart & "") & vbTab & Trim(tb!PatName & "") & vbTab & "Neg"
170           Else
180               grdHiv.AddItem n & vbTab & vbTab & vbTab & "Neg"
190           End If
200       Next

210       FixG grdHiv

220       cmdSave.Enabled = True

230       Exit Sub

cmdSet_Click_Error:

          Dim strES As String
          Dim intEL As Integer

240       intEL = Erl
250       strES = Err.Description
260       LogError "frmBatchHiv", "cmdSet_Click", intEL, strES, sql


End Sub

Private Sub grdHiv_Click()

10        On Error GoTo grdHiv_Click_Error

20        If grdHiv.MouseRow = 0 Then Exit Sub


30        If grdHiv.Rows = 2 And grdHiv.TextMatrix(1, 0) = "" Then Exit Sub

40        If grdHiv.ColSel = 3 Then
50            If grdHiv.TextMatrix(grdHiv.RowSel, 3) = "Neg" Then
60                grdHiv.TextMatrix(grdHiv.RowSel, 3) = "Pos"
70            Else
80                grdHiv.TextMatrix(grdHiv.RowSel, 3) = "Neg"
90            End If
100       End If

110       Exit Sub

grdHiv_Click_Error:

          Dim strES As String
          Dim intEL As Integer

120       intEL = Erl
130       strES = Err.Description
140       LogError "frmBatchHiv", "grdHiv_Click", intEL, strES


End Sub
