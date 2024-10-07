VERSION 5.00
Begin VB.Form frmBarCodes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - BarCodes"
   ClientHeight    =   5625
   ClientLeft      =   1005
   ClientTop       =   900
   ClientWidth     =   4695
   ForeColor       =   &H8000000F&
   Icon            =   "frmBarCodes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtMonoSpot 
      Height          =   315
      Left            =   1230
      TabIndex        =   10
      Top             =   4995
      Width           =   1575
   End
   Begin VB.TextBox txtFBC 
      Height          =   315
      Left            =   1230
      TabIndex        =   7
      Top             =   4020
      Width           =   1575
   End
   Begin VB.TextBox txtRetics 
      Height          =   315
      Left            =   1230
      TabIndex        =   9
      Top             =   4680
      Width           =   1575
   End
   Begin VB.TextBox txtESR 
      Height          =   315
      Left            =   1230
      TabIndex        =   8
      Top             =   4350
      Width           =   1575
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   735
      Left            =   3150
      Picture         =   "frmBarCodes.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1125
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.TextBox txtB 
      Height          =   315
      Left            =   1230
      TabIndex        =   6
      Top             =   3540
      Width           =   1575
   End
   Begin VB.TextBox txtA 
      Height          =   315
      Left            =   1230
      TabIndex        =   5
      Top             =   3210
      Width           =   1575
   End
   Begin VB.TextBox txtFasting 
      Height          =   315
      Left            =   1230
      TabIndex        =   4
      Top             =   2790
      Width           =   1575
   End
   Begin VB.TextBox txtRandom 
      Height          =   315
      Left            =   1230
      TabIndex        =   3
      Top             =   2460
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   645
      Left            =   3150
      Picture         =   "frmBarCodes.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2970
      Width           =   1245
   End
   Begin VB.TextBox txtClear 
      Height          =   315
      Left            =   1230
      TabIndex        =   2
      Top             =   2010
      Width           =   1575
   End
   Begin VB.TextBox txtSave 
      Height          =   315
      Left            =   1230
      TabIndex        =   1
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox txtCancel 
      Height          =   315
      Left            =   1230
      TabIndex        =   0
      Top             =   900
      Width           =   1575
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Retics"
      Height          =   195
      Left            =   720
      TabIndex        =   24
      Top             =   4740
      Width           =   450
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "MonoSpot"
      Height          =   195
      Left            =   420
      TabIndex        =   23
      Top             =   5070
      Width           =   735
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "ESR"
      Height          =   195
      Left            =   825
      TabIndex        =   22
      Top             =   4410
      Width           =   330
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "FBC"
      Height          =   195
      Left            =   855
      TabIndex        =   21
      Top             =   4080
      Width           =   300
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Random"
      Height          =   195
      Left            =   555
      TabIndex        =   20
      Top             =   2520
      Width           =   600
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Fasting"
      Height          =   195
      Left            =   645
      TabIndex        =   19
      Top             =   2850
      Width           =   510
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Set Analyser 'B'"
      Height          =   195
      Left            =   60
      TabIndex        =   18
      Top             =   3570
      Width           =   1095
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Set Analyser 'A'"
      Height          =   195
      Left            =   60
      TabIndex        =   17
      Top             =   3270
      Width           =   1095
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Cancel"
      Height          =   195
      Left            =   660
      TabIndex        =   16
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Save"
      Height          =   195
      Left            =   780
      TabIndex        =   15
      Top             =   1500
      Width           =   375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Clear"
      Height          =   195
      Left            =   795
      TabIndex        =   14
      Top             =   2070
      Width           =   360
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Scan Entries using BarCode Reader"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   645
      Left            =   1260
      TabIndex        =   13
      Top             =   90
      Width           =   1560
   End
End
Attribute VB_Name = "frmBarCodes"
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



10        On Error GoTo cmdSave_Click_Error

20        sql = "UPDATE BarCodeControl " & _
                "Set Code = '" & txtCancel & "' " & _
                "WHERE Text = 'ctlCancel'"
30        Cnxn(0).Execute sql

40        sql = "UPDATE BarCodeControl " & _
                "Set Code = '" & txtSave & "' " & _
                "WHERE Text = 'ctlSave'"
50        Cnxn(0).Execute sql

60        sql = "UPDATE BarCodeControl " & _
                "Set Code = '" & txtClear & "' " & _
                "WHERE Text = 'ctlClear'"
70        Cnxn(0).Execute sql

80        sql = "UPDATE BarCodeControl " & _
                "Set Code = '" & txtRandom & "' " & _
                "WHERE Text = 'ctlRandom'"
90        Cnxn(0).Execute sql

100       sql = "UPDATE BarCodeControl " & _
                "Set Code = '" & txtFasting & "' " & _
                "WHERE Text = 'ctlFasting'"
110       Cnxn(0).Execute sql

120       sql = "UPDATE BarCodeControl " & _
                "Set Code = '" & txtA & "' " & _
                "WHERE Text = 'ctlA'"
130       Cnxn(0).Execute sql

140       sql = "UPDATE BarCodeControl " & _
                "Set Code = '" & txtB & "' " & _
                "WHERE Text = 'ctlB'"
150       Cnxn(0).Execute sql

160       sql = "UPDATE BarCodeControl " & _
                "Set Code = '" & txtFBC & "' " & _
                "WHERE Text = 'ctlFBC'"
170       Cnxn(0).Execute sql

180       sql = "UPDATE BarCodeControl " & _
                "Set Code = '" & txtESR & "' " & _
                "WHERE Text = 'ctlESR'"
190       Cnxn(0).Execute sql

200       sql = "UPDATE BarCodeControl " & _
                "Set Code = '" & txtRetics & "' " & _
                "WHERE Text = 'ctlRetics'"
210       Cnxn(0).Execute sql

220       sql = "UPDATE BarCodeControl " & _
                "Set Code = '" & txtMonoSpot & "' " & _
                "WHERE Text = 'ctlMonoSpot'"
230       Cnxn(0).Execute sql

240       cmdSave.Visible = False



250       Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

260       intEL = Erl
270       strES = Err.Description
280       LogError "frmBarCodes", "cmdsave_Click", intEL, strES, sql


End Sub

Private Sub Form_Load()

          Dim sql As String
          Dim tb As New Recordset



10        On Error GoTo Form_Load_Error

20        sql = "SELECT * from BarCodeControl"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        With tb

60            Do While Not .EOF
70                Select Case UCase$(Trim$(!Text))
                  Case "CTLCANCEL": txtCancel = !Code & ""
80                Case "CTLSAVE": txtSave = !Code & ""
90                Case "CTLCLEAR": txtClear = !Code & ""
100               Case "CTLRANDOM": txtRandom = !Code & ""
110               Case "CTLFASTING": txtFasting = !Code & ""
120               Case "CTLA": txtA = !Code & ""
130               Case "CTLB": txtB = !Code & ""
140               Case "CTLFBC": txtFBC = !Code & ""
150               Case "CTLESR": txtESR = !Code & ""
160               Case "CTLRETICS": txtRetics = !Code & ""
170               Case "CTLMONOSPOT": txtMonoSpot = !Code & ""
180               End Select
190               .MoveNext
200           Loop
210       End With

220       Set_Font Me


230       Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

240       intEL = Erl
250       strES = Err.Description
260       LogError "frmBarCodes", "Form_Load", intEL, strES, sql


End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

10        On Error GoTo Form_QueryUnload_Error

20        If cmdSave.Visible Then
30            If iMsg("Cancel without Saving?", vbQuestion + vbYesNo) = vbNo Then
40                Cancel = True
50            End If
60        End If

70        Exit Sub

Form_QueryUnload_Error:

          Dim strES As String
          Dim intEL As Integer

80        intEL = Erl
90        strES = Err.Description
100       LogError "frmBarCodes", "Form_QueryUnload", intEL, strES


End Sub

Private Sub txtA_KeyPress(KeyAscii As Integer)

10        On Error GoTo txtA_KeyPress_Error

20        cmdSave.Visible = True

30        Exit Sub

txtA_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmBarCodes", "txtA_KeyPress", intEL, strES


End Sub

Private Sub txtB_KeyPress(KeyAscii As Integer)

10        On Error GoTo txtB_KeyPress_Error

20        cmdSave.Visible = True

30        Exit Sub

txtB_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmBarCodes", "txtB_KeyPress", intEL, strES


End Sub

Private Sub txtCancel_KeyPress(KeyAscii As Integer)

10        On Error GoTo txtCancel_KeyPress_Error

20        cmdSave.Visible = True

30        Exit Sub

txtCancel_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmBarCodes", "txtCancel_KeyPress", intEL, strES


End Sub

Private Sub txtClear_KeyPress(KeyAscii As Integer)

10        On Error GoTo txtClear_KeyPress_Error

20        cmdSave.Visible = True

30        Exit Sub

txtClear_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmBarCodes", "txtClear_KeyPress", intEL, strES


End Sub

Private Sub txtESR_KeyPress(KeyAscii As Integer)

10        On Error GoTo txtESR_KeyPress_Error

20        cmdSave.Visible = True

30        Exit Sub

txtESR_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmBarCodes", "txtESR_KeyPress", intEL, strES


End Sub

Private Sub txtFasting_KeyPress(KeyAscii As Integer)

10        On Error GoTo txtFasting_KeyPress_Error

20        cmdSave.Visible = True

30        Exit Sub

txtFasting_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmBarCodes", "txtFasting_KeyPress", intEL, strES


End Sub

Private Sub txtFBC_KeyPress(KeyAscii As Integer)

10        On Error GoTo txtFBC_KeyPress_Error

20        cmdSave.Visible = True

30        Exit Sub

txtFBC_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmBarCodes", "txtFBC_KeyPress", intEL, strES


End Sub

Private Sub txtMonoSpot_KeyPress(KeyAscii As Integer)

10        On Error GoTo txtMonoSpot_KeyPress_Error

20        cmdSave.Visible = True

30        Exit Sub

txtMonoSpot_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmBarCodes", "txtMonoSpot_KeyPress", intEL, strES


End Sub

Private Sub txtRandom_KeyPress(KeyAscii As Integer)

10        On Error GoTo txtRandom_KeyPress_Error

20        cmdSave.Visible = True

30        Exit Sub

txtRandom_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmBarCodes", "txtRandom_KeyPress", intEL, strES


End Sub

Private Sub txtRetics_KeyPress(KeyAscii As Integer)

10        On Error GoTo txtRetics_KeyPress_Error

20        cmdSave.Visible = True

30        Exit Sub

txtRetics_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmBarCodes", "txtRetics_KeyPress", intEL, strES


End Sub

Private Sub txtSave_KeyPress(KeyAscii As Integer)

10        On Error GoTo txtSave_KeyPress_Error

20        cmdSave.Visible = True

30        Exit Sub

txtSave_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmBarCodes", "txtSave_KeyPress", intEL, strES


End Sub
