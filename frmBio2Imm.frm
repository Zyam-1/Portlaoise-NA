VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmBio2Imm 
   Caption         =   "NetAcquire - Bio Related Immunology Results"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmBio2Imm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Exit"
      Height          =   1050
      Left            =   3450
      Picture         =   "frmBio2Imm.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5160
      Width           =   960
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy to Immunology"
      Height          =   1050
      Left            =   3450
      Picture         =   "frmBio2Imm.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   570
      Width           =   960
   End
   Begin MSFlexGridLib.MSFlexGrid grdB2I 
      Height          =   6030
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   3030
      _ExtentX        =   5345
      _ExtentY        =   10636
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      FormatString    =   "Test Name           | Result    "
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
End
Attribute VB_Name = "frmBio2Imm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

10        Unload Me

End Sub

Private Sub cmdCopy_Click()

          Dim sql As String
          Dim tb As New Recordset
          Dim sn As New Recordset
          Dim n As Long
          Dim Code As String

10        On Error GoTo cmdCopy_Click_Error

20        For n = 1 To grdB2I.Rows - 1
30            grdB2I.Row = n
40            If grdB2I.CellBackColor = vbRed Then
50                Code = CodeForShortName(grdB2I.TextMatrix(n, 0))

60                sql = "SELECT * FROM BioResults WHERE " & _
                        "SampleID = '" & frmEditAll.txtSampleID & "' " & _
                        "AND Code = '" & Code & "'"
70                Set tb = New Recordset
80                RecOpenServer 0, tb, sql
90                If Not tb.EOF Then
                      '90          Code = ICodeForShortName(grdB2I.TextMatrix(n, 0))
                      '100         If Code <> "???" Then
100                   sql = "SELECT * FROM ImmResults WHERE " & _
                            "SampleID = '" & frmEditAll.txtSampleID & "' " & _
                            "AND Code = '" & Code & "'"
110                   Set sn = New Recordset
120                   RecOpenServer 0, sn, sql
130                   If sn.EOF Then
140                       sn.AddNew
150                   Else
160                       sql = "SELECT * FROM ImmRepeats WHERE " & _
                                "SampleID = '" & frmEditAll.txtSampleID & "' " & _
                                "AND Code = '" & Code & "'"
170                       Set sn = New Recordset
180                       RecOpenServer 0, sn, sql
190                       If sn.EOF Then
200                           sn.AddNew
210                       End If
220                   End If
230                   sn!SampleID = tb!SampleID
240                   sn!Code = tb!Code
250                   sn!Result = tb!Result
260                   sn!Valid = tb!Valid
270                   sn!Printed = 0
280                   sn!Rundate = Format(tb!RunTime, "dd/MMM/yyyy")
290                   sn!RunTime = Format(tb!RunTime, "dd/MMM/yyyy hh:mm:ss")
300                   sn!Operator = tb!Operator
310                   sn!Units = tb!Units
320                   sn!SampleType = tb!SampleType
330                   sn.Update
                      '350         End If
340               End If
350           End If
360       Next

370       Unload Me

380       Exit Sub

cmdCopy_Click_Error:

          Dim strES As String
          Dim intEL As Integer

390       intEL = Erl
400       strES = Err.Description
410       LogError "frmBio2Imm", "cmdCopy_Click", intEL, strES, sql

End Sub

Private Sub Form_Load()

          Dim n As Long
          Dim s As String

10        On Error GoTo Form_Load_Error

20        With grdB2I
30            .Rows = 2
40            .AddItem ""
50            .RemoveItem 1
60        End With

70        For n = 1 To frmEditAll.gBio.Rows - 1
80            If InStr(frmEditAll.gBio.TextMatrix(n, 6), "V") > 0 Then
90                s = frmEditAll.gBio.TextMatrix(n, 0) & vbTab & frmEditAll.gBio.TextMatrix(n, 1)
100               grdB2I.AddItem s
110           End If
120       Next

130       If grdB2I.Rows > 2 And grdB2I.TextMatrix(1, 0) = "" Then
140           grdB2I.RemoveItem 1
150       End If

160       Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

170       intEL = Erl
180       strES = Err.Description
190       LogError "frmBio2Imm", "Form_Load", intEL, strES

End Sub

Private Sub grdB2I_Click()

          Dim n As Long

10        On Error GoTo grdB2I_Click_Error

20        If grdB2I.CellBackColor = vbRed Then
30            For n = 0 To 1
40                grdB2I.Col = n
50                grdB2I.CellBackColor = vbWhite
60            Next
70        Else
80            For n = 0 To 1
90                grdB2I.Col = n
100               grdB2I.CellBackColor = vbRed
110           Next
120       End If

130       Exit Sub

grdB2I_Click_Error:

          Dim strES As String
          Dim intEL As Integer

140       intEL = Erl
150       strES = Err.Description
160       LogError "frmBio2Imm", "grdB2I_Click", intEL, strES

End Sub
