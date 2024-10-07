VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmViewIQ200Repeat 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - IQ200 Repeats"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7245
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   7245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   1065
      Left            =   5775
      Picture         =   "frmViewIQ200Repeat.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmdTransfer 
      Caption         =   "Copy to Main &File"
      Height          =   1065
      Left            =   5775
      Picture         =   "frmViewIQ200Repeat.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   435
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete All Repeats"
      Height          =   1065
      Left            =   5775
      Picture         =   "frmViewIQ200Repeat.frx":11D4
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1740
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   4965
      Left            =   90
      TabIndex        =   4
      Top             =   450
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   8758
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      AllowUserResizing=   1
      FormatString    =   "<Test        |<Name                         |<Result                |<Date/Time          |Counter"
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tests in RED will be Transfered"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   180
      Width           =   3705
   End
End
Attribute VB_Name = "frmViewIQ200Repeat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Activated As Boolean

Private pSampleID As Double

Private Sub cmdCancel_Click()

10        Unload Me

End Sub


Private Sub cmdDelete_Click()

          Dim sql As String

10        On Error GoTo cmdDelete_Click_Error

20        If iMsg("DELETE All Repeats?" & vbCrLf & _
                  "You will not be able to undo this process!" & vbCrLf & _
                  "Continue?", vbQuestion + vbYesNo) = vbYes Then

30            sql = "DELETE from IQ200Repeats WHERE " & _
                    "SampleID = '" & pSampleID & "'"

40            Cnxn(0).Execute sql

50            Unload Me

60        End If

70        Exit Sub

cmdDelete_Click_Error:

          Dim strES As String
          Dim intEL As Integer

80        intEL = Erl
90        strES = Err.Description
100       LogError "frmViewIQ200Repeat", "cmdDelete_Click", intEL, strES, sql

End Sub

Private Sub cmdTransfer_Click()

          Dim Y As Long
          Dim sql As String
          Dim Deleted As Boolean

10        On Error GoTo cmdTransfer_Click_Error

20        Deleted = False

30        g.Col = 0
40        For Y = 1 To g.Rows - 1
50            g.Row = Y
60            If g.CellBackColor = vbRed Then
70                If Not Deleted Then
80                    sql = "DELETE FROM IQ200 " & _
                            "WHERE SampleID = " & pSampleID
90                    Cnxn(0).Execute sql
100                   Deleted = True
110               End If
120               sql = "INSERT INTO IQ200 " & _
                        "(SampleID, TestCode, ShortName, LongName, Range, Result, WorkListPrinted, DateTimeOfRecord, " & _
                        "Validated, ValidatedBy, Printed, PrintedBy) " & _
                        "SELECT SampleID, TestCode, ShortName, LongName, Range, Result, WorkListPrinted, DateTimeOfRecord, " & _
                        "Validated, ValidatedBy, Printed, PrintedBy FROM IQ200Repeats " & _
                        "WHERE SampleID = " & pSampleID & " " & _
                        "AND Counter = '" & g.TextMatrix(Y, 4) & "'"
130               Cnxn(0).Execute sql


                  '    sql = "IF EXISTS (SELECT * FROM IQ200 " & _
                       '          "           WHERE SampleID = '" & pSampleID & "' " & _
                       '          "           AND TestCode = '" & g.TextMatrix(y, 0) & "') " & _
                       '          "  UPDATE IQ200 " & _
                       '          "  SET Result = '" & g.TextMatrix(y, 2) & "', " & _
                       '          "  WorklistPrinted = 0, " & _
                       '          "  DateTimeOfRecord = '" & Format$(g.TextMatrix(y, 3), "dd/MMM/yyyy HH:nn") & "', " & _
                       '          "  Validated = 0, ValidatedBy = '', Printed = 0, PrintedBy = '' " & _
                       '          "  WHERE SampleID = '" & pSampleID & "' " & _
                       '          "  AND TestCode = '" & g.TextMatrix(y, 0) & "' " & _
                       '          "ELSE " & _
                       '          "  INSERT INTO IQ200 " & _
                       '          "  (SampleID, TestCode, ShortName, LongName, Range, Result, WorkListPrinted, DateTimeOfRecord, " & _
                       '          "  Validated, ValidatedBy, Printed, PrintedBy) " & _
                       '          "  SELECT SampleID, TestCode, ShortName, LongName, Range, Result, WorkListPrinted, DateTimeOfRecord, " & _
                       '          "  Validated, ValidatedBy, Printed, PrintedBy FROM IQ200Repeats " & _
                       '          "  WHERE SampleID = '" & pSampleID & "' " & _
                       '          "  AND Counter = '" & g.TextMatrix(y, 4) & "'"
                  '    Cnxn(0).Execute sql
                  '
140               sql = "DELETE FROM IQ200Repeats " & _
                        "WHERE SampleID = " & pSampleID & " " & _
                        "AND Counter = '" & g.TextMatrix(Y, 4) & "'"
150               Cnxn(0).Execute sql
160           End If
170       Next

          'FillG

180       Unload Me

190       Exit Sub

cmdTransfer_Click_Error:

          Dim strES As String
          Dim intEL As Integer

200       intEL = Erl
210       strES = Err.Description
220       LogError "frmViewIQ200Repeat", "cmdTransfer_Click", intEL, strES, sql

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
80        LogError "frmViewIQ200Repeat", "Form_Activate", intEL, strES

End Sub
Private Sub FillG()

          Dim sql As String
          Dim tb As Recordset
          Dim s As String

10        On Error GoTo FillG_Error

20        With g
30            .Rows = 2
40            .AddItem ""
50            .RemoveItem 1
60        End With

70        sql = "SELECT * FROM IQ200Repeats " & _
                "WHERE Sampleid = '" & pSampleID & "' AND Result <> '[none]'"
80        Set tb = New Recordset
90        RecOpenServer 0, tb, sql
100       Do While Not tb.EOF
110           s = tb!TestCode & vbTab & _
                  tb!LongName & vbTab & _
                  tb!Result & vbTab & _
                  Format$(tb!DateTimeOfRecord, "dd/MM/yy HH:nn") & vbTab & _
                  tb!Counter
120           g.AddItem s
130           tb.MoveNext
140       Loop

150       If g.Rows > 2 Then
160           g.RemoveItem 1
170       End If

180       Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

190       intEL = Erl
200       strES = Err.Description
210       LogError "frmViewIQ200Repeat", "FillG", intEL, strES, sql

End Sub

Private Sub Form_Load()

10        Activated = False
20        g.ColWidth(4) = 0

End Sub


Private Sub Form_Unload(Cancel As Integer)

10        Activated = False

End Sub



Public Property Let SampleID(ByVal NewValue As Double)

10        pSampleID = NewValue

End Property

Private Sub g_Click()

          Dim Y As Long
          Dim TimeOfRecord As String

10        On Error GoTo g_Click_Error

20        If g.MouseRow = 0 Then Exit Sub

30        cmdTransfer.Visible = False

40        g.Row = g.MouseRow
50        TimeOfRecord = g.TextMatrix(g.Row, 3)

60        g.Col = 0
70        If g.CellBackColor = vbRed Then
80            For Y = 1 To g.Rows - 1
90                g.Row = Y
100               g.CellBackColor = 0
110           Next
120       Else
130           For Y = 1 To g.Rows - 1
140               g.Row = Y
150               g.CellBackColor = 0
160           Next
170           For Y = 1 To g.Rows - 1
180               If g.TextMatrix(Y, 3) = TimeOfRecord Then
190                   g.Row = Y
200                   g.CellBackColor = vbRed
210               End If
220           Next
230       End If

240       For Y = 1 To g.Rows - 1
250           g.Row = Y
260           If g.CellBackColor = vbRed Then
270               cmdTransfer.Visible = True
280               Exit For
290           End If
300       Next

310       Exit Sub

g_Click_Error:

          Dim strES As String
          Dim intEL As Integer

320       intEL = Erl
330       strES = Err.Description
340       LogError "frmViewIQ200Repeat", "g_Click", intEL, strES

End Sub


