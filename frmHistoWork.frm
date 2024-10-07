VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmHistoWork 
   Caption         =   "NetAcquire - Histology Work List"
   ClientHeight    =   6345
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9630
   Icon            =   "frmHistoWork.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   9630
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   765
      Left            =   8100
      Picture         =   "frmHistoWork.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3600
      Width           =   1245
   End
   Begin VB.OptionButton opt 
      Caption         =   "Stain"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   4530
      TabIndex        =   5
      Top             =   450
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.OptionButton opt 
      Caption         =   "Pieces"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   3240
      TabIndex        =   4
      Top             =   450
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.CommandButton bPrint 
      Caption         =   "Print"
      Height          =   720
      Left            =   8100
      Picture         =   "frmHistoWork.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4500
      Width           =   1245
   End
   Begin VB.CommandButton bcancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   720
      Left            =   8100
      Picture         =   "frmHistoWork.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5355
      Width           =   1245
   End
   Begin MSComCtl2.DTPicker dtRundate 
      Height          =   285
      Left            =   480
      TabIndex        =   1
      Top             =   480
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "dd/MMM/yyyy"
      Format          =   59179011
      CurrentDate     =   38028
   End
   Begin MSFlexGridLib.MSFlexGrid grdHist 
      Height          =   5085
      Left            =   420
      TabIndex        =   0
      Top             =   1020
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   8969
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      FormatString    =   "Sample No  |Specimen                                               |Blocks| Pieces | Correct|Sampleid|tYPE"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid grdComm 
      Height          =   5085
      Left            =   390
      TabIndex        =   6
      Top             =   1020
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   8969
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      FormatString    =   "Sample No  |Specimen                                               |Blocks| Pieces|              "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmHistoWork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub bcancel_Click()

10        Unload Me

End Sub

Private Sub bprint_Click()

10        On Error GoTo bprint_Click_Error

20        If opt(0) Then
30            Print_Piece
40        Else
50            Print_Stain
60        End If

70        Exit Sub

bprint_Click_Error:

          Dim strES As String
          Dim intEL As Integer



80        intEL = Erl
90        strES = Err.Description
100       LogError "frmHistoWork", "bPrint_Click", intEL, strES


End Sub

Private Sub cmdUpdate_Click()
          Dim sql As String
          Dim n As Long
          Dim s As Long


10        On Error GoTo cmdUpdate_Click_Error

20        For n = 1 To grdHist.Rows - 1
30            If grdHist.TextMatrix(n, 0) <> "" Then
40                If grdHist.TextMatrix(n, 4) = "Y" Then s = 1 Else s = 0
50                sql = "Update histoblock set checked = " & s & " where " & _
                        "sampleid = " & grdHist.TextMatrix(n, 5) & " " & _
                        "and block = '" & Right(grdHist.TextMatrix(n, 2), 1) & "' " & _
                        "and type = '" & grdHist.TextMatrix(n, 6) & "'"
60                Cnxn(0).Execute sql
70            End If
80        Next

90        cmdUpdate.Enabled = False

100       Exit Sub

cmdUpdate_Click_Error:

          Dim strES As String
          Dim intEL As Integer



110       intEL = Erl
120       strES = Err.Description
130       LogError "frmHistoWork", "cmdUpdate_Click", intEL, strES, sql


End Sub

Private Sub dtRunDate_CloseUp()

10        On Error GoTo dtRunDate_CloseUp_Error

20        FillG

30        Exit Sub

dtRunDate_CloseUp_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmHistoWork", "dtRunDate_CloseUp", intEL, strES


End Sub

Private Sub Fill_Piece()

          Dim tb As New Recordset
          Dim sql As String
          Dim VarSamp As String
          Dim VarType As String
          Dim s As String
          Dim sn As New Recordset
          Dim n As Long
          Dim v As Long
          Dim Lett As String
          Dim D As String

10        On Error GoTo Fill_Piece_Error

20        VarSamp = ""

30        With grdHist
40            .Rows = 2
50            .AddItem ""
60            .RemoveItem 1
70            .TextMatrix(0, 3) = "Pieces"
80            .ColWidth(3) = 1000
90        End With

100       With grdComm
110           .Rows = 2
120           .AddItem ""
130           .RemoveItem 1
140           .TextMatrix(0, 3) = "Pieces"
150           .ColWidth(3) = 1000
160       End With

170       sql = "SELECT * from histospecimen WHERE rundate = '" & Format(dtRunDate, "dd/MMM/yyyy") & "'"

180       Set tb = New Recordset
190       RecOpenServer 0, tb, sql

200       Do While Not tb.EOF
210           If VarSamp <> tb!SampleID Or VarType <> tb!Type Then
220               If VarSamp <> tb!SampleID And VarSamp <> "" Then grdHist.AddItem ""

230               VarSamp = tb!SampleID
240               VarType = tb!Type

250               s = Trim(tb!Hyear) & "/" & Val(tb!SampleID) - (SysOptHistoOffset(0) + (Val(Swap_Year(Trim(tb!Hyear))) * 1000)) & "H" & vbTab & Trim(tb!Type) & vbTab
260               Lett = FindLetter(tb!specimen)
270           Else
280               Lett = FindLetter(tb!specimen)
290               s = vbTab & vbTab
300           End If

310           For n = 1 To Val(Trim(tb!blocks))
320               sql = "SELECT * from histoblock WHERE sampleid = '" & tb!SampleID & "' and type = '" & tb!Type & "' and block = '" & n & "'"
330               Set sn = New Recordset
340               RecOpenServer 0, sn, sql
350               If Not sn.EOF Then
360                   If sn!Checked = 1 Then D = "Y" Else D = "N"
370                   s = s & Lett & n & vbTab & sn!pieces & vbTab & D & vbTab & VarSamp & vbTab & VarType
380               End If
390               grdHist.AddItem s
400               For v = 0 To 1
410                   grdHist.Row = grdHist.Rows - 1
420                   grdHist.Col = v
430                   If grdHist.TextMatrix(grdHist.Row, 0) <> "" Then grdHist.CellBackColor = vbCyan
440               Next
450               If Not sn.EOF Then
460                   If Trim(sn!picomm & "") <> "" Then
470                       grdHist.Row = grdHist.Rows - 1
480                       grdHist.Col = 3
490                       grdHist.CellBackColor = vbYellow
500                       s = vbTab & vbTab & vbTab
510                       If sn!Checked = 1 Then D = "Y" Else D = "N"
520                       s = s & D & vbTab & sn!SampleID
530                       grdComm.AddItem s
540                   Else
550                       grdComm.AddItem ""
560                   End If
570               End If
580               s = vbTab & vbTab
590           Next

600           tb.MoveNext
610       Loop

620       If grdHist.Rows > 2 And grdHist.TextMatrix(1, 0) = "" Then
630           grdHist.RemoveItem 1
640       End If

650       If grdComm.Rows > 2 And grdComm.TextMatrix(1, 0) = "" Then
660           grdComm.RemoveItem 1
670       End If

680       Exit Sub

Fill_Piece_Error:

          Dim strES As String
          Dim intEL As Integer



690       intEL = Erl
700       strES = Err.Description
710       LogError "frmHistoWork", "Fill_Piece", intEL, strES, sql

End Sub

Private Sub FillG()

10        On Error GoTo FillG_Error

20        If opt(0) Then
30            Fill_Piece
40        Else
50            Fill_Stain
60        End If

70        Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer



80        intEL = Erl
90        strES = Err.Description
100       LogError "frmHistoWork", "FillG", intEL, strES


End Sub

Private Function FindLetter(ByVal sptype As String)

10        On Error GoTo FindLetter_Error

20        Select Case sptype
          Case "0"
30            FindLetter = "A"
40        Case "1"
50            FindLetter = "B"
60        Case "2"
70            FindLetter = "C"
80        Case "3"
90            FindLetter = "D"
100       Case "4"
110           FindLetter = "E"
120       Case "5"
130           FindLetter = "F"
140       End Select

150       Exit Function

FindLetter_Error:

          Dim strES As String
          Dim intEL As Integer



160       intEL = Erl
170       strES = Err.Description
180       LogError "frmHistoWork", "FindLetter", intEL, strES

End Function

Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        grdHist.ColWidth(5) = 0
30        grdHist.ColWidth(6) = 0

40        dtRunDate = Format(Now, "dd/MMM/yyyy")
50        FillG

60        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer



70        intEL = Erl
80        strES = Err.Description
90        LogError "frmHistoWork", "Form_Load", intEL, strES


End Sub



Private Sub grdHist_Click()


10        On Error GoTo grdHist_Click_Error

20        If grdHist.RowSel > 0 Then
30            If grdHist.ColSel = 4 Then
40                Select Case Trim(grdHist.TextMatrix(grdHist.RowSel, grdHist.ColSel))
                  Case "": grdHist.TextMatrix(grdHist.RowSel, grdHist.ColSel) = "Y"
50                Case "Y": grdHist.TextMatrix(grdHist.RowSel, grdHist.ColSel) = "N"
60                Case "N": grdHist.TextMatrix(grdHist.RowSel, grdHist.ColSel) = ""
70                End Select
80            End If
90            cmdUpdate.Enabled = True
100       End If

110       Exit Sub

grdHist_Click_Error:

          Dim strES As String
          Dim intEL As Integer



120       intEL = Erl
130       strES = Err.Description
140       LogError "frmHistoWork", "grdHist_Click", intEL, strES


End Sub

Private Sub grdHist_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo grdHist_MouseMove_Error

20        grdHist.Col = grdHist.MouseCol
30        grdHist.Row = grdHist.MouseRow

40        If grdHist.MouseCol > 3 Then Exit Sub

50        grdHist.ToolTipText = ""
60        If grdHist.CellBackColor = vbYellow Then
70            grdHist.ToolTipText = grdComm.TextMatrix(grdHist.MouseRow, grdHist.MouseCol)
80        End If

90        Exit Sub

grdHist_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer



100       intEL = Erl
110       strES = Err.Description
120       LogError "frmHistoWork", "grdHist_MouseMove", intEL, strES


End Sub

Private Sub opt_Click(Index As Integer)

10        On Error GoTo opt_Click_Error

20        FillG

30        Exit Sub

opt_Click_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmHistoWork", "opt_Click", intEL, strES


End Sub

Private Sub Print_Piece()

          Dim n As Long
          Dim Pgn As Long

10        On Error GoTo Print_Piece_Error

20        Pgn = 1

30        Printer.Font.Size = 16
40        Printer.Print Tab(15); "Histology Work List for " & dtRunDate & "."
50        Printer.Print
60        Printer.Font.Size = 10

70        Printer.Print grdHist.TextMatrix(0, 0);
80        Printer.Print Tab(17); grdHist.TextMatrix(0, 1);
90        Printer.Print Tab(67); grdHist.TextMatrix(0, 2);
100       Printer.Print Tab(77); grdHist.TextMatrix(0, 3);
110       Printer.Print Tab(90); "Checked"

120       For n = 1 To grdHist.Rows - 1
130           If grdHist.TextMatrix(n, 2) <> "" Then
140               Printer.Print grdHist.TextMatrix(n, 0);
150               If InStr(grdHist.TextMatrix(n, 1), "Block") = 0 Then
160                   Printer.Font.Bold = True
170                   Printer.Print Tab(16); StrConv(grdHist.TextMatrix(n, 1), vbProperCase);
180                   Printer.Print Tab(61); grdHist.TextMatrix(n, 2);
190                   Printer.Font.Bold = False
200               Else
210                   Printer.Print Tab(47); StrConv(grdHist.TextMatrix(n, 1), vbProperCase);
220                   Printer.Print Tab(67); grdHist.TextMatrix(n, 2);
230               End If
240               Printer.Print Tab(77); grdHist.TextMatrix(n, 3);
250               If grdHist.TextMatrix(n, 3) <> "" Then Printer.Print Tab(93); "___" Else Printer.Print
260           Else
270               Printer.Print
280           End If
290           If Printer.CurrentY = 16000 Then
300               Printer.Print "Page " & n
310               Pgn = Pgn + 1
320               Printer.NewPage
330               Printer.Print grdHist.TextMatrix(0, 0);
340               Printer.Print Tab(17); grdHist.TextMatrix(0, 1);
350               Printer.Print Tab(67); grdHist.TextMatrix(0, 2);
360               Printer.Print Tab(77); grdHist.TextMatrix(0, 3);
370               Printer.Print Tab(90); "Checked"
380           End If
390       Next
400       Printer.Print
410       Printer.Font.Size = 16
420       Printer.Print Tab(25); "End of Worklist"
430       Printer.EndDoc

440       Exit Sub

Print_Piece_Error:

          Dim strES As String
          Dim intEL As Integer



450       intEL = Erl
460       strES = Err.Description
470       LogError "frmHistoWork", "Print_Piece", intEL, strES


End Sub

Private Sub Fill_Stain()
          Dim tb As New Recordset
          Dim sql As String
          Dim VarSamp As String
          Dim VarType As String
          Dim s As String
          Dim sn As New Recordset
          Dim n As Long

10        On Error GoTo Fill_Stain_Error

20        VarSamp = ""

30        With grdHist
40            .Rows = 2
50            .AddItem ""
60            .RemoveItem 1
70            .TextMatrix(0, 3) = "Stain"
80            .ColWidth(3) = 2500
90        End With


100       sql = "SELECT * from histospecimen WHERE rundate = '" & Format(dtRunDate, "dd/MMM/yyyy") & "'"

110       Set tb = New Recordset
120       RecOpenServer 0, tb, sql

130       Do While Not tb.EOF
140           If VarSamp <> tb!SampleID Or VarType <> tb!Type Then
150               If VarSamp <> tb!SampleID And VarSamp <> "" Then grdHist.AddItem ""

160               VarSamp = tb!SampleID
170               VarType = tb!Type

180               s = tb!SampleID & vbTab & Trim(tb!Type) & vbTab & FindLetter(tb!specimen)
190           Else
200               s = vbTab & vbTab & FindLetter(tb!specimen)
210           End If
220           grdHist.AddItem s
230           For n = 0 To 2
240               grdHist.Row = grdHist.Rows - 1
250               grdHist.Col = n
260               grdHist.CellBackColor = vbCyan
270           Next

280           For n = 1 To tb!blocks
290               sql = "SELECT STAIN from histostain WHERE sampleid = '" & tb!SampleID & "' and block = '" & n & "'"
300               Set sn = New Recordset
310               RecOpenServer 0, sn, sql
320               Do While Not sn.EOF
330                   s = vbTab & "Block Number" & vbTab
340                   s = s & n & vbTab & sn!stain
350                   grdHist.AddItem s
360                   sn.MoveNext
370               Loop
380           Next
390           tb.MoveNext
400       Loop

410       If grdHist.Rows > 2 And grdHist.TextMatrix(1, 0) = "" Then
420           grdHist.RemoveItem 1
430       End If

440       Exit Sub

Fill_Stain_Error:

          Dim strES As String
          Dim intEL As Integer



450       intEL = Erl
460       strES = Err.Description
470       LogError "frmHistoWork", "Fill_Stain", intEL, strES, sql


End Sub

Private Sub Print_Stain()

          Dim n As Long

10        On Error GoTo Print_Stain_Error

20        Printer.Font.Size = 16
30        Printer.Print Tab(15); "Histology Pieces Work List for " & dtRunDate & "."
40        Printer.Print
50        Printer.Font.Size = 10

60        Printer.Print grdHist.TextMatrix(0, 0);
70        Printer.Print Tab(12); grdHist.TextMatrix(0, 1);
80        Printer.Print Tab(62); grdHist.TextMatrix(0, 2);
90        Printer.Print Tab(72); grdHist.TextMatrix(0, 3);
100       Printer.Print Tab(77); "Checked"

110       For n = 1 To grdHist.Rows - 1
120           Printer.Print grdHist.TextMatrix(n, 0);
130           Printer.Print Tab(12); grdHist.TextMatrix(n, 1);
140           Printer.Print Tab(62); grdHist.TextMatrix(n, 2);
150           Printer.Print Tab(72); grdHist.TextMatrix(n, 3);
160           Printer.Print Tab(82); "___"
170           If Printer.CurrentY = 16000 Then
180               Printer.NewPage
190               Printer.Print grdHist.TextMatrix(0, 0);
200               Printer.Print Tab(12); grdHist.TextMatrix(0, 1);
210               Printer.Print Tab(62); grdHist.TextMatrix(0, 2);
220               Printer.Print Tab(72); grdHist.TextMatrix(0, 3);
230               Printer.Print Tab(77); "Checked"
240           End If
250       Next
260       Printer.Print
270       Printer.Print Tab(10); "End of Worklist"

280       Printer.EndDoc

290       Exit Sub

Print_Stain_Error:

          Dim strES As String
          Dim intEL As Integer



300       intEL = Erl
310       strES = Err.Description
320       LogError "frmHistoWork", "Print_Stain", intEL, strES


End Sub

