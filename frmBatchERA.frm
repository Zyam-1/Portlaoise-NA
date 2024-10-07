VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmBatchERA 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire -  Batch Entry"
   ClientHeight    =   8505
   ClientLeft      =   90
   ClientTop       =   510
   ClientWidth     =   13860
   ControlBox      =   0   'False
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
   Icon            =   "frmBatchERA.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8505
   ScaleWidth      =   13860
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   900
      Left            =   9360
      Picture         =   "frmBatchERA.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   585
      Width           =   1100
   End
   Begin VB.OptionButton optUnResulted 
      Caption         =   "Show Unresulted Requests"
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   180
      TabIndex        =   14
      Top             =   1050
      Value           =   -1  'True
      Width           =   2985
   End
   Begin VB.OptionButton optNotValid 
      Caption         =   "Show Unvalidated Results"
      ForeColor       =   &H00008000&
      Height          =   195
      Left            =   180
      TabIndex        =   13
      Top             =   1290
      Width           =   2985
   End
   Begin VB.CommandButton bsearch 
      Caption         =   "&Search"
      Default         =   -1  'True
      Height          =   900
      Left            =   3480
      Picture         =   "frmBatchERA.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   180
      Width           =   1100
   End
   Begin VB.CommandButton cmdSave 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
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
      Height          =   900
      Left            =   11370
      Picture         =   "frmBatchERA.frx":0A56
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   585
      Width           =   1100
   End
   Begin VB.CommandButton cmdSet 
      Appearance      =   0  'Flat
      Caption         =   "Set All - Not Detected"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   2
      Left            =   6660
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2040
      Width           =   1845
   End
   Begin VB.CommandButton cmdSet 
      Appearance      =   0  'Flat
      Caption         =   "Set All - Negative"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   3
      Left            =   10890
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2040
      Width           =   1845
   End
   Begin VB.CommandButton cmdSet 
      Appearance      =   0  'Flat
      Caption         =   "Set All Negative"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   2880
      TabIndex        =   3
      Top             =   2040
      Width           =   1365
   End
   Begin VB.CommandButton cmdSet 
      Appearance      =   0  'Flat
      Caption         =   "Set All Negative"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   1380
      TabIndex        =   2
      Top             =   2040
      Width           =   1425
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
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
      Height          =   900
      Left            =   12510
      Picture         =   "frmBatchERA.frx":0D60
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   585
      Width           =   1100
   End
   Begin MSFlexGridLib.MSFlexGrid grdBatch 
      Height          =   5955
      Index           =   2
      Left            =   5490
      TabIndex        =   6
      Top             =   2385
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   10504
      _Version        =   393216
      Cols            =   3
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      HighLight       =   0
      GridLines       =   3
      ScrollBars      =   2
      FormatString    =   "<Sample ID  |<Result                     |^Valid  "
   End
   Begin MSFlexGridLib.MSFlexGrid grdBatch 
      Height          =   5955
      Index           =   3
      Left            =   9690
      TabIndex        =   7
      Top             =   2385
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   10504
      _Version        =   393216
      Cols            =   3
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      ScrollBars      =   2
      FormatString    =   "<Sample ID  |<Result                     |^Valid  "
   End
   Begin MSFlexGridLib.MSFlexGrid grdBatch 
      Height          =   5955
      Index           =   0
      Left            =   195
      TabIndex        =   8
      Top             =   2385
      Width           =   5085
      _ExtentX        =   8969
      _ExtentY        =   10504
      _Version        =   393216
      Cols            =   4
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      HighLight       =   0
      GridLines       =   3
      ScrollBars      =   2
      FormatString    =   "<Sample ID  |<Rota Result      |<Adeno Result     |^Valid  "
   End
   Begin MSComCtl2.DTPicker dtStart 
      Height          =   315
      Left            =   1650
      TabIndex        =   15
      Top             =   180
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   59244545
      CurrentDate     =   40267
   End
   Begin MSComCtl2.DTPicker dtEnd 
      Height          =   315
      Left            =   1650
      TabIndex        =   16
      Top             =   600
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   59244545
      CurrentDate     =   40267
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      Height          =   345
      Left            =   9360
      TabIndex        =   20
      Top             =   180
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Filter:         From"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   180
      TabIndex        =   18
      Top             =   210
      Width           =   1350
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1260
      TabIndex        =   17
      Top             =   630
      Width           =   255
   End
   Begin VB.Image imgSquareCross 
      Height          =   225
      Left            =   12540
      Picture         =   "frmBatchERA.frx":106A
      Top             =   60
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgSquareTick 
      Height          =   225
      Left            =   11760
      Picture         =   "frmBatchERA.frx":1340
      Top             =   120
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label lblHeading 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "H.Pylori"
      Height          =   255
      Index           =   3
      Left            =   9660
      TabIndex        =   11
      Top             =   1680
      Width           =   4005
   End
   Begin VB.Label lblHeading 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "c.Difficile"
      Height          =   255
      Index           =   2
      Left            =   5490
      TabIndex        =   10
      Top             =   1680
      Width           =   4005
   End
   Begin VB.Label lblHeading 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Rota / Adeno Virus"
      Height          =   255
      Index           =   0
      Left            =   195
      TabIndex        =   1
      Top             =   1680
      Width           =   5145
   End
End
Attribute VB_Name = "frmBatchERA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim ListCDiffToxinAB() As ListColour
Dim ListRota() As ListColour
Dim ListAdeno() As ListColour
Dim ListHPylori() As ListColour

Private Sub bsearch_Click()
10        FillG "A"
20        FillG "G"
30        FillG "Y"
End Sub

Private Sub cmdCancel_Click()

10        Unload Me

End Sub

Private Sub cmdSave_Click()

          Dim tb As Recordset
          Dim sql As String
          Dim Num As Integer
          Dim Y As Integer
          Dim strDept As String
          Dim strDeptFull As String

10        On Error GoTo cmdSave_Click_Error

20        For Num = 0 To 3
30            strDept = Choose(Num + 1, "A", "A", "G", "Y")
40            strDeptFull = Choose(Num + 1, "ROTAADENO", "ROTAADENO", "CDIFF", "HPYLORI")
50            If Num <> 1 Then
60                For Y = 1 To grdBatch(Num).Rows - 1
70                    If grdBatch(Num).TextMatrix(Y, 0) <> "" And grdBatch(Num).TextMatrix(Y, 1) <> "?" Then
80                        If ValidStatus4MicroDept(Val(SysOptMicroOffset(0) + grdBatch(Num).TextMatrix(Y, 0)), strDept) = False Then
90                            sql = "SELECT * FROM Faeces WHERE " & _
                                    "SampleID = '" & SysOptMicroOffset(0) + grdBatch(Num).TextMatrix(Y, 0) & "'"
100                           Set tb = New Recordset
110                           RecOpenServer 0, tb, sql
120                           If tb.EOF Then
130                               tb.AddNew
140                               tb!SampleID = SysOptMicroOffset(0) + grdBatch(Num).TextMatrix(Y, 0)
150                           End If
160                           If Num = 3 Then    'H Pylori
170                               grdBatch(Num).Row = Y
180                               grdBatch(Num).Col = 1
190                               tb!HPylori = grdBatch(Num).TextMatrix(Y, 1) & "|" & grdBatch(Num).CellForeColor & "|" & grdBatch(Num).CellBackColor
200                           ElseIf Num = 0 Then     'Roda/Ademo
210                               grdBatch(Num).Row = Y
220                               grdBatch(Num).Col = 1
230                               tb!Rota = grdBatch(Num).TextMatrix(Y, 1) & "|" & grdBatch(Num).CellForeColor & "|" & grdBatch(Num).CellBackColor
240                               grdBatch(Num).Row = Y
250                               grdBatch(Num).Col = 2
260                               tb!Adeno = grdBatch(Num).TextMatrix(Y, 2) & "|" & grdBatch(Num).CellForeColor & "|" & grdBatch(Num).CellBackColor
270                           ElseIf Num = 2 Then    'C Diff
280                               grdBatch(Num).Row = Y
290                               grdBatch(Num).Col = 1
300                               tb!ToxinAB = grdBatch(Num).TextMatrix(Y, 1) & "|" & grdBatch(Num).CellForeColor & "|" & grdBatch(Num).CellBackColor
310                           End If
320                           tb!Username = Username
330                           tb.Update
340                       End If
350                   End If
360                   If Num = 0 Then
370                       grdBatch(Num).Row = Y
380                       grdBatch(Num).Col = 3
390                   Else
400                       grdBatch(Num).Row = Y
410                       grdBatch(Num).Col = 2
420                   End If
430                   If grdBatch(Num).CellPicture = imgSquareTick.Picture Then
440                       UpdatePrintValidLog SysOptMicroOffset(0) + grdBatch(Num).TextMatrix(Y, 0), strDeptFull, 1, 0
450                   ElseIf grdBatch(Num).CellPicture = imgSquareCross.Picture Then
460                       UpdatePrintValidLog SysOptMicroOffset(0) + grdBatch(Num).TextMatrix(Y, 0), strDeptFull, 0, 0
470                   End If
480               Next
490           End If
500       Next
510       cmdSave.Enabled = False
520       FillG "A"
530       FillG "G"
540       FillG "Y"

550       Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

560       intEL = Erl
570       strES = Err.Description
580       LogError "frmBatchERA", "cmdsave_Click", intEL, strES, sql

End Sub

Private Sub cmdSet_Click(Index As Integer)
          Dim Num As Long

10        On Error GoTo cmdSet_Click_Error

20        If grdBatch(Index).Rows = 1 And grdBatch(Index).TextMatrix(0, 0) = "" Then Exit Sub

30        For Num = 1 To grdBatch(Index).Rows - 1
40            If Index = 2 Then
50                grdBatch(Index).TextMatrix(Num, 1) = "Not Detected"
60            ElseIf Index = 1 Then
70                grdBatch(Index - 1).TextMatrix(Num, 2) = "Negative"
80            Else
90                grdBatch(Index).TextMatrix(Num, 1) = "Negative"
100           End If
110       Next

120       cmdSave.Enabled = True

130       Exit Sub

cmdSet_Click_Error:

          Dim strES As String
          Dim intEL As Integer

140       intEL = Erl
150       strES = Err.Description
160       LogError "frmBatchERA", "cmdSet_Click", intEL, strES


End Sub



Private Sub cmdXL_Click()
10        If grdBatch(0).Rows + grdBatch(2).Rows + grdBatch(3).Rows = 3 Then
20            iMsg "Nothing to export", vbInformation
30            Exit Sub
40        End If

          Dim strHeading As String
          Dim i As Integer
50        For i = 0 To 3
60            If i <> 1 Then
70                strHeading = "Batch Entry" & vbCr
80                strHeading = strHeading & Trim(lblHeading(i).Caption) & IIf(optUnResulted.Value = True, " Unresulted Sample Requests", " Unvalidated Sample Results")
90                strHeading = strHeading & vbCr & vbCr
100               ExportFlexGrid grdBatch(i), Me, strHeading
110           End If
120       Next i
End Sub

Private Sub Form_Load()

10        dtStart.Value = Date - 3
20        dtEnd.Value = Date

30        MarkBatchEntryOpen4Use ("OPTBATCHENTRYERA")

40        LoadListGenericColour ListCDiffToxinAB(), "CDiffToxinAB"
50        LoadListGenericColour ListAdeno(), "Adeno"
60        LoadListGenericColour ListRota(), "Rota"
70        LoadListGenericColour ListHPylori(), "HPylori"

80        FillG "A"
90        FillG "G"
100       FillG "Y"

End Sub

Private Sub Form_Unload(Cancel As Integer)

10        delBatchEntryOpenStatus ("OPTBATCHENTRYERA")

End Sub


Private Sub grdBatch_Click(Index As Integer)

10        If grdBatch(Index).Rows = 1 Then Exit Sub


          'To check right mouse button pass 2
20        If GetAsyncKeyState(2) <> 0 Then
30            Exit Sub
40        End If

50        With grdBatch(Index)

60            If .MouseRow = 0 Then Exit Sub
70            If .MouseCol = 0 Then Exit Sub

80            If Index = 0 Then
90                If .MouseCol = 1 Then
100                   CycleGridCell ListRota, .Row, .Col, grdBatch(0)
110               ElseIf .MouseCol = 2 Then
120                   CycleGridCell ListAdeno, .Row, .Col, grdBatch(0)
130               ElseIf .MouseCol = 3 Then
140                   .Col = 3
150                   If .CellPicture = imgSquareTick.Picture Then
160                       Set .CellPicture = imgSquareCross.Picture
170                   Else
180                       Set .CellPicture = imgSquareTick.Picture
190                   End If

200               End If

210           ElseIf Index = 2 Then    'C Diff
220               If .MouseCol = 1 Then
230                   CycleGridCell ListCDiffToxinAB, .Row, .Col, grdBatch(2)
240               ElseIf .MouseCol = 2 Then
250                   .Col = 2
260                   If .CellPicture = imgSquareTick.Picture Then
270                       Set .CellPicture = imgSquareCross.Picture
280                   Else
290                       Set .CellPicture = imgSquareTick.Picture
300                   End If
310               End If
320           ElseIf Index = 3 Then    'H.Pylori
330               If .MouseCol = 1 Then
340                   CycleGridCell ListHPylori, .Row, .Col, grdBatch(3)
350               ElseIf .MouseCol = 2 Then
360                   .Col = 2
370                   If .CellPicture = imgSquareTick.Picture Then
380                       Set .CellPicture = imgSquareCross.Picture
390                   Else
400                       Set .CellPicture = imgSquareTick.Picture
410                   End If
420               End If
430           End If


440       End With

450       cmdSave.Enabled = True

End Sub

Private Sub FillG(Department As String)

          Dim tb As Recordset
          Dim sql As String
          Dim Num As Integer
          Dim Field As String
          Dim Field2 As String
          Dim s As String
          Dim T() As String
          Dim ForeColour As Long
          Dim BackColour As Long
          Dim ForeColour2 As Long
          Dim BackColour2 As Long
          Dim ValidCol As Byte

10        On Error GoTo FillG_Error

20        Select Case Department
          Case "A":
30            Num = 0
40            Field = "Rota"
50            Field2 = "Adeno"
60            ValidCol = 3
70        Case "G":
80            Num = 2
90            Field = "ToxinAB"
100           ValidCol = 2
110       Case "Y":
120           Num = 3
130           Field = "HPylori"
140           ValidCol = 2
150       End Select

160       With grdBatch(Num)
170           .Rows = 2
180           .AddItem ""
190           .RemoveItem 1
200           .Rows = 1

210       End With

220       If Num = 0 Then

230           If optUnResulted Then
240               sql = "Select R.SampleID, F.%field1, F.%field2, COALESCE(L.Valid, 0) As Valid From FaecalRequests R " & _
                        "Left Join (Select * From PrintValidLog Where Department = '%department') L On R.SampleID = L.SampleID " & _
                        "Left Join Faeces F On R.SampleID = F.SampleID " & _
                        "Where (R.%field1 = 1 Or R.%field2 = 1) And COALESCE(L.Valid, 0) = 0 " & _
                        "And (COALESCE(F.%field1, '') = '' And COALESCE(F.%field2, '') = '') " & _
                        "And R.DateTimeOfRecord Between '%date1' And '%date2' " & _
                        "Order By R.SampleID"
250           ElseIf optNotValid Then
260               sql = "Select F.SampleID, F.%field1, F.%field2, COALESCE(L.Valid, 0) As Valid From Faeces F " & _
                        "Left Join (Select * From PrintValidLog Where Department = '%department') L On F.SampleID = L.SampleID " & _
                        "Inner Join FaecalRequests R On R.SampleID = F.SampleID " & _
                        "Where (COALESCE(F.%field1, '') <> '' Or COALESCE(F.%field2, '') <> '') " & _
                        "And COALESCE(L.Valid, 0) = 0 " & _
                        "And R.DateTimeOfRecord Between '%date1' And '%date2' " & _
                        "Order By F.SampleID"
270           End If
280           sql = Replace(sql, "%field1", Field)
290           sql = Replace(sql, "%field2", Field2)
300           sql = Replace(sql, "%department", Department)
310           sql = Replace(sql, "%date1", Format(dtStart.Value, "dd/MMM/yyyy hh:mm:ss"))
320           sql = Replace(sql, "%date2", Format(dtEnd.Value + 1, "dd/MMM/yyyy hh:mm:ss"))
330       Else
340           If optUnResulted Then
350               sql = "Select R.SampleID, F.%field1, COALESCE(L.Valid, 0) As Valid From FaecalRequests R " & _
                        "Left Join (Select * From PrintValidLog Where Department = '%department') L On R.SampleID = L.SampleID " & _
                        "Left Join Faeces F On R.SampleID = F.SampleID " & _
                        "Where R.%field1 = 1 And COALESCE(L.Valid, 0) = 0 " & _
                        "And COALESCE(F.%field1, '') = '' " & _
                        "And R.DateTimeOfRecord Between '%date1' And '%date2'"
360           ElseIf optNotValid Then
370               sql = "Select F.SampleID, F.%field1, COALESCE(L.Valid, 0) As Valid From Faeces F " & _
                        "Left Join (Select * From PrintValidLog Where Department = '%department') L On F.SampleID = L.SampleID " & _
                        "Inner Join FaecalRequests R On R.SampleID = F.SampleID " & _
                        "Where COALESCE(F.%field1, '') <> '' " & _
                        "And COALESCE(L.Valid, 0) = 0 " & _
                        "And R.DateTimeOfRecord Between '%date1' And '%date2'"
380           End If
390           sql = Replace(sql, "%field1", Field)
400           sql = Replace(sql, "%department", Department)
410           sql = Replace(sql, "%date1", Format(dtStart.Value, "dd/MMM/yyyy hh:mm:ss"))
420           sql = Replace(sql, "%date2", Format(dtEnd.Value + 1, "dd/MMM/yyyy hh:mm:ss"))
430       End If


440       Set tb = New Recordset
450       RecOpenClient 0, tb, sql
460       Do While Not tb.EOF
470           s = tb!SampleID - SysOptMicroOffset(0) & vbTab
480           T = Split(tb(Field) & "", "|")
490           If UBound(T) = -1 Then
500               s = s & "" & vbTab
510           ElseIf UBound(T) > 1 Then
520               s = s & T(0) & vbTab
530               ForeColour = T(1)
540               BackColour = T(2)
550           Else
560               s = s & T(0) & vbTab
570           End If

580           If Num = 0 Then
590               T = Split(tb(Field2) & "", "|")
600               If UBound(T) = -1 Then
610                   s = s & "" & vbTab
620               ElseIf UBound(T) > 1 Then
630                   s = s & T(0) & vbTab
640                   ForeColour2 = T(1)
650                   BackColour2 = T(2)
660               Else
670                   s = s & vbTab
680               End If
690           End If

700           grdBatch(Num).AddItem s
710           grdBatch(Num).Row = grdBatch(Num).Rows - 1
720           grdBatch(Num).Col = 1
730           grdBatch(Num).CellBackColor = BackColour
740           grdBatch(Num).CellForeColor = ForeColour
750           If Num = 0 Then
760               grdBatch(Num).Col = 2
770               grdBatch(Num).CellBackColor = BackColour2
780               grdBatch(Num).CellForeColor = ForeColour2
790           End If
800           grdBatch(Num).Col = ValidCol
810           grdBatch(Num).CellPictureAlignment = flexAlignCenterCenter
820           Set grdBatch(Num).CellPicture = IIf(tb!Valid, imgSquareTick, imgSquareCross)

830           tb.MoveNext
840       Loop

850       Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

860       intEL = Erl
870       strES = Err.Description
880       LogError "frmBatchERA", "FillG", intEL, strES, sql

End Sub

Private Sub optNotValid_Click()
10        FillG "A"
20        FillG "G"
30        FillG "Y"
End Sub

Private Sub optUnResulted_Click()
10        FillG "A"
20        FillG "G"
30        FillG "Y"
End Sub
