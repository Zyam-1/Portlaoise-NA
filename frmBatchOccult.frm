VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmBatchOccult 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire Batch Entry - Occult Blood"
   ClientHeight    =   9045
   ClientLeft      =   330
   ClientTop       =   480
   ClientWidth     =   10230
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
   Icon            =   "frmBatchOccult.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9045
   ScaleWidth      =   10230
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   900
      Left            =   8940
      Picture         =   "frmBatchOccult.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4620
      Width           =   1100
   End
   Begin VB.CommandButton bsearch 
      Caption         =   "&Search"
      Default         =   -1  'True
      Height          =   735
      Left            =   3480
      Picture         =   "frmBatchOccult.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   180
      Width           =   1095
   End
   Begin VB.OptionButton optNotValid 
      Caption         =   "Show Unvalidated Results"
      ForeColor       =   &H00008000&
      Height          =   195
      Left            =   180
      TabIndex        =   8
      Top             =   1290
      Width           =   2985
   End
   Begin VB.OptionButton optUnResulted 
      Caption         =   "Show Unresulted Requests"
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   180
      TabIndex        =   7
      Top             =   1050
      Value           =   -1  'True
      Width           =   2985
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
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
      Left            =   8940
      Picture         =   "frmBatchOccult.frx":0A56
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7020
      Width           =   1100
   End
   Begin MSFlexGridLib.MSFlexGrid grdOccult 
      Height          =   6135
      Left            =   180
      TabIndex        =   5
      Top             =   2760
      Width           =   8580
      _ExtentX        =   15134
      _ExtentY        =   10821
      _Version        =   393216
      Cols            =   5
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
      FormatString    =   "<Sample ID           |<Result  1                  |<Result  2                  |<Result  3                  |^Valid    "
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
   Begin VB.CommandButton cmdSetAll 
      Caption         =   "S e t    1,  2,  a n d  3    A l l    N e g a t i v e"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   425
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1710
      Width           =   5745
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "Set All Negative (3)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   425
      Index           =   2
      Left            =   5820
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2220
      Width           =   1815
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "Set All Negative (2)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   425
      Index           =   1
      Left            =   3870
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2220
      Width           =   1815
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "Set All Negative (1)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   425
      Index           =   0
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2220
      Width           =   1815
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
      Height          =   900
      Left            =   8940
      Picture         =   "frmBatchOccult.frx":0D60
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7995
      Width           =   1100
   End
   Begin MSComCtl2.DTPicker dtStart 
      Height          =   315
      Left            =   1650
      TabIndex        =   10
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
      Format          =   16515073
      CurrentDate     =   40267
   End
   Begin MSComCtl2.DTPicker dtEnd 
      Height          =   315
      Left            =   1650
      TabIndex        =   11
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
      Format          =   16515073
      CurrentDate     =   40267
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      Height          =   345
      Left            =   8940
      TabIndex        =   15
      Top             =   4200
      Visible         =   0   'False
      Width           =   1100
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
      TabIndex        =   13
      Top             =   630
      Width           =   255
   End
   Begin VB.Label Label1 
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
      TabIndex        =   12
      Top             =   210
      Width           =   1350
   End
   Begin VB.Image imgSquareTick 
      Height          =   225
      Left            =   6840
      Picture         =   "frmBatchOccult.frx":106A
      Top             =   660
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgSquareCross 
      Height          =   225
      Left            =   7215
      Picture         =   "frmBatchOccult.frx":1340
      Top             =   660
      Visible         =   0   'False
      Width           =   210
   End
End
Attribute VB_Name = "frmBatchOccult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Dim ListFOB() As ListColour

Private Sub bsearch_Click()
10        FillGrid

End Sub

Private Sub cmdCancel_Click()

10        On Error GoTo cmdCancel_Click_Error

20        If cmdSave.Enabled Then
30            If iMsg("Cancel without Saving?", vbQuestion + vbYesNo) = vbNo Then
40                Exit Sub
50            End If
60        End If

70        Unload Me

80        Exit Sub

cmdCancel_Click_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmBatchOccult", "cmdCancel_Click", intEL, strES

End Sub

Private Sub cmdSave_Click()

          Dim tb As New Recordset
          Dim sql As String
          Dim Y As Integer
          Dim X As Integer
          Dim Result As String
          Dim FieldName As String

10        On Error GoTo cmdSave_Click_Error

20        With grdOccult
30            For Y = 1 To .Rows - 1
40                If LTrim(RTrim(.TextMatrix(Y, 1) & .TextMatrix(Y, 2) & .TextMatrix(Y, 3))) <> "" Then
50                    For X = 1 To 3
60                        FieldName = Choose(X, "OB0", "OB1", "OB2")

                          '                Select Case UCase(Trim$(.TextMatrix(Y, X)))
                          '                    Case "NEGATIVE": Result = "N"
                          '                    Case "POSITIVE": Result = "P"
                          '                    Case "INSUFFICIENT SAMPLE": Result = "U"
                          '                    Case "?": Result = "?"
                          '                    Case Else: Result = ""
                          '                End Select

70                        If .TextMatrix(Y, X) <> "?" Then
80                            If ValidStatus4MicroDept(Val(.TextMatrix(Y, 0)) + SysOptMicroOffset(0), "F") = False Then
90                                sql = "SELECT * FROM Faeces WHERE " & _
                                        "SampleID = '" & Val(.TextMatrix(Y, 0)) + SysOptMicroOffset(0) & "'"
100                               Set tb = New Recordset
110                               RecOpenServer 0, tb, sql
120                               If tb.EOF Then
130                                   tb.AddNew
140                                   tb!SampleID = Val(.TextMatrix(Y, 0)) + SysOptMicroOffset(0)
150                               End If
160                               .Col = X
170                               .Row = Y
180                               tb(FieldName) = .TextMatrix(Y, X) & "|" & .CellForeColor & "|" & .CellBackColor
190                               tb!Username = Username
200                               tb.Update

210                           End If
220                       End If
230                       Result = ""
240                   Next
                      'update printvalid log here
250                   .Col = 4
260                   .Row = Y
270                   If .CellPicture = imgSquareTick.Picture Then
280                       UpdatePrintValidLog .TextMatrix(Y, 0) + SysOptMicroOffset(0), "FOB", 1, 0
290                   ElseIf .CellPicture = imgSquareCross.Picture Then
300                       UpdatePrintValidLog .TextMatrix(Y, 0) + SysOptMicroOffset(0), "FOB", 0, 0
310                   End If
320               End If
330           Next
340       End With

350       FillGrid
360       cmdSave.Enabled = False

370       Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

380       intEL = Erl
390       strES = Err.Description
400       LogError "frmBatchOccult", "cmdsave_Click", intEL, strES, sql


End Sub

Private Sub cmdSet_Click(Index As Integer)

          Dim n As Long

10        On Error GoTo cmdSet_Click_Error

20        If grdOccult.Rows = 2 And grdOccult.TextMatrix(1, 0) = "" Then Exit Sub

30        With grdOccult()
40            .Col = Index + 1
50            For n = 1 To .Rows - 1
60                .Row = n
70                If .TextMatrix(n, 0) <> "" Then
80                    .Text = "Negative"
90                    .CellForeColor = vbBlue
100                   .CellBackColor = vbGreen
110               End If
120           Next
130       End With

140       cmdSave.Enabled = True

150       Exit Sub

cmdSet_Click_Error:

          Dim strES As String
          Dim intEL As Integer

160       intEL = Erl
170       strES = Err.Description
180       LogError "frmBatchOccult", "cmdSet_Click", intEL, strES


End Sub

Private Sub cmdSetAll_Click()

          Dim n As Long

10        On Error GoTo cmdSetAll_Click_Error

20        For n = 0 To 2
30            cmdSet_Click (n)
40        Next

50        Exit Sub

cmdSetAll_Click_Error:

          Dim strES As String
          Dim intEL As Integer

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmBatchOccult", "cmdSetAll_Click", intEL, strES


End Sub

Private Sub FillGrid()

          Dim tb As Recordset
          Dim sql As String
          Dim n As Integer
          Dim s As String
          Dim T() As String
          Dim ForeColour As Long
          Dim BackColour As Long

10        On Error GoTo FillGrid_Error

20        With grdOccult
30            .Rows = 2
40            .AddItem ""
50            .RemoveItem 1
60            .Rows = 1

70            If optUnResulted Then
80                sql = "Select R.SampleID, F.OB0, F.OB1, F.OB2, COALESCE(L.Valid, 0) As Valid From FaecalRequests R " & _
                        "Left Join (Select * From PrintValidLog Where Department = 'F') L On R.SampleID = L.SampleID " & _
                        "Left Join Faeces F On R.SampleID = F.SampleID " & _
                        "Where (R.OB0 = 1 Or R.OB1 = 1 Or R.OB2 = 1) And COALESCE(L.Valid, 0) = 0 " & _
                        "And COALESCE(F.OB0, '') = '' And COALESCE(F.OB1, '') = '' And COALESCE(F.OB2, '') = '' " & _
                        "And R.DateTimeOfRecord Between '%date1' And '%date2' " & _
                        "Order By R.SampleID"
90            Else
100               sql = "Select F.SampleID, F.OB0, F.OB1, F.OB2, COALESCE(L.Valid, 0) As Valid From Faeces F " & _
                        "Left Join (Select * From PrintValidLog Where Department = 'F') L On F.SampleID = L.SampleID " & _
                        "Inner Join FaecalRequests R On R.SampleID = F.SampleID " & _
                        "Where (COALESCE(F.OB0, '') <> '' Or COALESCE(F.OB1, '') <> '' Or COALESCE(F.OB2, '') <> '') " & _
                        "And COALESCE(L.Valid, 0) = 0 " & _
                        "And (COALESCE(F.OB0, '') <> '' OR COALESCE(F.OB1, '') = '' Or COALESCE(F.OB2, '') = '') " & _
                        "And R.DateTimeOfRecord Between '%date1' And '%date2' " & _
                        "Order By F.SampleID"
110           End If

120           sql = Replace(sql, "%date1", Format(dtStart.Value, "dd/MMM/yyyy hh:mm:ss"))
130           sql = Replace(sql, "%date2", Format(dtEnd.Value + 1, "dd/MMM/yyyy hh:mm:ss"))

140           Set tb = New Recordset
150           RecOpenClient 0, tb, sql
160           Do While Not tb.EOF
170               s = tb!SampleID - SysOptMicroOffset(0)
180               .AddItem s
190               .Row = .Rows - 1
200               For n = 0 To 2
210                   .Col = n + 1
220                   T = Split(tb("OB" & Format$(n)) & "", "|")
230                   If UBound(T) = -1 Then
240                       s = ""
250                   ElseIf UBound(T) > 1 Then
260                       s = T(0)
270                       ForeColour = T(1)
280                       BackColour = T(2)
290                   Else
300                       s = T(0)
310                   End If
320                   .TextMatrix(.Row, n + 1) = s
330                   .CellBackColor = BackColour
340                   .CellForeColor = ForeColour
350               Next
360               .Col = 4
370               .CellPictureAlignment = flexAlignCenterCenter
380               Set .CellPicture = IIf(tb!Valid, imgSquareTick.Picture, imgSquareCross.Picture)
390               tb.MoveNext
400           Loop


410       End With

420       Exit Sub

FillGrid_Error:

          Dim strES As String
          Dim intEL As Integer

430       intEL = Erl
440       strES = Err.Description
450       LogError "frmBatchOccult", "FillGrid", intEL, strES, sql

End Sub

Private Sub cmdXL_Click()
10        If grdOccult.Rows = 1 Then
20            iMsg "Nothing to export", vbInformation
30            Exit Sub
40        End If

          Dim strHeading As String
50        strHeading = "Batch Entry" & vbCr
60        strHeading = strHeading & "Urine " & IIf(optUnResulted, "Unresulted Sample Requests", "Unvalidated Sample Results")
70        strHeading = strHeading & vbCr & vbCr
80        ExportFlexGrid grdOccult, Me, strHeading
End Sub

Private Sub Form_Load()
10        dtStart.Value = Date - 3
20        dtEnd.Value = Date

30        MarkBatchEntryOpen4Use ("OPTBATCHENTRYOCCULTBLOOD")
40        LoadListGenericColour ListFOB(), "OccultBlood"

50        FillGrid

End Sub



Private Sub Form_Unload(Cancel As Integer)

10        delBatchEntryOpenStatus ("OPTBATCHENTRYOCCULTBLOOD")

End Sub


Private Sub grdOccult_Click()

10        On Error GoTo grdOccult_Click_Error

20        If grdOccult.Rows = 1 Then Exit Sub

          'To check right mouse button pass 2
30        If GetAsyncKeyState(2) <> 0 Then
40            Exit Sub
50        End If


60        With grdOccult()
70            If .MouseRow = 0 Then Exit Sub
80            If .MouseCol = 0 Then Exit Sub
90            Select Case .Col
              Case 1, 2, 3:
100               CycleGridCell ListFOB, .Row, .Col, grdOccult
110           Case 4:
120               If .CellPicture = imgSquareTick.Picture Then
130                   Set .CellPicture = imgSquareCross.Picture
140               Else
150                   Set .CellPicture = imgSquareTick.Picture
160               End If
170           End Select
180       End With

190       cmdSave.Enabled = True

200       Exit Sub

grdOccult_Click_Error:

          Dim strES As String
          Dim intEL As Integer

210       intEL = Erl
220       strES = Err.Description
230       LogError "frmBatchOccult", "grdOccult_Click", intEL, strES


End Sub


Private Sub optNotValid_Click()
10        FillGrid
End Sub

Private Sub optUnResulted_Click()
10        FillGrid
End Sub
