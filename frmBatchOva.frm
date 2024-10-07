VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmBatchOva 
   Appearance      =   0  'Flat
   Caption         =   "NetAcquire - Batch Ova/Parasites"
   ClientHeight    =   8115
   ClientLeft      =   285
   ClientTop       =   870
   ClientWidth     =   5820
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
   Icon            =   "frmBatchOva.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8115
   ScaleWidth      =   5820
   Begin VB.CommandButton bsearch 
      Caption         =   "&Search"
      Default         =   -1  'True
      Height          =   900
      Left            =   3480
      Picture         =   "frmBatchOva.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   240
      Width           =   1100
   End
   Begin VB.OptionButton optNotValid 
      Caption         =   "Show Unvalidated Results"
      ForeColor       =   &H00008000&
      Height          =   195
      Left            =   180
      TabIndex        =   7
      Top             =   1350
      Width           =   2985
   End
   Begin VB.OptionButton optUnResulted 
      Caption         =   "Show Unresulted Requests"
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   180
      TabIndex        =   6
      Top             =   1110
      Value           =   -1  'True
      Width           =   2985
   End
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   900
      Left            =   4500
      Picture         =   "frmBatchOva.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4005
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
      Left            =   4500
      Picture         =   "frmBatchOva.frx":0A56
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6000
      Width           =   1100
   End
   Begin VB.CommandButton cmdSetAll 
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
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2160
      Width           =   1845
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
      Left            =   4500
      Picture         =   "frmBatchOva.frx":0D60
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7035
      Width           =   1100
   End
   Begin MSFlexGridLib.MSFlexGrid grdOva 
      Height          =   5355
      Left            =   210
      TabIndex        =   1
      Top             =   2580
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   9446
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
      GridLinesFixed  =   3
      ScrollBars      =   2
      FormatString    =   "<Sample ID       |<Result                              |^Valid   "
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
   Begin MSComCtl2.DTPicker dtStart 
      Height          =   315
      Left            =   1650
      TabIndex        =   9
      Top             =   240
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
      Format          =   59637761
      CurrentDate     =   40267
   End
   Begin MSComCtl2.DTPicker dtEnd 
      Height          =   315
      Left            =   1650
      TabIndex        =   10
      Top             =   660
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
      Format          =   59637761
      CurrentDate     =   40267
   End
   Begin VB.Image imgSquareTick 
      Height          =   225
      Left            =   4920
      Picture         =   "frmBatchOva.frx":106A
      Top             =   600
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgSquareCross 
      Height          =   225
      Left            =   5100
      Picture         =   "frmBatchOva.frx":1340
      Top             =   600
      Visible         =   0   'False
      Width           =   210
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
      Top             =   690
      Width           =   255
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
      TabIndex        =   12
      Top             =   270
      Width           =   1350
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      Height          =   345
      Left            =   4500
      TabIndex        =   11
      Top             =   3600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cryptosporidium"
      Height          =   255
      Left            =   210
      TabIndex        =   3
      Top             =   1830
      Width           =   4065
   End
End
Attribute VB_Name = "frmBatchOva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim ListCrypto() As ListColour

Private Sub bsearch_Click()
10        FillGrid
End Sub

Private Sub cmdCancel_Click()

10        delBatchEntryOpenStatus ("OPTBATCHENTRYOVA")

20        Unload Me

End Sub

Private Sub cmdSave_Click()

10        SaveOP

20        cmdSave.Enabled = False

End Sub

Private Sub cmdSetAll_Click()

          Dim Num As Long

10        On Error GoTo cmdSetAll_Click_Error

20        If grdOva.Rows = 2 And grdOva.TextMatrix(1, 0) = "" Then Exit Sub

30        For Num = 1 To grdOva.Rows - 1
40            grdOva.TextMatrix(Num, 1) = "Not Detected"
50        Next

60        cmdSave.Enabled = True

70        Exit Sub

cmdSetAll_Click_Error:

          Dim strES As String
          Dim intEL As Integer

80        intEL = Erl
90        strES = Err.Description
100       LogError "frmBatchOva", "cmdSetAll_Click", intEL, strES

End Sub



Private Sub FillGrid()

          Dim tb As Recordset
          Dim sql As String
          Dim s As String
          Dim T() As String
          Dim ForeColour As Long
          Dim BackColour As Long

10        On Error GoTo FillGrid_Error

20        With grdOva
30            .Rows = 2
40            .AddItem ""
50            .RemoveItem 1
60            .Rows = 1

70            If optUnResulted Then
80                sql = "Select R.SampleID, F.Cryptosporidium, COALESCE(L.Valid, 0) As Valid From FaecalRequests R " & _
                        "Left Join (Select * From PrintValidLog Where Department = 'O') L On R.SampleID = L.SampleID " & _
                        "Left Join Faeces F On R.SampleID = F.SampleID " & _
                        "Where R.Cryptosporidium = 1 And COALESCE(l.Valid, 0) = 0 " & _
                        "And COALESCE(F.Cryptosporidium, '') = '' " & _
                        "And R.DateTimeOfRecord Between '%date1' And '%date2' " & _
                        "Order By R.SampleID"
90            Else
100               sql = "Select F.SampleID, F.Cryptosporidium, COALESCE(L.Valid, 0) As Valid From Faeces F " & _
                        "Left Join (Select * From PrintValidLog Where Department = 'O') L On F.SampleID = L.SampleID " & _
                        "Inner Join FaecalRequests R On R.SampleID = F.SampleID " & _
                        "Where COALESCE(F.Cryptosporidium, '') <> '' " & _
                        "And COALESCE(L.Valid, 0) = 0 " & _
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
200               .Col = 1
210               T = Split(tb!Cryptosporidium & "", "|")
220               If UBound(T) = -1 Then
230                   s = ""
240               ElseIf UBound(T) > 1 Then
250                   s = T(0)
260                   ForeColour = T(1)
270                   BackColour = T(2)
280               Else
290                   s = T(0)
300               End If
310               .TextMatrix(.Row, 1) = s
320               .CellBackColor = BackColour
330               .CellForeColor = ForeColour
340               .Col = 2
350               .CellPictureAlignment = flexAlignCenterCenter
360               Set .CellPicture = IIf(tb!Valid, imgSquareTick.Picture, imgSquareCross.Picture)
370               tb.MoveNext
380           Loop


390       End With

400       Exit Sub

FillGrid_Error:

          Dim strES As String
          Dim intEL As Integer

410       intEL = Erl
420       strES = Err.Description
430       LogError "frmBatchOccult", "FillGrid", intEL, strES, sql

End Sub

Private Sub cmdXL_Click()
10        If grdOva.Rows = 1 Then
20            iMsg "Nothing to export", vbInformation
30            Exit Sub
40        End If

          Dim strHeading As String
50        strHeading = "Batch Entry" & vbCr
60        strHeading = strHeading & "Ova/Parasites " & IIf(optUnResulted, "Unresulted Sample Requests", "Unvalidated Sample Results")
70        strHeading = strHeading & vbCr & vbCr
80        ExportFlexGrid grdOva, Me, strHeading

End Sub

Private Sub Form_Load()

10        dtStart.Value = Date - 3
20        dtEnd.Value = Date

30        MarkBatchEntryOpen4Use ("OPTBATCHENTRYOVA")

40        LoadListGenericColour ListCrypto(), "Crypto"

50        FillGrid

End Sub

Private Sub grdOva_Click()

10        On Error GoTo grdOva_Click_Error

20        If grdOva.Rows = 1 Then Exit Sub

          'To check right mouse button pass 2
30        If GetAsyncKeyState(2) <> 0 Then
40            Exit Sub
50        End If

60        With grdOva
70            If .MouseRow = 0 Then Exit Sub
80            If .MouseCol = 0 Then Exit Sub
90            If .TextMatrix(1, 0) = "" Then Exit Sub

100           Select Case .Col
              Case 1:
110               CycleGridCell ListCrypto(), .Row, .Col, grdOva
120           Case 2:
130               .Col = 2
140               If .CellPicture = imgSquareTick.Picture Then
150                   Set .CellPicture = imgSquareCross.Picture
160               Else
170                   Set .CellPicture = imgSquareTick.Picture
180               End If
190           End Select

200       End With

210       cmdSave.Enabled = True

220       Exit Sub

grdOva_Click_Error:

          Dim strES As String
          Dim intEL As Integer

230       intEL = Erl
240       strES = Err.Description
250       LogError "frmBatchOva", "grdOva_Click", intEL, strES

End Sub

Private Sub SaveOP()

          Dim tb As Recordset
          Dim sql As String
          Dim Y As Integer

10        On Error GoTo SaveOp_Error

20        With grdOva
30            For Y = 1 To .Rows - 1
40                If .TextMatrix(Y, 0) <> "" And Left$(grdOva.TextMatrix(Y, 1), 1) <> "?" Then
50                    If ValidStatus4MicroDept(Val(.TextMatrix(Y, 0)) + SysOptMicroOffset(0), "O") = False Then

60                        sql = "SELECT * FROM Faeces WHERE " & _
                                "SampleID = '" & Val(.TextMatrix(Y, 0)) + SysOptMicroOffset(0) & "'"
70                        Set tb = New Recordset
80                        RecOpenServer 0, tb, sql
90                        If tb.EOF Then
100                           tb.AddNew
110                           tb!SampleID = Val(.TextMatrix(Y, 0)) + SysOptMicroOffset(0)
120                       End If
130                       .Col = 1
140                       .Row = Y
150                       tb!Cryptosporidium = .TextMatrix(Y, 1) & "|" & .CellForeColor & "|" & .CellBackColor
160                       tb!Username = Username

170                       tb.Update

180                   End If
190                   .Col = 2
200                   .Row = Y
210                   If .CellPicture = imgSquareTick Then
220                       UpdatePrintValidLog Val(.TextMatrix(Y, 0)) + SysOptMicroOffset(0), "OP", 1, 0
230                   ElseIf .CellPicture = imgSquareCross Then
240                       UpdatePrintValidLog Val(.TextMatrix(Y, 0)) + SysOptMicroOffset(0), "OP", 0, 0
250                   End If
260               End If
270           Next
280       End With
290       FillGrid

300       Exit Sub

SaveOp_Error:

          Dim strES As String
          Dim intEL As Integer

310       intEL = Erl
320       strES = Err.Description
330       LogError "frmBatchOva", "SaveOp", intEL, strES

End Sub


Private Sub optNotValid_Click()
10        FillGrid
End Sub

Private Sub optUnResulted_Click()
10        FillGrid
End Sub
