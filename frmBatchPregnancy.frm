VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmBatchPregnancy 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire -Pregnancy Batch Entry "
   ClientHeight    =   6885
   ClientLeft      =   690
   ClientTop       =   1530
   ClientWidth     =   6075
   Icon            =   "frmBatchPregnancy.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   885
      Left            =   4740
      Picture         =   "frmBatchPregnancy.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2760
      Width           =   1125
   End
   Begin VB.OptionButton optUnResulted 
      Caption         =   "Show Unresulted Requests"
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   180
      TabIndex        =   6
      Top             =   1050
      Value           =   -1  'True
      Width           =   2265
   End
   Begin VB.OptionButton optNotValid 
      Caption         =   "Show Unvalidated Results"
      ForeColor       =   &H00008000&
      Height          =   195
      Left            =   180
      TabIndex        =   5
      Top             =   1290
      Width           =   2265
   End
   Begin VB.CommandButton bsearch 
      Caption         =   "&Search"
      Default         =   -1  'True
      Height          =   735
      Left            =   3480
      Picture         =   "frmBatchPregnancy.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   180
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   885
      Left            =   4740
      Picture         =   "frmBatchPregnancy.frx":0A56
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4680
      Width           =   1125
   End
   Begin MSFlexGridLib.MSFlexGrid grdBatch 
      Height          =   5055
      Left            =   180
      TabIndex        =   1
      Top             =   1620
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   8916
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
      GridLines       =   2
      ScrollBars      =   2
      FormatString    =   "<Sample ID     |<Pregnancy                 |^Valid    "
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
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   885
      Left            =   4740
      Picture         =   "frmBatchPregnancy.frx":0D60
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5745
      Width           =   1125
   End
   Begin MSComCtl2.DTPicker dtStart 
      Height          =   315
      Left            =   1650
      TabIndex        =   4
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
      Format          =   59965441
      CurrentDate     =   40267
   End
   Begin MSComCtl2.DTPicker dtEnd 
      Height          =   315
      Left            =   1650
      TabIndex        =   7
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
      Format          =   59965441
      CurrentDate     =   40267
   End
   Begin VB.Image imgSquareCross 
      Height          =   225
      Left            =   5535
      Picture         =   "frmBatchPregnancy.frx":106A
      Top             =   1620
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgSquareTick 
      Height          =   225
      Left            =   5160
      Picture         =   "frmBatchPregnancy.frx":1340
      Top             =   1620
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4740
      TabIndex        =   11
      Top             =   3660
      Visible         =   0   'False
      Width           =   1125
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
      TabIndex        =   9
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
      TabIndex        =   8
      Top             =   630
      Width           =   255
   End
End
Attribute VB_Name = "frmBatchPregnancy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ListPregnancy() As String

Private Sub bsearch_Click()
10        FillG
End Sub

Private Sub cmdCancel_Click()

10        Unload Me

End Sub

Private Sub cmdSave_Click()

          Dim sql As String
          Dim Y As Integer
          Dim Pregnancy As String
          Dim SampleID As Double
          Dim ValidFlag As Integer

10        On Error GoTo cmdSave_Click_Error

20        If grdBatch.Rows = 1 Then Exit Sub

30        For Y = 1 To grdBatch.Rows - 1
40            ValidFlag = 0
50            grdBatch.Col = 2
60            If grdBatch.CellPicture = imgSquareTick.Picture Then
70                ValidFlag = 1
80            End If
90            SampleID = Val(grdBatch.TextMatrix(Y, 0)) + SysOptMicroOffset(0)
100           Pregnancy = IIf(Trim(grdBatch.TextMatrix(Y, 1)) <> "", grdBatch.TextMatrix(Y, 1), "")
110           sql = "IF EXISTS (SELECT * FROM Urine WHERE SampleID = '" & SampleID & "') " & _
                    "  UPDATE Urine " & _
                    "  SET Pregnancy = '" & Pregnancy & "', " & _
                    "  UserName = '" & Username & "' " & _
                    "  WHERE SampleID = " & SampleID & " " & _
                    "ELSE " & _
                    " INSERT INTO Urine " & _
                    " ( SampleID, Pregnancy, Valid, Printed, UserName) VALUES " & _
                    " (" & SampleID & ", " & _
                    "  '" & Pregnancy & "', " & _
                    "  '" & ValidFlag & "', " & _
                    "  '0', " & _
                    "  '" & Username & "' )"
120           Cnxn(0).Execute sql
130           UpdatePrintValidLog SampleID, "URINE", ValidFlag, 0
140       Next

150       FillG

160       Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

170       intEL = Erl
180       strES = Err.Description
190       LogError "frmBatchPregnancy", "cmdSave_Click", intEL, strES, sql

End Sub

Sub FillG()

          Dim tb As Recordset
          Dim sql As String
          Dim s As String

10        On Error GoTo FillG_Error

20        If optUnResulted Then
30            sql = "SELECT R.SampleID, U.Pregnancy, COALESCE(L.Valid, 0) Valid FROM UrineRequests AS R " & _
                    "Left Join (Select SampleID, COALESCE(Valid, 0) As Valid From PrintValidLog Where Department = 'U') L On R.SampleID = L.SampleID " & _
                    "Left Join Urine AS U On R.SampleID = U.SampleID " & _
                    "Where R.Pregnancy = 1 And COALESCE(U.Pregnancy, '') = '' " & _
                    "AND R.DateTimeOfRecord Between '%date1' And '%date2' " & _
                    "ORDER By R.SampleID"


40        Else
50            sql = "Select U.SampleID, U.Pregnancy, COALESCE(L.Valid, 0) Valid FROM Urine U " & _
                    "Left Join (Select SampleID, COALESCE(Valid, 0) As Valid From PrintValidLog Where Department = 'U') L On U.SampleID = L.SampleID " & _
                    "Inner Join UrineRequests R On R.SampleID = U.SampleID " & _
                    "Where COALESCE(U.Pregnancy, '') <> '' And COALESCE(L.Valid, 0) = 0 " & _
                    "AND R.DateTimeOfRecord Between '%date1' And '%date2' " & _
                    "ORDER By U.SampleID"

60        End If

70        sql = Replace(sql, "%date1", Format(dtStart.Value, "dd/MMM/yyyy hh:mm:ss"))
80        sql = Replace(sql, "%date2", Format(dtEnd.Value + 1, "dd/MMM/yyyy hh:mm:ss"))


90        grdBatch.Rows = 2
100       grdBatch.AddItem ""
110       grdBatch.RemoveItem 1
120       grdBatch.Rows = 1

130       Set tb = New Recordset
140       RecOpenClient 0, tb, sql
150       Do While Not tb.EOF
160           s = Val(tb!SampleID) - SysOptMicroOffset(0) & vbTab
170           Select Case tb!Pregnancy & ""
              Case "N": s = s & "Negative"
180           Case "P": s = s & "Positive"
190           Case "I": s = s & "Inconclusive"
200           Case Else: s = s & tb!Pregnancy & ""
210           End Select


220           grdBatch.AddItem s
230           grdBatch.Col = 2
240           grdBatch.Row = grdBatch.Rows - 1
250           grdBatch.CellPictureAlignment = flexAlignCenterCenter
260           Set grdBatch.CellPicture = IIf(tb!Valid, imgSquareTick.Picture, imgSquareCross.Picture)

270           tb.MoveNext
280       Loop



290       Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

300       intEL = Erl
310       strES = Err.Description
320       LogError "frmBatchPregnancy", "FillG", intEL, strES, sql

End Sub

Private Sub cmdXL_Click()
10        If grdBatch.Rows = 1 Then
20            iMsg "Nothing to export", vbInformation
30            Exit Sub
40        End If

          Dim strHeading As String
50        strHeading = "Batch Entry" & vbCr
60        strHeading = strHeading & "Pregnancy " & IIf(optUnResulted, "Unresulted Sample Requests", "Unvalidated Sample Results")
70        strHeading = strHeading & vbCr & vbCr
80        ExportFlexGrid grdBatch, Me, strHeading
End Sub

Private Sub Form_Activate()

10        FillG

End Sub

Private Sub Form_Load()
10        dtStart.Value = Date - 3
20        dtEnd.Value = Date
30        LoadListPregnancy
End Sub

Private Sub grdBatch_Click()

10        If grdBatch.Rows = 1 Then Exit Sub

20        If grdBatch.Row < 1 Then Exit Sub

30        Select Case grdBatch.Col
          Case 1:    'Pregnancy
40            CycleControlValue ListPregnancy, grdBatch

              '        Select Case grdBatch
              '            Case "": grdBatch = "Negative"
              '            Case "Negative": grdBatch = "Positive"
              '            Case "Positive": grdBatch = "Inconclusive"
              '            Case Else: grdBatch = ""
              '        End Select
50        Case 2:
60            If grdBatch.CellPicture = imgSquareTick.Picture Then
70                Set grdBatch.CellPicture = imgSquareCross.Picture
80            Else
90                Set grdBatch.CellPicture = imgSquareTick.Picture
100           End If
110       End Select

End Sub

Private Sub grdBatch_KeyPress(KeyAscii As Integer)

10        If grdBatch.Row < 1 Then Exit Sub

20        Select Case grdBatch.Col
          Case 1:    'Pregnancy
30            CycleControlValue ListPregnancy, grdBatch
              '        Select Case Chr(KeyAscii)
              '            Case "n", "N": grdBatch = "Negative"
              '            Case "p", "P": grdBatch = "Positive"
              '            Case "i", "I": grdBatch = "Inconclusive"
              '            Case Else: grdBatch = ""
              '        End Select

40        End Select

End Sub

Private Sub optNotValid_Click()

10        FillG

End Sub


Private Sub optUnResulted_Click()

10        FillG

End Sub


Private Sub LoadListPregnancy()

          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo LoadListPregnancy_Error

20        ReDim ListPregnancy(0 To 0) As String
30        ListPregnancy(0) = ""

40        sql = "SELECT * FROM Lists WHERE " & _
                "ListType = 'PG' " & _
                "ORDER BY ListOrder"
50        Set tb = New Recordset
60        RecOpenServer 0, tb, sql
70        Do While Not tb.EOF
80            ReDim Preserve ListPregnancy(0 To UBound(ListPregnancy) + 1)
90            ListPregnancy(UBound(ListPregnancy)) = tb!Text & ""
100           tb.MoveNext
110       Loop

120       Exit Sub

LoadListPregnancy_Error:

          Dim strES As String
          Dim intEL As Integer

130       intEL = Erl
140       strES = Err.Description
150       LogError "frmEditMicrobiologyNew", "LoadListPregnancy", intEL, strES, sql


End Sub

