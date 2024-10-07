VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmBatchPrinting 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetAcquire - Batch Printing"
   ClientHeight    =   8460
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7155
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8460
   ScaleWidth      =   7155
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optDateValidated 
      Caption         =   "Filter by Date Validated"
      ForeColor       =   &H00008000&
      Height          =   195
      Left            =   180
      TabIndex        =   13
      Top             =   1140
      Value           =   -1  'True
      Width           =   2265
   End
   Begin VB.OptionButton optDateRequested 
      Caption         =   "Filter by Date Requested"
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   180
      TabIndex        =   12
      Top             =   1440
      Width           =   2265
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   1000
      Index           =   1
      Left            =   5785
      Picture         =   "frmBatchPrinting.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Tag             =   "bprint"
      ToolTipText     =   "Print Result"
      Top             =   7320
      Width           =   1100
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   1000
      Index           =   0
      Left            =   2245
      Picture         =   "frmBatchPrinting.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   10
      Tag             =   "bprint"
      ToolTipText     =   "Print Result"
      Top             =   7320
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   1000
      Left            =   5785
      Picture         =   "frmBatchPrinting.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   180
      Width           =   1100
   End
   Begin VB.CommandButton bsearch 
      Caption         =   "&Search"
      Default         =   -1  'True
      Height          =   1000
      Left            =   3480
      Picture         =   "frmBatchPrinting.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   180
      Width           =   1100
   End
   Begin MSFlexGridLib.MSFlexGrid grdBatchPrint 
      Height          =   5055
      Index           =   0
      Left            =   180
      TabIndex        =   1
      Top             =   2160
      Width           =   3165
      _ExtentX        =   5583
      _ExtentY        =   8916
      _Version        =   393216
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
      FormatString    =   "<Sample ID             |^Print         "
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
   Begin MSComCtl2.DTPicker dtStart 
      Height          =   315
      Left            =   1650
      TabIndex        =   3
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
      Format          =   59768833
      CurrentDate     =   40267
   End
   Begin MSComCtl2.DTPicker dtEnd 
      Height          =   315
      Left            =   1650
      TabIndex        =   4
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
      Format          =   59768833
      CurrentDate     =   40267
   End
   Begin MSFlexGridLib.MSFlexGrid grdBatchPrint 
      Height          =   5055
      Index           =   1
      Left            =   3720
      TabIndex        =   7
      Top             =   2160
      Width           =   3165
      _ExtentX        =   5583
      _ExtentY        =   8916
      _Version        =   393216
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
      FormatString    =   "<Sample ID             |^Print         "
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
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Faeces samples to print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   3720
      TabIndex        =   9
      Top             =   1800
      Width           =   3165
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Urine samples to print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   180
      TabIndex        =   8
      Top             =   1800
      Width           =   3165
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
      TabIndex        =   6
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
      TabIndex        =   5
      Top             =   210
      Width           =   1350
   End
   Begin VB.Image imgSquareTick 
      Height          =   225
      Left            =   4860
      Picture         =   "frmBatchPrinting.frx":0D60
      Top             =   300
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgSquareCross 
      Height          =   225
      Left            =   5235
      Picture         =   "frmBatchPrinting.frx":1036
      Top             =   300
      Visible         =   0   'False
      Width           =   210
   End
End
Attribute VB_Name = "frmBatchPrinting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub bsearch_Click()
10        FillG
End Sub

Private Sub cmdCancel_Click()
10        Unload Me
End Sub



Sub FillG()

          Dim tb As Recordset
          Dim sql As String
          Dim s As String

10        On Error GoTo FillG_Error

20        sql = "Select U.SampleID, COALESCE(L.Printed, 0) As Printed From Urine U " & _
                "Left Join (Select SampleID, COALESCE(Valid, 0) As Valid, ValidatedDateTime, Printed From PrintValidLog Where Department = 'U') L On U.SampleID = L.SampleID " & _
                "Inner Join UrineRequests R On U.SampleID = R.SampleID " & _
                "Where COALESCE(L.Valid, 0) = 1 "
30        If optDateValidated Then
40            sql = sql & "AND L.ValidatedDateTime Between '%date1' And '%date2' " & _
                    "ORDER By U.SampleID"
50        ElseIf optDateRequested Then
60            sql = sql & "AND R.DateTimeOfRecord Between '%date1' And '%date2' " & _
                    "ORDER By U.SampleID"
70        End If


80        sql = Replace(sql, "%date1", Format(dtStart.Value, "dd/MMM/yyyy hh:mm:ss"))
90        sql = Replace(sql, "%date2", Format(dtEnd.Value + 1, "dd/MMM/yyyy hh:mm:ss"))


100       grdBatchPrint(0).Rows = 2
110       grdBatchPrint(0).AddItem ""
120       grdBatchPrint(0).RemoveItem 1
130       grdBatchPrint(0).Rows = 1

140       Set tb = New Recordset
150       RecOpenClient 0, tb, sql
160       Do While Not tb.EOF
170           s = Val(tb!SampleID) - SysOptMicroOffset(0)
180           grdBatchPrint(0).AddItem s
190           grdBatchPrint(0).Col = 1
200           grdBatchPrint(0).Row = grdBatchPrint(0).Rows - 1
210           grdBatchPrint(0).CellPictureAlignment = flexAlignCenterCenter
220           Set grdBatchPrint(0).CellPicture = IIf(tb!Printed, imgSquareCross.Picture, imgSquareTick.Picture)

230           tb.MoveNext
240       Loop


250       sql = "Select F.SampleID, COALESCE(L.Printed, 0) As Printed From Faeces F " & _
                "Left Join PrintValidLog L On F.SampleID = L.SampleID " & _
                "Inner Join FaecalRequests R On R.SampleID = F.SampleID " & _
                "Where COALESCE(L.Valid, 0) = 1 "
260       If optDateValidated Then
270           sql = sql & "AND L.ValidatedDateTime Between '%date1' And '%date2' " & _
                    "ORDER By F.SampleID"
280       ElseIf optDateRequested Then
290           sql = sql & "AND R.DateTimeOfRecord Between '%date1' And '%date2' " & _
                    "ORDER By F.SampleID"
300       End If

310       sql = Replace(sql, "%date1", Format(dtStart.Value, "dd/MMM/yyyy hh:mm:ss"))
320       sql = Replace(sql, "%date2", Format(dtEnd.Value + 1, "dd/MMM/yyyy hh:mm:ss"))


330       grdBatchPrint(1).Rows = 2
340       grdBatchPrint(1).AddItem ""
350       grdBatchPrint(1).RemoveItem 1
360       grdBatchPrint(1).Rows = 1

370       Set tb = New Recordset
380       RecOpenClient 0, tb, sql
390       Do While Not tb.EOF
400           s = Val(tb!SampleID) - SysOptMicroOffset(0)
410           grdBatchPrint(1).AddItem s
420           grdBatchPrint(1).Col = 1
430           grdBatchPrint(1).Row = grdBatchPrint(1).Rows - 1
440           grdBatchPrint(1).CellPictureAlignment = flexAlignCenterCenter
450           Set grdBatchPrint(1).CellPicture = IIf(tb!Printed, imgSquareCross.Picture, imgSquareTick.Picture)

460           tb.MoveNext
470       Loop

480       Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

490       intEL = Erl
500       strES = Err.Description
510       LogError "frmBatchPregnancy", "FillG", intEL, strES, sql

End Sub



Private Sub cmdPrint_Click(Index As Integer)

          Dim tb As Recordset
          Dim sql As String
          Dim i As Integer
          Dim Ward As String
          Dim Clinician As String
          Dim GP As String
          Dim SampleID As String

10        On Error GoTo cmdPrint_Click_Error



20        With grdBatchPrint(Index)
30            For i = 1 To .Rows - 1
40                SampleID = .TextMatrix(i, 0) + SysOptMicroOffset(0)
50                .Row = i
60                If .CellPicture = imgSquareTick Then
70                    sql = "Select Ward, Gp, Clinician From Demographics Where SampleID = '%sampleid'"
80                    sql = Replace(sql, "%sampleid", SampleID)
90                    Set tb = New Recordset
100                   RecOpenClient 0, tb, sql
110                   If Not tb.EOF Then
120                       Ward = tb!Ward & ""
130                       GP = tb!GP & ""
140                       Clinician = tb!Clinician & ""
150                   End If
160                   sql = "Select * from PrintPending where " & _
                            "Department = 'N' " & _
                            "and SampleID = '" & SampleID & "'"
170                   Set tb = New Recordset
180                   RecOpenClient 0, tb, sql
190                   If tb.EOF Then
200                       tb.AddNew
210                   End If
220                   tb!SampleID = SampleID
230                   tb!Ward = Ward
240                   tb!Clinician = Clinician
250                   tb!GP = GP
260                   tb!Department = "N"
270                   tb!Initiator = Username
280                   tb!UsePrinter = ""
290                   tb.Update
300               End If
310           Next i
320       End With

330       Exit Sub

cmdPrint_Click_Error:

          Dim strES As String
          Dim intEL As Integer

340       intEL = Erl
350       strES = Err.Description
360       LogError "frmBatchPrinting", "cmdPrint_Click", intEL, strES, sql

End Sub

Private Sub Form_Activate()

10        FillG

End Sub

Private Sub Form_Load()
10        dtStart.Value = Date - 3
20        dtEnd.Value = Date
End Sub


Private Sub grdBatchprint_Click(Index As Integer)

10        If grdBatchPrint(Index).Rows = 1 Then Exit Sub
20        If grdBatchPrint(Index).Row = 0 Then Exit Sub

30        Select Case grdBatchPrint(Index).Col
          Case 1:
40            If grdBatchPrint(Index).CellPicture = imgSquareTick.Picture Then
50                Set grdBatchPrint(Index).CellPicture = imgSquareCross.Picture
60            Else
70                Set grdBatchPrint(Index).CellPicture = imgSquareTick.Picture
80            End If
90        End Select

End Sub

Private Sub optDateRequested_Click()
10        FillG
End Sub

Private Sub optDateValidated_Click()
10        FillG
End Sub
