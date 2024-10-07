VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmBatchUrine 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetAcquire - Urine Results Batch Entry"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   13080
   ClipControls    =   0   'False
   Icon            =   "frmBatchUrine.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   13080
   Begin VB.CommandButton bsearch 
      Caption         =   "&Search"
      Height          =   735
      Left            =   3480
      Picture         =   "frmBatchUrine.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   210
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker dtStart 
      Height          =   315
      Left            =   1650
      TabIndex        =   10
      Top             =   203
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
   Begin VB.OptionButton optNotValid 
      Caption         =   "Show Unvalidated Results"
      ForeColor       =   &H00008000&
      Height          =   195
      Left            =   1020
      TabIndex        =   9
      Top             =   1320
      Width           =   2265
   End
   Begin VB.OptionButton optUnResulted 
      Caption         =   "Show Unresulted Requests"
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   1020
      TabIndex        =   8
      Top             =   1080
      Value           =   -1  'True
      Width           =   2265
   End
   Begin VB.ListBox lstU 
      Height          =   4935
      Left            =   9600
      TabIndex        =   6
      Top             =   2700
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   885
      Left            =   8040
      Picture         =   "frmBatchUrine.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   630
      Width           =   1125
   End
   Begin VB.TextBox txtInput 
      Height          =   285
      Left            =   11220
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   60
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   885
      Left            =   10560
      Picture         =   "frmBatchUrine.frx":0A56
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   630
      Width           =   1125
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   885
      Left            =   11790
      Picture         =   "frmBatchUrine.frx":0D60
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   630
      Width           =   1125
   End
   Begin MSFlexGridLib.MSFlexGrid grdUrn 
      Height          =   7095
      Left            =   180
      TabIndex        =   0
      Top             =   1680
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   12515
      _Version        =   393216
      Cols            =   11
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ForeColorSel    =   16776960
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      AllowUserResizing=   1
      FormatString    =   $"frmBatchUrine.frx":106A
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   135
      Left            =   180
      TabIndex        =   7
      Top             =   8790
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
      Max             =   600
      Scrolling       =   1
   End
   Begin MSComCtl2.DTPicker dtEnd 
      Height          =   315
      Left            =   1650
      TabIndex        =   13
      Top             =   630
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
      TabIndex        =   12
      Top             =   660
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
      TabIndex        =   11
      Top             =   240
      Width           =   1350
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
      Left            =   8040
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Image imgSquareTick 
      Height          =   225
      Left            =   9975
      Picture         =   "frmBatchUrine.frx":1144
      Top             =   345
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgSquareCross 
      Height          =   225
      Left            =   10350
      Picture         =   "frmBatchUrine.frx":141A
      Top             =   345
      Visible         =   0   'False
      Width           =   210
   End
End
Attribute VB_Name = "frmBatchUrine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private UsingKeys As Boolean

Private pFromTextBox As Boolean

Dim intGridCol As Integer
Dim intGridRow As Integer

Dim ListRCC() As ListColour
Dim ListWCC() As ListColour


Private Sub bsearch_Click()
10        FillG
End Sub

Private Sub cmdCancel_Click()

10        Unload Me

End Sub

Private Sub cmdSave_Click()

10        SaveDetails

20        cmdSave.Enabled = False

30        FillG

End Sub

Sub FillG()

          Dim tb As Recordset
          Dim sql As String
          Dim Str As String
          Dim Validated As Boolean
          Dim U() As String
          Dim WCC As String
          Dim WccForecolor As Long
          Dim WccBackcolor As Long
          Dim RCC As String
          Dim RccForecolor As Long
          Dim RccBackcolor As Long

10        On Error GoTo FillG_Error

20        With grdUrn
30            .Visible = False
40            .Rows = 2
50            .AddItem ""
60            .RemoveItem 1
70            .Rows = 1
80        End With


90        If optUnResulted Then
100           sql = "SELECT R.SampleID, U.WCC, U.RCC, U.Crystals, U.Casts, U.Misc0, U.Misc1, U.Misc2, COALESCE(L.Valid, 0) As Valid FROM UrineRequests AS R " & _
                    "Left Join (Select SampleID, COALESCE(Valid, 0) As Valid From PrintValidLog Where Department = 'U') L On R.SampleID = L.SampleID " & _
                    "Left Join Urine AS U On R.SampleID = U.SampleID " & _
                    "Where COALESCE(U.WCC, '') = '' And COALESCE(U.RCC, '') = '' " & _
                    "And COALESCE(U.Crystals, '') = '' And COALESCE(U.Casts, '') = '' " & _
                    "And COALESCE(U.Misc0, '') = '' And COALESCE(U.Misc1, '') = '' And COALESCE(U.Misc2, '') = '' " & _
                    "AND R.DateTimeOfRecord Between '%date1' And '%date2' " & _
                    "ORDER By R.SampleID"

110       ElseIf optNotValid Then
120           sql = "SELECT U.SampleID, U.WCC, U.RCC, U.Crystals, U.Casts, U.Misc0, U.Misc1, U.Misc2, COALESCE(L.Valid, 0) As Valid FROM Urine AS U " & _
                    "Left Join (Select SampleID, COALESCE(Valid, 0) As Valid From PrintValidLog Where Department = 'U') L On U.SampleID = L.SampleID " & _
                    "Inner Join UrineRequests AS R On R.SampleID = U.SampleID " & _
                    "Where COALESCE(L.Valid, 0) = 0 And (COALESCE(U.WCC, '') <> '' Or COALESCE(U.RCC, '') <> '' " & _
                    "Or COALESCE(U.Crystals, '') <> '' Or COALESCE(U.Casts, '') <> '' " & _
                    "Or COALESCE(U.Misc0, '') <> '' Or COALESCE(U.Misc1, '') <> '' Or COALESCE(U.Misc2, '') = '') " & _
                    "AND R.DateTimeOfRecord Between '%date1' And '%date2' " & _
                    "ORDER By U.SampleID"

130       End If
140       sql = Replace(sql, "%date1", Format(dtStart.Value, "dd/MMM/yyyy hh:mm:ss"))
150       sql = Replace(sql, "%date2", Format(dtEnd.Value + 1, "dd/MMM/yyyy hh:mm:ss"))

160       Set tb = New Recordset
170       RecOpenClient 0, tb, sql
180       Do While Not tb.EOF
190           U = Split(tb!WCC & "", "|")
200           If UBound(U) = -1 Then
210               WCC = ""
220           ElseIf UBound(U) > 1 Then
230               WCC = U(0)
240               WccForecolor = U(1)
250               WccBackcolor = U(2)
260           Else
270               WCC = U(0)
280           End If
290           U = Split(tb!RCC & "", "|")
300           If UBound(U) = -1 Then
310               RCC = ""
320           ElseIf UBound(U) > 1 Then
330               RCC = U(0)
340               RccForecolor = U(1)
350               RccBackcolor = U(2)
360           Else
370               RCC = U(0)
380           End If
390           Validated = False
400           Validated = tb!Valid

410           Str = Val(tb!SampleID) - SysOptMicroOffset(0) & vbTab & vbTab & _
                    Trim(WCC) & vbTab & _
                    Trim(RCC & "") & vbTab & _
                    Trim(tb!Crystals & "") & vbTab & _
                    Trim(tb!Casts & "") & vbTab & _
                    Trim(tb!Misc0 & "") & vbTab & _
                    Trim(tb!Misc1 & "") & vbTab & _
                    Trim(tb!Misc2 & "") & ""
420           grdUrn.AddItem Str
430           grdUrn.Col = 2
440           grdUrn.CellBackColor = WccBackcolor
450           grdUrn.CellForeColor = WccForecolor
460           grdUrn.Col = 3
470           grdUrn.CellBackColor = RccBackcolor
480           grdUrn.CellForeColor = RccForecolor

490           grdUrn.Col = 9
500           grdUrn.Row = grdUrn.Rows - 1
510           grdUrn.CellPictureAlignment = flexAlignCenterCenter
520           Set grdUrn.CellPicture = imgSquareTick.Picture

530           grdUrn.Col = 10
540           grdUrn.Row = grdUrn.Rows - 1
550           grdUrn.CellPictureAlignment = flexAlignCenterCenter
560           Set grdUrn.CellPicture = IIf(Validated, imgSquareTick.Picture, imgSquareCross)

570           tb.MoveNext
580       Loop


590       grdUrn.Visible = True
600       Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

610       intEL = Erl
620       strES = Err.Description
630       LogError "frmBatchUrine", "FillG", intEL, strES, sql
640       grdUrn.Visible = True

End Sub

Sub FillU(ByVal CCM As String)

          Dim sql As String
          Dim tb As Recordset
          Dim ListType As String

10        On Error GoTo FillU_Error

20        Select Case CCM
          Case "Crystals": ListType = "CR"
30        Case "Casts": ListType = "CA"
40        Case "Miscellaneous": ListType = "MI"
50        End Select

60        lstU.Clear
70        sql = "SELECT * FROM Lists WHERE " & _
                "ListType = '" & ListType & "' " & _
                "ORDER BY ListOrder"
80        Set tb = New Recordset
90        RecOpenServer 0, tb, sql
100       Do While Not tb.EOF
110           lstU.AddItem tb!Text & ""
120           tb.MoveNext
130       Loop

140       lstU.AddItem "", 0

150       FixListHeight lstU

160       Exit Sub

FillU_Error:

          Dim strES As String
          Dim intEL As Integer

170       intEL = Erl
180       strES = Err.Description
190       LogError "frmBatchUrine", "FillU", intEL, strES, sql

End Sub

Private Sub cmdXL_Click()
10        If grdUrn.Rows = 1 Then
20            iMsg "Nothing to export", vbInformation
30            Exit Sub
40        End If

          Dim strHeading As String
50        strHeading = "Batch Entry" & vbCr
60        strHeading = strHeading & "Urine " & IIf(optUnResulted, "Unresulted Sample Requests", "Unvalidated Sample Results")
70        strHeading = strHeading & vbCr & vbCr

80        ExportFlexGrid grdUrn, Me, strHeading

End Sub

Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        dtStart = Date - 3
30        dtEnd = Date

40        intGridCol = 0
50        intGridRow = 0
60        pFromTextBox = False

70        lstU.Visible = False
          'grdUrn.ColWidth(1) = 200
80        FillG


90        LoadListGenericColour ListRCC(), "RR"
100       LoadListGenericColour ListWCC(), "WW"

110       Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

120       intEL = Erl
130       strES = Err.Description
140       LogError "frmBatchUrine", "Form_Load", intEL, strES

End Sub

Private Sub grdUrn_Click()

10        On Error GoTo grdUrn_Click_Error

20        If grdUrn.Rows = 1 Then Exit Sub

30        If UsingKeys Then
40            UsingKeys = False
50            Exit Sub
60        End If

70        If grdUrn.MouseRow = 0 Then Exit Sub

80        txtInput.Visible = False
90        lstU.Visible = False

100       cmdSave.Enabled = True

110       Select Case grdUrn.Col
          Case 1:    'set neg
120           grdUrn.TextMatrix(grdUrn.Row, 2) = "Nill"   'WCC
130           grdUrn.TextMatrix(grdUrn.Row, 3) = "Nill"   'RCC
140           grdUrn.TextMatrix(grdUrn.Row, 4) = "None Seen"   'Crystals
150           grdUrn.TextMatrix(grdUrn.Row, 5) = "None Seen"   'Casts
160           grdUrn.TextMatrix(grdUrn.Row, 6) = "-"   'Miscellenous

170           grdUrn.Col = 2
180           grdUrn.CellBackColor = vbWhite
190           grdUrn.CellForeColor = vbBlack
200           grdUrn.Col = 3
210           grdUrn.CellBackColor = vbWhite
220           grdUrn.CellForeColor = vbBlack
230           grdUrn.Col = 10
240           Set grdUrn.CellPicture = imgSquareTick.Picture
250           grdUrn.Col = 0

260       Case 2, 3:    'WCC, RCC
270           intGridCol = grdUrn.Col
280           intGridRow = grdUrn.Row

290           txtInput.Top = grdUrn.CellTop + grdUrn.Top
300           txtInput.Left = grdUrn.CellLeft + grdUrn.Left
310           txtInput.Width = grdUrn.ColWidth(grdUrn.Col)
320           txtInput.Visible = True
330           txtInput = grdUrn
340           txtInput.SelStart = 0
350           txtInput.SelLength = Len(txtInput)
360           txtInput.SetFocus
370       Case 4, 5, 6, 7, 8:    'Crystals,Casts and Miscellaneous
380           FillU grdUrn.TextMatrix(0, grdUrn.Col)
390           lstU.Top = grdUrn.CellTop + grdUrn.Top
400           lstU.Left = grdUrn.CellLeft + grdUrn.Left
410           lstU.Width = grdUrn.CellWidth
420           lstU.Visible = True
430           lstU.SetFocus
440       Case 9:
450           If grdUrn.CellPicture = imgSquareTick.Picture Then
460               If iMsg("Are you sure you wish remove " & vbCrLf & _
                          "Sample ID " & grdUrn.TextMatrix(grdUrn.Row, 0) & _
                          " from batch urine report?", vbYesNo + vbQuestion) = vbYes Then
470                   Set grdUrn.CellPicture = imgSquareCross.Picture
480               End If
490           Else
500               Set grdUrn.CellPicture = imgSquareTick.Picture
510           End If

520       Case 10:
530           If grdUrn.CellPicture = imgSquareTick.Picture Then
540               Set grdUrn.CellPicture = imgSquareCross.Picture
550           Else
560               Set grdUrn.CellPicture = imgSquareTick.Picture
570           End If

580       End Select

590       Exit Sub

grdUrn_Click_Error:

          Dim strES As String
          Dim intEL As Integer

600       intEL = Erl
610       strES = Err.Description
620       LogError "frmBatchUrine", "grdUrn_Click", intEL, strES

630       Exit Sub




End Sub

Private Sub SaveDetails()

          Dim sql As String
          Dim Y As Long
          Dim strSampleId As Double

10        On Error GoTo SaveDetails_Error

20        With grdUrn
30            If .Rows = 2 And .TextMatrix(1, 0) = "" Then Exit Sub

40            For Y = 1 To .Rows - 1
50                strSampleId = Val(.TextMatrix(Y, 0)) + SysOptMicroOffset(0)

60                sql = "If Exists(Select 1 From Urine " & _
                        "Where SampleID = @SampleID0 ) " & _
                        "Begin " & _
                        "Update Urine Set " & _
                        "WCC = '@WCC13', RCC = '@RCC14', Crystals = '@Crystals15', Casts = '@Casts16', " & _
                        "Misc0 = '@Misc017', Misc1 = '@Misc118', Misc2 = '@Misc219', UserName = '@UserName25' " & _
                        "Where SampleID = @SampleID0  " & _
                        "End  " & _
                        "Else " & _
                        "Begin  " & _
                        "Insert Into Urine (SampleID, WCC, RCC, Crystals, Casts, Misc0, Misc1, Misc2, UserName) Values " & _
                        "(@SampleID0, '@WCC13', '@RCC14', '@Crystals15', '@Casts16', '@Misc017', '@Misc118', '@Misc219', '@UserName25') " & _
                        "End"

70                sql = Replace(sql, "@SampleID0", strSampleId)
80                .Row = .Row: .Col = 2
90                sql = Replace(sql, "@WCC13", IIf(Trim(.TextMatrix(Y, 2)) <> "", .TextMatrix(Y, 2) & "|" & .CellForeColor & "|" & .CellBackColor, ""))
100               .Row = .Row: .Col = 3
110               sql = Replace(sql, "@RCC14", IIf(Trim(.TextMatrix(Y, 3)) <> "", .TextMatrix(Y, 3) & "|" & .CellForeColor & "|" & .CellBackColor, ""))
120               sql = Replace(sql, "@Crystals15", IIf(Trim(.TextMatrix(Y, 4)) <> "", .TextMatrix(Y, 4), ""))
130               sql = Replace(sql, "@Casts16", IIf(Trim(.TextMatrix(Y, 5)) <> "", .TextMatrix(Y, 5), ""))
140               sql = Replace(sql, "@Misc017", IIf(Trim(.TextMatrix(Y, 6)) <> "", .TextMatrix(Y, 6), ""))
150               sql = Replace(sql, "@Misc118", IIf(Trim(.TextMatrix(Y, 7)) <> "", .TextMatrix(Y, 7), ""))
160               sql = Replace(sql, "@Misc219", IIf(Trim(.TextMatrix(Y, 8)) <> "", .TextMatrix(Y, 8), ""))
170               sql = Replace(sql, "@UserName25", Username)

180               Cnxn(0).Execute sql


190               .Col = 9
200               .Row = Y
210               If .CellPicture = imgSquareCross.Picture Then
220                   sql = "UPDATE UrineRequests SET DoNotDisplayInBatchEntry = 1 " & _
                            "WHERE SampleID = '" & strSampleId & "'"
230                   Cnxn(0).Execute sql
240               End If
250               .Col = 10
260               .Row = Y
270               If .CellPicture = imgSquareTick.Picture Then
280                   UpdatePrintValidLog strSampleId, "URINE", 1, 0
290               End If
300           Next
310       End With

320       Exit Sub

SaveDetails_Error:

          Dim strES As String
          Dim intEL As Integer

330       intEL = Erl
340       strES = Err.Description
350       LogError "frmBatchUrine", "SaveDetails", intEL, strES, sql

End Sub

Private Sub lstU_DblClick()

10        UsingKeys = True
20        grdUrn = lstU
30        lstU.Visible = False



End Sub

Private Sub lstU_KeyDown(KeyCode As Integer, Shift As Integer)

      'If pFromTextBox Then
      '  pFromTextBox = False
      '  Exit Sub
      'End If

10        With lstU

              '  If Shift <> 0 Then 'press shift key to move up and down between cells
20            Select Case KeyCode

              Case vbKeyUp:
30                If Shift <> 0 Then
40                    If grdUrn.Row > 1 Then
50                        grdUrn.Row = grdUrn.Row - 1
60                        .Text = grdUrn
70                        KeyCode = 0
80                    End If
                      '        .ListIndex = 0
90                    .Top = grdUrn.CellTop + grdUrn.Top
100                   .Left = grdUrn.CellLeft + grdUrn.Left
110               End If

120           Case vbKeyDown:
130               If Shift <> 0 Then
140                   If grdUrn.Row < grdUrn.Rows - 1 Then
150                       grdUrn.Row = grdUrn.Row + 1
160                       .Text = grdUrn
170                       KeyCode = 0
180                   End If
                      '        .ListIndex = 0
190                   .Top = grdUrn.CellTop + grdUrn.Top
200                   .Left = grdUrn.CellLeft + grdUrn.Left
210               End If

220           Case vbKeyLeft:
230               If grdUrn.Col > 4 Then
240                   grdUrn.Col = grdUrn.Col - 1
250                   FillU grdUrn.TextMatrix(0, grdUrn.Col)
260                   .Top = grdUrn.CellTop + grdUrn.Top
270                   .Left = grdUrn.CellLeft + grdUrn.Left
280                   .Text = grdUrn
290                   KeyCode = 0
300               Else
                      'move to textbox
310                   UsingKeys = True
320                   lstU.Visible = False
330                   grdUrn.Col = 3
340                   txtInput.Top = grdUrn.CellTop + grdUrn.Top
350                   txtInput.Left = grdUrn.CellLeft + grdUrn.Left
360                   txtInput.Width = grdUrn.ColWidth(grdUrn.Col)
370                   txtInput.Visible = True
380                   txtInput = grdUrn
390                   txtInput.SetFocus
400                   txtInput.SelStart = 0
410                   txtInput.SelLength = Len(txtInput)
420                   UsingKeys = False
430               End If

440           Case vbKeyRight:
450               If grdUrn.Col < 8 Then
460                   grdUrn.Col = grdUrn.Col + 1
470               End If
480               FillU grdUrn.TextMatrix(0, grdUrn.Col)
490               .Top = grdUrn.CellTop + grdUrn.Top
500               .Left = grdUrn.CellLeft + grdUrn.Left
510               .Text = grdUrn
520               KeyCode = 0

530           Case 13:
540               grdUrn = .Text
550               grdUrn.Col = 9
560               Set grdUrn.CellPicture = imgSquareCross.Picture
570               If grdUrn.Col < 8 Then
580                   grdUrn.Col = grdUrn.Col + 1
590                   FillU grdUrn.TextMatrix(0, grdUrn.Col)
600                   .Top = grdUrn.CellTop + grdUrn.Top
610                   .Left = grdUrn.CellLeft + grdUrn.Left
620               Else
630                   lstU.Visible = False
640                   grdUrn.Col = 2
650                   txtInput.Top = grdUrn.CellTop + grdUrn.Top
660                   txtInput.Left = grdUrn.CellLeft + grdUrn.Left
670                   txtInput.Width = grdUrn.ColWidth(2)
680                   txtInput.Visible = True
690                   txtInput = grdUrn
700                   txtInput.SetFocus
710                   txtInput.SelStart = 0
720                   txtInput.SelLength = Len(txtInput)
730               End If
740           Case vbKeyEscape:
750               lstU.Visible = False
760               txtInput.Visible = False
770               grdUrn.SetFocus
780           End Select

790       End With

End Sub

Private Sub lstU_LostFocus()

10        If UsingKeys Then Exit Sub

          'grdUrn = lstU
20        lstU.Visible = False
30        grdUrn.Enabled = True

End Sub


Private Sub optNotValid_Click()
10        FillG
End Sub

Private Sub optUnResulted_Click()
10        FillG
End Sub

Private Sub txtInput_Click()
10        If grdUrn.Col = 2 Then
20            CycleTextBox ListWCC(), txtInput
30        ElseIf grdUrn.Col = 3 Then
40            CycleTextBox ListRCC(), txtInput
50        End If
End Sub

Private Sub txtInput_KeyUp(KeyCode As Integer, Shift As Integer)

10        On Error GoTo txtInput_KeyDown_Error

20        cmdSave.Enabled = True


30        Select Case KeyCode

          Case vbKeyDown:
40            UsingKeys = True
50            grdUrn = txtInput
60            If grdUrn.Row < grdUrn.Rows - 1 Then    'bottom row
70                grdUrn.Row = grdUrn.Row + 1
80            End If
90            txtInput.Top = grdUrn.CellTop + grdUrn.Top
100           txtInput.Left = grdUrn.CellLeft + grdUrn.Left
110           txtInput.Width = grdUrn.ColWidth(grdUrn.Col)
120           txtInput.Visible = True
130           txtInput = grdUrn
140           txtInput.SelStart = 0
150           txtInput.SelLength = Len(txtInput)
160           txtInput.SetFocus
170           UsingKeys = False

180       Case vbKeyUp:
190           UsingKeys = True
200           grdUrn = txtInput
210           If grdUrn.Row > 1 Then
220               grdUrn.Row = grdUrn.Row - 1
230           End If
240           txtInput.Top = grdUrn.CellTop + grdUrn.Top
250           txtInput.Left = grdUrn.CellLeft + grdUrn.Left
260           txtInput.Width = grdUrn.ColWidth(grdUrn.Col)
270           txtInput.Visible = True
280           txtInput = grdUrn
290           txtInput.SelStart = 0
300           txtInput.SelLength = Len(txtInput)
310           txtInput.SetFocus
320           UsingKeys = False

330       Case vbKeyLeft:
340           UsingKeys = True
350           grdUrn = txtInput
360           If grdUrn.Col = 3 Then
370               grdUrn.Col = 2
380           End If
390           txtInput.Top = grdUrn.CellTop + grdUrn.Top
400           txtInput.Left = grdUrn.CellLeft + grdUrn.Left
410           txtInput.Width = grdUrn.ColWidth(grdUrn.Col)
420           txtInput.Visible = True
430           txtInput = grdUrn
440           txtInput.SetFocus
450           txtInput.SelStart = 0
460           txtInput.SelLength = Len(txtInput)
470           UsingKeys = False

480       Case vbKeyRight, vbKeyTab:
490           UsingKeys = True
500           grdUrn = txtInput
510           If grdUrn.Col = 2 Then
520               grdUrn.Col = 3
530               txtInput.Top = grdUrn.CellTop + grdUrn.Top
540               txtInput.Left = grdUrn.CellLeft + grdUrn.Left
550               txtInput.Width = grdUrn.ColWidth(grdUrn.Col)
560               txtInput.Visible = True
570               txtInput = grdUrn
580               txtInput.SetFocus
590               txtInput.SelStart = 0
600               txtInput.SelLength = Len(txtInput)
610           Else
620               grdUrn = txtInput
                  'switch to dropdowns
630               pFromTextBox = True
640               txtInput.Visible = False
650               grdUrn.Col = 4
660               txtInput = grdUrn
670               FillU grdUrn.TextMatrix(0, 4)
680               lstU.Top = grdUrn.CellTop + grdUrn.Top
690               lstU.Left = grdUrn.CellLeft + grdUrn.Left
700               lstU = grdUrn.TextMatrix(grdUrn.Row, 4)
710               lstU.Visible = True
720               lstU.SetFocus
730           End If
740           UsingKeys = False

750       End Select
760       If grdUrn.Col = 2 Or grdUrn.Col = 3 Then    'WCC or RCC

770           If KeyCode = vbKeyF2 Then
                  '50        grdUrn = "<5"
780               txtInput = "<5"
                  '60        If grdUrn.Col = 2 Then
                  '70          grdUrn.Col = 3
                  '80          grdUrn.SetFocus
                  '90        ElseIf grdUrn.Col = 3 Then
                  '100         grdUrn.Col = 2
                  '110         grdUrn.SetFocus
                  '120       End If

790           ElseIf KeyCode = vbKeyF3 Then
                  '140       grdUrn = ">200"
800               txtInput = ">200"
                  '150       If grdUrn.Col = 2 Then
                  '160         grdUrn.Col = 3
                  '170         grdUrn.SetFocus
                  '180       ElseIf grdUrn.Col = 3 Then
                  '190         grdUrn.Col = 2
                  '200         grdUrn.SetFocus
                  '210       End If
810           End If


820           Exit Sub

830       End If

840       Exit Sub

txtInput_KeyDown_Error:

          Dim strES As String
          Dim intEL As Integer

850       intEL = Erl
860       strES = Err.Description
870       LogError "frmBatchUrine", "txtInput_KeyDown", intEL, strES

End Sub

Private Sub txtInput_LostFocus()

          Dim intOrginalRow As Integer

10        On Error GoTo txtInput_LostFocus_Error

20        If UsingKeys Then Exit Sub

30        grdUrn.TextMatrix(intGridRow, intGridCol) = txtInput
40        If intGridCol = 2 Or intGridCol = 3 Then
50            grdUrn.Col = intGridCol
60            grdUrn.CellBackColor = txtInput.BackColor
70            grdUrn.CellForeColor = txtInput.ForeColor
80        End If
90        txtInput = ""
100       txtInput.Visible = False
110       grdUrn.Enabled = True

120       intOrginalRow = grdUrn.Row
130       grdUrn.Row = intGridRow
140       If grdUrn.TextMatrix(intGridRow, intGridCol) <> "" Then
150           grdUrn.Col = 9
160           Set grdUrn.CellPicture = imgSquareCross.Picture
170           grdUrn.Col = 0
180       End If
190       grdUrn.Row = intOrginalRow
200       Exit Sub

txtInput_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

210       intEL = Erl
220       strES = Err.Description
230       LogError "frmBatchUrine", "txtInput_LostFocus", intEL, strES

End Sub

