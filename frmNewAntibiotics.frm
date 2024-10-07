VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmNewAntibiotics 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - New Antibiotics"
   ClientHeight    =   7770
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   10725
   Icon            =   "frmNewAntibiotics.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   10725
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrUp 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   10080
      Top             =   600
   End
   Begin VB.Timer tmrDown 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   9780
      Top             =   600
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   150
      TabIndex        =   5
      Top             =   180
      Width           =   7455
      Begin VB.Frame Frame4 
         Caption         =   "Penicillin Allergy"
         Height          =   555
         Left            =   5430
         TabIndex        =   20
         Top             =   810
         Width           =   1875
         Begin VB.OptionButton optPenAll 
            Caption         =   "Exclude"
            Height          =   255
            Index           =   1
            Left            =   870
            TabIndex        =   22
            Top             =   240
            Width           =   885
         End
         Begin VB.OptionButton optPenAll 
            Alignment       =   1  'Right Justify
            Caption         =   "Allow"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   21
            Top             =   270
            Value           =   -1  'True
            Width           =   675
         End
      End
      Begin VB.TextBox txtCode 
         Height          =   285
         Left            =   210
         TabIndex        =   19
         Top             =   420
         Width           =   1620
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   795
         Left            =   2490
         Picture         =   "frmNewAntibiotics.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   180
         Width           =   825
      End
      Begin VB.TextBox txtAntibiotic 
         Height          =   285
         Left            =   210
         TabIndex        =   15
         Top             =   1020
         Width           =   3105
      End
      Begin VB.Frame Frame2 
         Caption         =   "Pregnancy"
         Height          =   555
         Left            =   5430
         TabIndex        =   12
         Top             =   210
         Width           =   1875
         Begin VB.OptionButton optPreg 
            Alignment       =   1  'Right Justify
            Caption         =   "Allow"
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   14
            Top             =   240
            Value           =   -1  'True
            Width           =   705
         End
         Begin VB.OptionButton optPreg 
            Caption         =   "Exclude"
            Height          =   195
            Index           =   1
            Left            =   870
            TabIndex        =   13
            Top             =   240
            Width           =   945
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Out Patients"
         Height          =   555
         Left            =   3510
         TabIndex        =   9
         Top             =   210
         Width           =   1875
         Begin VB.OptionButton optOP 
            Alignment       =   1  'Right Justify
            Caption         =   "Allow"
            Height          =   255
            Index           =   0
            Left            =   150
            TabIndex        =   11
            Top             =   240
            Value           =   -1  'True
            Width           =   675
         End
         Begin VB.OptionButton optOP 
            Caption         =   "Exclude"
            Height          =   255
            Index           =   1
            Left            =   870
            TabIndex        =   10
            Top             =   240
            Width           =   885
         End
      End
      Begin VB.Frame Children 
         Caption         =   "Children"
         Height          =   555
         Left            =   3510
         TabIndex        =   6
         Top             =   810
         Width           =   1875
         Begin VB.OptionButton optChildren 
            Caption         =   "Exclude"
            Height          =   255
            Index           =   1
            Left            =   870
            TabIndex        =   8
            Top             =   240
            Width           =   915
         End
         Begin VB.OptionButton optChildren 
            Alignment       =   1  'Right Justify
            Caption         =   "Allow"
            Height          =   255
            Index           =   0
            Left            =   150
            TabIndex        =   7
            Top             =   240
            Value           =   -1  'True
            Width           =   675
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Antibiotic"
         Height          =   195
         Left            =   210
         TabIndex        =   18
         Top             =   780
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Code"
         Height          =   195
         Left            =   210
         TabIndex        =   17
         Top             =   210
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdMoveDown 
      Caption         =   "Move Down"
      Enabled         =   0   'False
      Height          =   825
      Left            =   9780
      Picture         =   "frmNewAntibiotics.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4440
      Width           =   795
   End
   Begin VB.CommandButton cmdMoveUp 
      Caption         =   "Move Up"
      Enabled         =   0   'False
      Height          =   825
      Left            =   9780
      Picture         =   "frmNewAntibiotics.frx":0B8E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3600
      Width           =   795
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   735
      Left            =   9780
      Picture         =   "frmNewAntibiotics.frx":0FD0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6930
      Width           =   885
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   735
      Left            =   9780
      Picture         =   "frmNewAntibiotics.frx":12DA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2220
      Width           =   795
   End
   Begin MSFlexGridLib.MSFlexGrid grdAB 
      Height          =   5925
      Left            =   150
      TabIndex        =   0
      Top             =   1740
      Width           =   9525
      _ExtentX        =   16801
      _ExtentY        =   10451
      _Version        =   393216
      Cols            =   7
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
      FormatString    =   "<Code       |<Antibiotic Name                  |^Pregnancy|^Out-Patients|^Children|^Pen.All |<Report Name              "
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
   Begin VB.Image imgSquareCross 
      Height          =   225
      Left            =   10140
      Picture         =   "frmNewAntibiotics.frx":15E4
      Top             =   1740
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgSquareTick 
      Height          =   225
      Left            =   10140
      Picture         =   "frmNewAntibiotics.frx":18BA
      Top             =   1260
      Visible         =   0   'False
      Width           =   210
   End
End
Attribute VB_Name = "frmNewAntibiotics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FireCounter As Long
Private Sub FireDown()

          Dim n As Long
          Dim s As String
          Dim X As Long
          Dim TickSave(2 To 5) As Boolean
          Dim VisibleRows As Long

10        On Error GoTo FireDown_Error

20        With grdAB
30            If .Row = .Rows - 1 Then Exit Sub

40            n = .Row

50            FireCounter = FireCounter + 1
60            If FireCounter > 5 Then
70                tmrDown.Interval = 100
80            End If

90            VisibleRows = .Height \ .RowHeight(1) - 1

100           .Visible = False

110           s = .TextMatrix(n, 0) & vbTab & .TextMatrix(n, 1)
120           For X = 2 To 5
130               .Col = X
140               TickSave(X) = .CellPicture = imgSquareTick.Picture
150           Next

160           .RemoveItem n
170           If n < .Rows Then
180               .AddItem s, n + 1
190               .Row = n + 1
200           Else
210               .AddItem s
220               .Row = .Rows - 1
230           End If

240           For X = 0 To .Cols - 1
250               .Col = X
260               .CellBackColor = vbYellow
270           Next

280           For X = 2 To 5
290               .Col = X
300               .CellPictureAlignment = flexAlignCenterCenter
310               Set .CellPicture = IIf(TickSave(X), imgSquareTick.Picture, imgSquareCross.Picture)
320           Next

330           If Not .RowIsVisible(.Row) Or .Row = .Rows - 1 Then
340               If .Row - VisibleRows + 1 > 0 Then
350                   .TopRow = .Row - VisibleRows + 1
360               End If
370           End If

380           .Visible = True

390       End With

400       cmdSave.Enabled = True

410       Exit Sub

FireDown_Error:

          Dim strES As String
          Dim intEL As Integer



420       intEL = Erl
430       strES = Err.Description
440       LogError "frmNewAntibiotics", "FireDown", intEL, strES


End Sub
Private Sub FireUp()

          Dim n As Long
          Dim s As String
          Dim X As Long
          Dim TickSave(2 To 5) As Boolean

10        On Error GoTo FireUp_Error

20        With grdAB
30            If .Row = 1 Then Exit Sub

40            FireCounter = FireCounter + 1
50            If FireCounter > 5 Then
60                tmrUp.Interval = 100
70            End If

80            n = .Row

90            .Visible = False

100           s = .TextMatrix(n, 0) & vbTab & .TextMatrix(n, 1)

110           For X = 2 To 5
120               .Col = X
130               TickSave(X) = .CellPicture = imgSquareTick.Picture
140           Next

150           .RemoveItem n
160           .AddItem s, n - 1

170           .Row = n - 1
180           For X = 0 To .Cols - 1
190               .Col = X
200               .CellBackColor = vbYellow
210           Next

220           For X = 2 To 5
230               .Col = X
240               .CellPictureAlignment = flexAlignCenterCenter
250               Set .CellPicture = IIf(TickSave(X), imgSquareTick.Picture, imgSquareCross.Picture)
260           Next

270           If Not .RowIsVisible(.Row) Then
280               .TopRow = .Row
290           End If

300           .Visible = True

310       End With

320       cmdSave.Enabled = True

330       Exit Sub

FireUp_Error:

          Dim strES As String
          Dim intEL As Integer



340       intEL = Erl
350       strES = Err.Description
360       LogError "frmNewAntibiotics", "FireUp", intEL, strES


End Sub





Private Sub FillG()

          Dim tb As Recordset
          Dim sql As String
          Dim s As String

10        On Error GoTo FillG_Error

20        grdAB.Rows = 2
30        grdAB.AddItem ""
40        grdAB.RemoveItem 1

50        sql = "SELECT * from Antibiotics " & _
                "order by ListOrder asc"
60        Set tb = New Recordset
70        RecOpenClient 0, tb, sql
80        Do While Not tb.EOF
90            s = tb!Code & "" & vbTab & _
                  tb!AntibioticName & "" & vbTab & _
                  vbTab & vbTab & vbTab & vbTab & tb!ReportName & ""
100           grdAB.AddItem s
110           grdAB.Row = grdAB.Rows - 1

120           grdAB.Col = 2
130           grdAB.CellPictureAlignment = flexAlignCenterCenter
140           If Not IsNull(tb!AllowIfPregnant) Then
150               If tb!AllowIfPregnant Then
160                   Set grdAB.CellPicture = imgSquareTick.Picture
170               Else
180                   Set grdAB.CellPicture = imgSquareCross.Picture
190               End If
200           Else
210               Set grdAB.CellPicture = imgSquareTick.Picture
220           End If

230           grdAB.Col = 3
240           grdAB.CellPictureAlignment = flexAlignCenterCenter
250           If Not IsNull(tb!AllowIfOutPatient) Then
260               If tb!AllowIfOutPatient Then
270                   Set grdAB.CellPicture = imgSquareTick.Picture
280               Else
290                   Set grdAB.CellPicture = imgSquareCross.Picture
300               End If
310           Else
320               Set grdAB.CellPicture = imgSquareTick.Picture
330           End If

340           grdAB.Col = 4
350           grdAB.CellPictureAlignment = flexAlignCenterCenter
360           If Not IsNull(tb!AllowIfChild) Then
370               If tb!AllowIfChild Then
380                   Set grdAB.CellPicture = imgSquareTick.Picture
390               Else
400                   Set grdAB.CellPicture = imgSquareCross.Picture
410               End If
420           Else
430               Set grdAB.CellPicture = imgSquareTick.Picture
440           End If

450           grdAB.Col = 5
460           grdAB.CellPictureAlignment = flexAlignCenterCenter
470           If Not IsNull(tb!AllowIfPenAll) Then
480               If tb!AllowIfPenAll Then
490                   Set grdAB.CellPicture = imgSquareTick.Picture
500               Else
510                   Set grdAB.CellPicture = imgSquareCross.Picture
520               End If
530           Else
540               Set grdAB.CellPicture = imgSquareTick.Picture
550           End If

560           tb.MoveNext
570       Loop

580       If grdAB.Rows > 2 Then
590           grdAB.RemoveItem 1
600       End If

610       Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer



620       intEL = Erl
630       strES = Err.Description
640       LogError "frmNewAntibiotics", "FillG", intEL, strES, sql


End Sub

Private Sub cmdadd_Click()

          Dim n As Integer
          Dim s As String

10        On Error GoTo cmdadd_Click_Error

20        If Trim$(txtAntibiotic) = "" Then Exit Sub
30        If Trim$(txtCode) = "" Then Exit Sub

40        For n = 1 To grdAB.Rows - 1
50            If UCase$(Trim$(txtCode)) = UCase$(Trim$(grdAB.TextMatrix(n, 0))) Then
60                iMsg "Code already used.", vbCritical
70                txtCode = ""
80                txtCode.SetFocus
90                Exit Sub
100           End If
110       Next

120       s = txtCode & vbTab & txtAntibiotic
130       grdAB.AddItem s
140       grdAB.Row = grdAB.Rows - 1

150       grdAB.Col = 2
160       grdAB.CellPictureAlignment = flexAlignCenterCenter
170       If optPreg(0) Then
180           Set grdAB.CellPicture = imgSquareTick.Picture
190       Else
200           Set grdAB.CellPicture = imgSquareCross.Picture
210       End If

220       grdAB.Col = 3
230       grdAB.CellPictureAlignment = flexAlignCenterCenter
240       If optOP(0) Then
250           Set grdAB.CellPicture = imgSquareTick.Picture
260       Else
270           Set grdAB.CellPicture = imgSquareCross.Picture
280       End If

290       grdAB.Col = 4
300       grdAB.CellPictureAlignment = flexAlignCenterCenter
310       If optChildren(0) Then
320           Set grdAB.CellPicture = imgSquareTick.Picture
330       Else
340           Set grdAB.CellPicture = imgSquareCross.Picture
350       End If

360       grdAB.Col = 5
370       grdAB.CellPictureAlignment = flexAlignCenterCenter
380       If optPenAll(0) Then
390           Set grdAB.CellPicture = imgSquareTick.Picture
400       Else
410           Set grdAB.CellPicture = imgSquareCross.Picture
420       End If

430       txtAntibiotic = ""
440       txtCode = ""
450       cmdSave.Enabled = True

460       Exit Sub

cmdadd_Click_Error:

          Dim strES As String
          Dim intEL As Integer



470       intEL = Erl
480       strES = Err.Description
490       LogError "frmNewAntibiotics", "cmdAdd_Click", intEL, strES


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
110       LogError "frmNewAntibiotics", "cmdCancel_Click", intEL, strES


End Sub

Private Sub cmdMoveDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo cmdMoveDown_MouseDown_Error

20        FireDown

30        tmrDown.Interval = 250
40        FireCounter = 0

50        tmrDown.Enabled = True

60        Exit Sub

cmdMoveDown_MouseDown_Error:

          Dim strES As String
          Dim intEL As Integer



70        intEL = Erl
80        strES = Err.Description
90        LogError "frmNewAntibiotics", "cmdMoveDown_MouseDown", intEL, strES


End Sub


Private Sub cmdMoveDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo cmdMoveDown_MouseUp_Error

20        tmrDown.Enabled = False

30        Exit Sub

cmdMoveDown_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmNewAntibiotics", "cmdMoveDown_MouseUp", intEL, strES


End Sub


Private Sub cmdMoveUp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo cmdMoveUp_MouseDown_Error

20        FireUp

30        tmrUp.Interval = 250
40        FireCounter = 0

50        tmrUp.Enabled = True

60        Exit Sub

cmdMoveUp_MouseDown_Error:

          Dim strES As String
          Dim intEL As Integer



70        intEL = Erl
80        strES = Err.Description
90        LogError "frmNewAntibiotics", "cmdMoveUp_MouseDown", intEL, strES


End Sub


Private Sub cmdMoveUp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo cmdMoveUp_MouseUp_Error

20        tmrUp.Enabled = False

30        Exit Sub

cmdMoveUp_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmNewAntibiotics", "cmdMoveUp_MouseUp", intEL, strES


End Sub


Private Sub cmdSave_Click()

          Dim tb As Recordset
          Dim sql As String
          Dim n As Long

10        On Error GoTo cmdSave_Click_Error

20        For n = 1 To grdAB.Rows - 1
30            If grdAB.TextMatrix(n, 1) <> "" Then
40                sql = "SELECT * from Antibiotics WHERE " & _
                        "AntibioticName = '" & grdAB.TextMatrix(n, 1) & "'"

50                Set tb = New Recordset
60                RecOpenClient 0, tb, sql
70                If tb.EOF Then
80                    tb.AddNew
90                End If
100               tb!Code = grdAB.TextMatrix(n, 0)
110               tb!AntibioticName = grdAB.TextMatrix(n, 1)
120               tb!ListOrder = n
130               grdAB.Row = n
140               grdAB.Col = 2
150               tb!AllowIfPregnant = IIf(grdAB.CellPicture = imgSquareTick.Picture, 1, 0)
160               grdAB.Col = 3
170               tb!AllowIfOutPatient = IIf(grdAB.CellPicture = imgSquareTick.Picture, 1, 0)
180               grdAB.Col = 4
190               tb!AllowIfChild = IIf(grdAB.CellPicture = imgSquareTick.Picture, 1, 0)
200               grdAB.Col = 5
210               tb!AllowIfPenAll = IIf(grdAB.CellPicture = imgSquareTick.Picture, 1, 0)
220               tb!ReportName = grdAB.TextMatrix(n, 6)
230               tb.Update
240           End If
250       Next

260       cmdSave.Enabled = False

270       Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer



280       intEL = Erl
290       strES = Err.Description
300       LogError "frmNewAntibiotics", "cmdsave_Click", intEL, strES, sql


End Sub

Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        FillG

30        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmNewAntibiotics", "Form_Load", intEL, strES


End Sub

Private Sub grdAB_Click()

          Static SortOrder As Boolean
          Dim X As Long
          Dim Y As Long
          Dim ySave As Long
          Dim xSave As Long

10        On Error GoTo grdAB_Click_Error

20        With grdAB
30            ySave = .Row
40            xSave = .Col
50            .Visible = False
60            .Col = 0
70            For Y = 1 To .Rows - 1
80                .Row = Y
90                If .CellBackColor = vbYellow Then
100                   For X = 0 To .Cols - 1
110                       .Col = X
120                       .CellBackColor = 0
130                   Next
140                   Exit For
150               End If
160           Next
170           .Row = ySave
180           .Col = xSave
190           .Visible = True
200       End With

210       If grdAB.MouseRow = 0 Then
220           If SortOrder Then
230               grdAB.Sort = flexSortGenericAscending
240           Else
250               grdAB.Sort = flexSortGenericDescending
260           End If
270           SortOrder = Not SortOrder
280           Exit Sub
290       End If

300       Select Case grdAB.Col
          Case 0:
310           grdAB.Enabled = False
320           grdAB.TextMatrix(grdAB.Row, 0) = Trim$(UCase$(iBOX("Code for " & Trim$(grdAB.TextMatrix(grdAB.Row, 1)) & " ?", "Antibiotic Code", grdAB.TextMatrix(grdAB.Row, 0))))
330           grdAB.Enabled = True
340           cmdSave.Enabled = True
350       Case 1:
360           For X = 0 To grdAB.Cols - 1
370               grdAB.Col = X
380               grdAB.CellBackColor = vbYellow
390           Next
400           cmdMoveUp.Enabled = True
410           cmdMoveDown.Enabled = True
420       Case 2, 3, 4, 5:
430           If grdAB.CellPicture = imgSquareTick.Picture Then
440               Set grdAB.CellPicture = imgSquareCross.Picture
450           Else
460               Set grdAB.CellPicture = imgSquareTick.Picture
470           End If
480       Case 6:
490           cmdSave.Enabled = True
500           grdAB.Enabled = False
510           grdAB = iBOX("Report Name", , grdAB)
520           grdAB.Enabled = True
530           cmdSave.Enabled = True
540           cmdSave.Visible = True

550       End Select

560       Exit Sub

grdAB_Click_Error:

          Dim strES As String
          Dim intEL As Integer



570       intEL = Erl
580       strES = Err.Description
590       LogError "frmNewAntibiotics", "grdAB_Click", intEL, strES


End Sub

Private Sub tmrDown_Timer()

10        On Error GoTo tmrDown_Timer_Error

20        FireDown

30        Exit Sub

tmrDown_Timer_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmNewAntibiotics", "tmrDown_Timer", intEL, strES


End Sub


Private Sub tmrUp_Timer()

10        On Error GoTo tmrUp_Timer_Error

20        FireUp

30        Exit Sub

tmrUp_Timer_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmNewAntibiotics", "tmrUp_Timer", intEL, strES


End Sub


