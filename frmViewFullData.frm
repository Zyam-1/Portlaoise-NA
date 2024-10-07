VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmViewFullDataHBA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetAcquire - View Full HbA1c Data"
   ClientHeight    =   6990
   ClientLeft      =   2040
   ClientTop       =   510
   ClientWidth     =   4950
   Icon            =   "frmViewFullData.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   4950
   Begin MSFlexGridLib.MSFlexGrid grdPoint 
      Height          =   2175
      Left            =   135
      TabIndex        =   6
      Top             =   990
      Width           =   4650
      _ExtentX        =   8202
      _ExtentY        =   3836
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      BackColorBkg    =   -2147483633
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   0
      GridLinesFixed  =   0
      ScrollBars      =   2
      FormatString    =   "<                 |<                |>Sec   |>Area       |>      %"
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
   Begin VB.CommandButton bCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   735
      Left            =   3510
      Picture         =   "frmViewFullData.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   135
      Width           =   1245
   End
   Begin VB.PictureBox picGraph 
      AutoRedraw      =   -1  'True
      DrawWidth       =   2
      Height          =   3075
      Left            =   135
      ScaleHeight     =   310.185
      ScaleMode       =   0  'User
      ScaleWidth      =   235.149
      TabIndex        =   0
      Top             =   3465
      Width           =   4665
   End
   Begin VB.Label Label4 
      Caption         =   "0                      60                      120                               180"
      Height          =   240
      Left            =   135
      TabIndex        =   9
      Top             =   6570
      Width           =   4650
   End
   Begin VB.Label lblHbA1c 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   945
      TabIndex        =   8
      Top             =   225
      Width           =   645
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "HbA1c"
      Height          =   195
      Left            =   360
      TabIndex        =   7
      Top             =   315
      Width           =   495
   End
   Begin VB.Label lblHbF 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2250
      TabIndex        =   5
      Top             =   225
      Width           =   645
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "HbF"
      Height          =   195
      Left            =   1845
      TabIndex        =   4
      Top             =   315
      Width           =   300
   End
   Begin VB.Label lblHbA1 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   945
      TabIndex        =   3
      Top             =   630
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "HbA1"
      Height          =   195
      Left            =   450
      TabIndex        =   2
      Top             =   675
      Width           =   405
   End
End
Attribute VB_Name = "frmViewFullDataHBA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mSampleID As String
Private mhba1c As String

Private Activated As Boolean


Private Type Menarini
    PeakTimeSecs As Long
    Analyte As String
    Area As String
    AreaRatio As String
End Type
Private HbData(1 To 10) As Menarini

Private Ssecs As Long
Private Fsecs As Long

Private Sub DrawGraph()

          Dim n As Single
          Dim Position As Long
          Dim gdArray(0 To 179) As Single
          Dim HbStart As Single
          Dim HbEnd As Single
          Dim Max As Long
          Dim tb As New Recordset
          Dim sql As String
          Dim Block3 As String
          Dim TempY As Long
          Dim TempX As Long
          Dim blnF As Boolean
          Dim blnC As Boolean


10        On Error GoTo DrawGraph_Error

20        sql = "SELECT * from HbA1c WHERE " & _
                "SampleID = '" & mSampleID & "'"
30        Set tb = New Recordset
40        RecOpenClient 0, tb, sql

50        If tb.EOF Then Exit Sub

60        lblHbA1 = tb!hba1
70        lblHbF = tb!hbf
80        lblHbA1c = mhba1c

90        FillG

100       Block3 = Mid$(tb!block & "", 269)
          'Block3 = Mid$(tb!block & "", 433)
110       If Block3 = "" Then Exit Sub

120       Position = 0
130       For n = 2 To 158 Step 4
140           gdArray(Position) = Val(Mid$(Block3, n, 3))
150           Position = Position + 1
160       Next
170       For n = 165 To 321 Step 4
180           gdArray(Position) = Val(Mid$(Block3, n, 3))
190           Position = Position + 1
200       Next
210       For n = 328 To 477 Step 4
220           gdArray(Position) = Val(Mid$(Block3, n, 3))
230           Position = Position + 1
240       Next
250       For n = 491 To 727 Step 4
260           gdArray(Position) = Val(Mid$(Block3, n, 3))
270           Position = Position + 1
280       Next

290       Max = 0
300       For n = 0 To 179
310           If gdArray(n) > Max Then
320               Max = gdArray(n)
330           End If
340       Next
350       picGraph.ScaleHeight = Max

360       HbStart = Val(Mid$(Block3, 828, 3))
370       HbEnd = Val(Mid$(Block3, 832, 3))

380       HbStart = (HbStart * 115) / 180
390       HbEnd = (HbEnd * 115) / 180

400       blnF = False
410       blnC = False

420       For n = 0 To 115    '179
430           If n = 0 Then
440               picGraph.PSet (n * 179 / 115, picGraph.ScaleHeight - gdArray(n)), vbBlack
450           Else
460               Debug.Print Fsecs, n * 179 / 115
470               If Fsecs <= n * 179 / 115 And Not blnF Then
480                   blnF = True
490                   TempY = picGraph.CurrentY
500                   TempX = picGraph.CurrentX
510                   picGraph.CurrentY = picGraph.CurrentY - (1.5 * picGraph.TextHeight("H"))
520                   picGraph.CurrentX = n * 179 / 115 - picGraph.TextWidth("I")
530                   picGraph.ForeColor = vbBlack
540                   picGraph.Print "F";
550                   picGraph.CurrentY = TempY
560                   picGraph.CurrentX = TempX
570               End If
580               If Ssecs <= n * 179 / 115 And Not blnC Then
590                   blnC = True
600                   TempY = picGraph.CurrentY
610                   TempX = picGraph.CurrentX
620                   picGraph.CurrentY = picGraph.CurrentY - (1.5 * picGraph.TextHeight("H"))
630                   picGraph.CurrentX = n * 179 / 115 - picGraph.TextWidth("W")
640                   picGraph.ForeColor = vbBlack
650                   picGraph.Print "s-A1c";
660                   picGraph.CurrentY = TempY
670                   picGraph.CurrentX = TempX
680               End If

690               If n > HbStart And n < HbEnd Then
700                   picGraph.Line -(n * 179 / 115, picGraph.ScaleHeight - gdArray(n)), vbYellow
710               Else
720                   picGraph.Line -(n * 179 / 115, picGraph.ScaleHeight - gdArray(n)), vbBlack
730               End If
740           End If
750       Next
760       picGraph.Line (HbStart * 179 / 115, picGraph.ScaleHeight - gdArray(HbStart))-(HbEnd * 179 / 115, picGraph.ScaleHeight - gdArray(HbEnd)), vbYellow

770       On Error GoTo 0


780       Exit Sub

DrawGraph_Error:

          Dim strES As String
          Dim intEL As Integer



790       intEL = Erl
800       strES = Err.Description
810       LogError "frmViewFullDataHBA", "DrawGraph", intEL, strES, sql


End Sub

Private Sub FillG()

          Dim tb As New Recordset
          Dim sql As String
          Dim Block2 As String
          Dim n As Long
          Dim Y As Long


10        On Error GoTo FillG_Error

20        grdPoint.Rows = 2
30        grdPoint.AddItem ""
40        grdPoint.RemoveItem 1

50        For n = 1 To 10
60            HbData(n).Analyte = ""
70            HbData(n).Area = ""
80            HbData(n).AreaRatio = ""
90            HbData(n).PeakTimeSecs = 0
100       Next
110       Fsecs = 0
120       Ssecs = 0

130       sql = "SELECT * from HbA1c WHERE " & _
                "SampleID = '" & mSampleID & "'"
140       Set tb = New Recordset
150       RecOpenClient 0, tb, sql

160       If tb.EOF Then Exit Sub

170       Block2 = Mid$(tb!block & "", 74)

180       If Len(Block2) < 80 Then Exit Sub

190       Y = 1
200       For n = 2 To 142 Step 35

210           HbData(Y).PeakTimeSecs = Val(Mid$(Block2, n, 1)) * 60 + Val(Mid$(Block2, n + 2, 2))
220           Select Case Mid$(Block2, n + 4, 1)
              Case "S": HbData(Y).Analyte = "s-A1c": Ssecs = HbData(Y).PeakTimeSecs
230           Case "F": HbData(Y).Analyte = "HbF": Fsecs = HbData(Y).PeakTimeSecs
240           Case "#": HbData(Y).Analyte = "L-A1c"
250           Case Else: HbData(Y).Analyte = ""
260           End Select
270           HbData(Y).Area = Format$(Val(Mid$(Block2, n + 5, 6)))
280           HbData(Y).AreaRatio = Format$(Val(Mid$(Block2, n + 12, 4)), "0.0")

290           Y = Y + 1
300           HbData(Y).PeakTimeSecs = Val(Mid$(Block2, n + 17, 1)) * 60 + Val(Mid$(Block2, n + 19, 2))
310           Select Case Mid$(Block2, n + 21, 1)
              Case "S": HbData(Y).Analyte = "s-A1c": Ssecs = HbData(Y).PeakTimeSecs
320           Case "F": HbData(Y).Analyte = "HbF": Fsecs = HbData(Y).PeakTimeSecs
330           Case "#": HbData(Y).Analyte = "L-A1c"
340           Case Else: HbData(Y).Analyte = ""
350           End Select
360           HbData(Y).Area = Format$(Val(Mid$(Block2, n + 22, 6)))
370           HbData(Y).AreaRatio = Format$(Val(Mid$(Block2, n + 29, 4)), "0.0")

380           Y = Y + 1

390       Next

400       For n = 1 To 10
410           If HbData(n).PeakTimeSecs <> 0 Then
420               grdPoint.AddItem "P" & Format$(n) & vbTab & _
                                   HbData(n).Analyte & vbTab & _
                                   Format$(HbData(n).PeakTimeSecs) & vbTab & _
                                   HbData(n).Area & vbTab & _
                                   HbData(n).AreaRatio
430           End If
440       Next

450       If grdPoint.Rows > 2 Then
460           grdPoint.RemoveItem 1
470       End If

480       Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer



490       intEL = Erl
500       strES = Err.Description
510       LogError "frmViewFullDataHBA", "FillG", intEL, strES, sql


End Sub


Private Sub bcancel_Click()

10        Unload Me

End Sub


Private Sub Form_Activate()

10        On Error GoTo Form_Activate_Error

20        If Activated Then
30            Exit Sub
40        End If
50        Activated = True

60        DrawGraph

70        Exit Sub

Form_Activate_Error:

          Dim strES As String
          Dim intEL As Integer



80        intEL = Erl
90        strES = Err.Description
100       LogError "frmViewFullDataHBA", "Form_Activate", intEL, strES


End Sub


Public Property Let SampleID(ByVal sNewValue As String)

10        On Error GoTo SampleID_Error

20        mSampleID = sNewValue

30        Exit Property

SampleID_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmViewFullDataHBA", "SampleID", intEL, strES


End Property

Public Property Let HbA1c(ByVal sNewValue As String)

10        On Error GoTo HbA1c_Error

20        mhba1c = sNewValue

30        Exit Property

HbA1c_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmViewFullDataHBA", "HbA1c", intEL, strES


End Property

Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        Activated = False

30        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmViewFullDataHBA", "Form_Load", intEL, strES


End Sub

Private Sub Form_Unload(Cancel As Integer)

10        On Error GoTo Form_Unload_Error

20        Activated = False

30        Exit Sub

Form_Unload_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmViewFullDataHBA", "Form_Unload", intEL, strES


End Sub

