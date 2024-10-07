VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmBioSplitList 
   Caption         =   "NetAcquire - Biochemistry Splits"
   ClientHeight    =   7725
   ClientLeft      =   75
   ClientTop       =   375
   ClientWidth     =   10635
   Icon            =   "frmBioSplitList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   10635
   Begin VB.CommandButton cmdMove 
      Caption         =   "Remove Item"
      Enabled         =   0   'False
      Height          =   735
      Index           =   0
      Left            =   3090
      Picture         =   "frmBioSplitList.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3240
      Width           =   1245
   End
   Begin MSFlexGridLib.MSFlexGrid grdSplit 
      Height          =   6855
      Index           =   1
      Left            =   4440
      TabIndex        =   6
      Top             =   690
      Width           =   2925
      _ExtentX        =   5159
      _ExtentY        =   12091
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      FormatString    =   "^Type   |<Analyte Name                    "
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "Move to Split 2"
      Enabled         =   0   'False
      Height          =   735
      Index           =   2
      Left            =   3090
      Picture         =   "frmBioSplitList.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2370
      Width           =   1245
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "Move to Split 1"
      Enabled         =   0   'False
      Height          =   705
      Index           =   1
      Left            =   3090
      Picture         =   "frmBioSplitList.frx":0B8E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1530
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   735
      Left            =   3060
      Picture         =   "frmBioSplitList.frx":0FD0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6750
      Width           =   1245
   End
   Begin MSFlexGridLib.MSFlexGrid grdSplit 
      Height          =   6855
      Index           =   2
      Left            =   7470
      TabIndex        =   7
      Top             =   690
      Width           =   2925
      _ExtentX        =   5159
      _ExtentY        =   12091
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      FormatString    =   "^Type   |<Analyte Name                    "
   End
   Begin MSFlexGridLib.MSFlexGrid grdSplit 
      Height          =   7365
      Index           =   0
      Left            =   60
      TabIndex        =   8
      Top             =   180
      Width           =   2925
      _ExtentX        =   5159
      _ExtentY        =   12991
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      FormatString    =   "^Type   |<Analyte Name                    "
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Secondary"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8160
      TabIndex        =   5
      Top             =   330
      Width           =   1140
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Primary"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4950
      TabIndex        =   4
      Top             =   330
      Width           =   810
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Highlight item to be moved then click appropriate arrow."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3060
      TabIndex        =   3
      Top             =   180
      Width           =   1245
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   9300
      Picture         =   "frmBioSplitList.frx":12DA
      Top             =   210
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   5760
      Picture         =   "frmBioSplitList.frx":15E4
      Top             =   210
      Width           =   480
   End
End
Attribute VB_Name = "frmBioSplitList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Disp As String
Private intLastGridUsed As Long

Private Sub cmdCancel_Click()

10        Unload Me

End Sub

Private Sub cmdMove_Click(Index As Integer)

      'Index is 'Move To'

          Dim sql As String
          Dim intFromIndex As Integer
          Dim strAnalyte As String
          Dim strSampleType As String

10        On Error GoTo cmdMove_Click_Error

20        If Index = 0 Then
30            intFromIndex = intLastGridUsed
40        Else
50            intFromIndex = 0
60        End If

70        With grdSplit(intFromIndex)
80            strAnalyte = .TextMatrix(.Row, 1)
90            strSampleType = Left$(.TextMatrix(.Row, 0), 1)
100       End With

110       sql = "UPDATE " & Disp & "TestDefinitions " & _
                "Set Splitlist = " & Index & " " & _
                "WHERE LongName = '" & strAnalyte & "' " & _
                "and SampleType = '" & strSampleType & "'"
120       Cnxn(0).Execute sql

130       cmdMove(0).Enabled = False
140       cmdMove(1).Enabled = False
150       cmdMove(2).Enabled = False

160       FillGrids

170       Exit Sub

cmdMove_Click_Error:

          Dim strES As String
          Dim intEL As Integer

180       intEL = Erl
190       strES = Err.Description
200       LogError "frmBioSplitList", "cmdMove_Click", intEL, strES


End Sub

Private Sub FillGrids()

          Dim tb As New Recordset
          Dim sql As String
          Dim SampleType As String
          Dim intY As Long

10        On Error GoTo FillGrids_Error

20        For intY = 0 To 2
30            ClearFGrid grdSplit(intY)
40        Next

50        If Disp = "End" Then
60            Me.Caption = "NetAcquire - Endocrinology Splits"
70        ElseIf Disp = "Imm" Then
80            Me.Caption = "NetAcquire - Immunology Splits"
90        End If

100       sql = "SELECT distinct LongName, SampleType, SplitList " & _
                "from " & Disp & "TestDefinitions "

110       Set tb = New Recordset
120       RecOpenServer 0, tb, sql
130       Do While Not tb.EOF
140           SampleType = ListText("ST", tb!SampleType & "")    'IIf(tb!SampleType = "U", "Urine", "Serum")
150           If Not IsNull(tb!SplitList) Then
160               grdSplit(tb!SplitList).AddItem SampleType & vbTab & tb!LongName
170           Else
180               grdSplit(0).AddItem SampleType & vbTab & tb!LongName
190           End If
200           tb.MoveNext
210       Loop

220       For intY = 0 To 2
230           If grdSplit(intY).Rows > 2 Then
240               grdSplit(intY).RemoveItem 1
250           End If
260           grdSplit(intY).Visible = True
270       Next

280       Exit Sub

FillGrids_Error:

          Dim strES As String
          Dim intEL As Integer

290       intEL = Erl
300       strES = Err.Description
310       LogError "frmBioSplitList", "FillGrids", intEL, strES


End Sub

Private Sub Form_Load()


10        On Error GoTo Form_Load_Error

20        FillGrids

30        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmBioSplitList", "Form_Load", intEL, strES


End Sub

Private Sub grdSplit_Click(Index As Integer)

          Dim intY As Long
          Dim intRowSave As Long
          Dim intActive As Long

10        On Error GoTo grdSplit_Click_Error

20        If grdSplit(Index).MouseRow = 0 Then Exit Sub
30        If grdSplit(Index).TextMatrix(1, 0) = "" Then Exit Sub

40        intLastGridUsed = Index

50        intRowSave = grdSplit(Index).Row

60        For intActive = 0 To 2
70            If intActive <> Index Then
80                With grdSplit(intActive)
90                    For intY = 1 To .Rows - 1
100                       .Row = intY
110                       .Col = 0
120                       .CellBackColor = &H80000018
130                       .Col = 1
140                       .CellBackColor = &H80000018
150                   Next
160               End With
170           End If
180       Next

190       With grdSplit(Index)
200           For intY = 1 To .Rows - 1
210               .Row = intY
220               .Col = 0
230               .CellBackColor = &H80000018
240               .Col = 1
250               .CellBackColor = &H80000018
260           Next
270           .Row = intRowSave
280           .Col = 0
290           .CellBackColor = vbYellow
300           .Col = 1
310           .CellBackColor = vbYellow
320       End With

330       If Index = 0 Then
340           cmdMove(0).Enabled = False
350           cmdMove(1).Enabled = True
360           cmdMove(2).Enabled = True
370       Else
380           cmdMove(0).Enabled = True
390           cmdMove(1).Enabled = False
400           cmdMove(2).Enabled = False
410       End If

420       Exit Sub

grdSplit_Click_Error:

          Dim strES As String
          Dim intEL As Integer

430       intEL = Erl
440       strES = Err.Description
450       LogError "frmBioSplitList", "grdSplit_Click", intEL, strES


End Sub
