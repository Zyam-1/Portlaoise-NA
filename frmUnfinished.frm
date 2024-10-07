VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmUnfinished 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Unfinished Samples"
   ClientHeight    =   8550
   ClientLeft      =   2115
   ClientTop       =   1530
   ClientWidth     =   8535
   Icon            =   "frmUnfinished.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   8535
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Department"
      Height          =   3075
      Left            =   6750
      TabIndex        =   5
      Top             =   1860
      Width           =   1545
      Begin VB.OptionButton optDept 
         Caption         =   "Red Sub"
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   15
         Top             =   2730
         Width           =   945
      End
      Begin VB.OptionButton optDept 
         Caption         =   "C && S"
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   14
         Top             =   2460
         Width           =   765
      End
      Begin VB.OptionButton optDept 
         Caption         =   "Ova/Parasites"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   13
         Top             =   2190
         Width           =   1365
      End
      Begin VB.OptionButton optDept 
         Caption         =   "CSF"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   12
         Top             =   1920
         Width           =   705
      End
      Begin VB.OptionButton optDept 
         Caption         =   "Rota/Adeno"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   11
         Top             =   1650
         Width           =   1215
      End
      Begin VB.OptionButton optDept 
         Caption         =   "Urine"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   10
         Top             =   1380
         Width           =   705
      End
      Begin VB.OptionButton optDept 
         Caption         =   "Faeces"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Top             =   1110
         Width           =   945
      End
      Begin VB.OptionButton optDept 
         Caption         =   "c. Diff"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   765
      End
      Begin VB.OptionButton optDept 
         Caption         =   "H. Pylori"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   570
         Width           =   945
      End
      Begin VB.OptionButton optDept 
         Caption         =   "All"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   300
         Value           =   -1  'True
         Width           =   555
      End
   End
   Begin MSComCtl2.DTPicker dt 
      Height          =   315
      Left            =   6810
      TabIndex        =   3
      Top             =   1410
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   556
      _Version        =   393216
      Format          =   16515073
      CurrentDate     =   39730
   End
   Begin VB.CommandButton bcancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   720
      Left            =   6720
      Picture         =   "frmUnfinished.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7620
      Width           =   1470
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   8265
      Left            =   300
      TabIndex        =   0
      Top             =   45
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   14579
      _Version        =   393216
      Cols            =   3
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
      FormatString    =   "<Sample Number    |<Outstanding                     |<Date                      "
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
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "View Samples dated later than"
      Height          =   465
      Left            =   6810
      TabIndex        =   4
      Top             =   930
      Width           =   1485
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   6600
      Picture         =   "frmUnfinished.frx":0614
      Top             =   30
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Click Heading to Sort"
      Height          =   435
      Left            =   7080
      TabIndex        =   2
      Top             =   60
      Width           =   1125
   End
End
Attribute VB_Name = "frmUnfinished"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private SortOrder As Boolean

Private Sub FillG(ByVal DeptCode As String, _
                  ByVal DeptNumber As Integer)

          Dim sql As String
          Dim tb As Recordset
          Dim Department(1 To 9) As String
          Dim s As String

10        Department(1) = "H. Pylori"
20        Department(2) = "c. Diff"
30        Department(3) = "Faeces"
40        Department(4) = "Urine"
50        Department(5) = "Rota/Adeno"
60        Department(6) = "CSF"
70        Department(7) = "Ova/Parasites"
80        Department(8) = "C & S"
90        Department(9) = "Red Sub"

100       sql = "SELECT P.SampleID, D.RunDate FROM PrintValidLog P, Demographics D WHERE " & _
                "P.Department = '" & DeptCode & "' " & _
                "AND P.Valid = 0 " & _
                "AND P.SampleID = D.SampleID " & _
                "AND D.RunDate >= '" & Format$(dt, "dd/MMM/yyyy") & "'"
110       Set tb = New Recordset
120       RecOpenServer 0, tb, sql
130       Do While Not tb.EOF
140           s = Format$(Val(tb!SampleID) - SysOptMicroOffset(0)) & vbTab & _
                  Department(DeptNumber) & vbTab & _
                  Format$(tb!Rundate, "dd/MM/yyyy")
150           g.AddItem s
160           tb.MoveNext
170       Loop

End Sub

Private Sub RefreshGrid()

          Dim Dept As String
          Dim n As Integer
          Dim OptionSelected As Integer

10        On Error GoTo RefreshGrid_Error

20        g.Rows = 2
30        g.AddItem ""
40        g.RemoveItem 1
50        g.Visible = False
60        Screen.MousePointer = vbHourglass

70        OptionSelected = 0
80        For n = 0 To 9
90            If optDept(n).Value Then
100               OptionSelected = n
110               Exit For
120           End If
130       Next

140       Dept = "YGFUACODR"

150       If OptionSelected = 0 Then
160           For n = 1 To Len(Dept)
170               FillG Mid$(Dept, n, 1), n
180           Next
190       Else
200           FillG Mid$(Dept, OptionSelected, 1), n
210       End If

220       g.Col = 0
230       g.Sort = flexSortGenericAscending
240       If g.Rows > 2 Then
250           g.RemoveItem 1
260       End If
270       g.Visible = True
280       Screen.MousePointer = vbNormal

290       Exit Sub

RefreshGrid_Error:

          Dim strES As String
          Dim intEL As Integer

300       intEL = Erl
310       strES = Err.Description
320       LogError "frmUnfinished", "RefreshGrid", intEL, strES
330       Screen.MousePointer = vbNormal

End Sub

Private Sub bcancel_Click()

10        Unload Me

End Sub


Private Sub dt_CloseUp()

10        RefreshGrid

End Sub


Private Sub Form_Activate()

10        RefreshGrid

End Sub

Private Sub Form_Load()

10        dt = Format$(Now - 3, "dd/MM/yyyy")

End Sub


Private Sub g_Click()


10        On Error GoTo g_Click_Error

20        If g.Rows = 2 Then Exit Sub

30        If g.MouseRow = 0 Then
40            If InStr(UCase$(g.TextMatrix(0, g.MouseCol)), "DATE") <> 0 Then
50                g.Sort = 9
60            ElseIf SortOrder Then
70                g.Sort = flexSortGenericAscending
80            Else
90                g.Sort = flexSortGenericDescending
100           End If
110           SortOrder = Not SortOrder
120           Exit Sub
130       End If

140       frmEditMicrobiologyNew.ForcedSID = g.TextMatrix(g.Row, 0)
150       frmEditMicrobiologyNew.Show 1

160       Exit Sub

g_Click_Error:

          Dim strES As String
          Dim intEL As Integer

170       intEL = Erl
180       strES = Err.Description
190       LogError "frmUnfinished", "g_Click", intEL, strES

End Sub
Private Sub g_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)

          Dim d1 As String
          Dim d2 As String

10        If Not IsDate(g.TextMatrix(Row1, g.Col)) Then
20            Cmp = 0
30            Exit Sub
40        End If

50        If Not IsDate(g.TextMatrix(Row2, g.Col)) Then
60            Cmp = 0
70            Exit Sub
80        End If

90        d1 = Format(g.TextMatrix(Row1, g.Col), "dd/MMM/yyyy HH:mm:ss")
100       d2 = Format(g.TextMatrix(Row2, g.Col), "dd/MMM/yyyy HH:mm:ss")

110       If SortOrder Then
120           Cmp = Sgn(DateDiff("s", d1, d2))
130       Else
140           Cmp = Sgn(DateDiff("s", d2, d1))
150       End If

End Sub


Private Sub optDept_Click(Index As Integer)

10        RefreshGrid

End Sub


