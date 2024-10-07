VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmEndPlausible 
   Caption         =   "NetAcquire - Endocrinology Plausible Ranges"
   ClientHeight    =   7155
   ClientLeft      =   1260
   ClientTop       =   1020
   ClientWidth     =   5415
   Icon            =   "frmEndPlausible.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   5415
   Begin MSComctlLib.ProgressBar pb 
      Height          =   225
      Left            =   90
      TabIndex        =   3
      Top             =   0
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton bcancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   675
      Left            =   4185
      Picture         =   "frmEndPlausible.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6345
      Width           =   1035
   End
   Begin VB.CommandButton bPrint 
      Caption         =   "&Print"
      Height          =   765
      Left            =   4185
      Picture         =   "frmEndPlausible.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5355
      Width           =   1035
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   6855
      Left            =   90
      TabIndex        =   2
      Top             =   210
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   12091
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   2
      ScrollBars      =   2
      FormatString    =   "<Parameter               |^Low       |<High      "
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
End
Attribute VB_Name = "frmEndPlausible"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub FillG()

          Dim sql As String
          Dim tb As New Recordset
          Dim Found As Boolean
          Dim n As Long


10        On Error GoTo FillG_Error

20        ClearFGrid g

30        sql = "SELECT * from ImmTestDefinitions " & _
                "Order by PrintPriority"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql

60        Do While Not tb.EOF
70            With tb
80                Found = False
90                For n = 1 To g.Rows - 1
100                   If g.TextMatrix(n, 0) = !ShortName Then
110                       Found = True
120                       Exit For
130                   End If
140               Next
150               If Not Found Then
160                   g.AddItem !ShortName & vbTab & _
                                Format$(!PlausibleLow) & vbTab & _
                                Format$(!PlausibleHigh)
170               End If
180               .MoveNext
190           End With
200       Loop

210       FixG g




220       Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer



230       intEL = Erl
240       strES = Err.Description
250       LogError "frmEndPlausible", "FillG", intEL, strES, sql


End Sub
Private Sub SaveG()

          Dim sql As String
          Dim Y As Long


10        On Error GoTo SaveG_Error

20        Y = g.Row

30        sql = "UPDATE ImmTestDefinitions " & _
                "Set PlausibleLow = " & Val(g.TextMatrix(Y, 1)) & ", " & _
                "PlausibleHigh = " & Val(g.TextMatrix(Y, 2)) & " " & _
                "WHERE ShortName = '" & g.TextMatrix(Y, 0) & "'"
40        Cnxn(0).Execute sql



50        Exit Sub

SaveG_Error:

          Dim strES As String
          Dim intEL As Integer



60        intEL = Erl
70        strES = Err.Description
80        LogError "frmEndPlausible", "SaveG", intEL, strES, sql


End Sub



Private Sub bcancel_Click()

10        Unload Me

End Sub


Private Sub bprint_Click()

          Dim Y As Long

10        On Error GoTo bprint_Click_Error

20        Printer.Print
30        Printer.Font.Name = "Courier New"
40        Printer.Font.Size = 12

50        Printer.Print "List of Biochemistry Plausible Ranges."
60        Printer.Print

70        For Y = 0 To g.Rows - 1
80            g.Row = Y
90            g.Col = 0
100           Printer.Print g; Tab(20);
110           g.Col = 1
120           Printer.Print g; Tab(30);
130           g.Col = 2
140           Printer.Print g
150       Next

160       Printer.EndDoc



170       Exit Sub

bprint_Click_Error:

          Dim strES As String
          Dim intEL As Integer



180       intEL = Erl
190       strES = Err.Description
200       LogError "frmEndPlausible", "bPrint_Click", intEL, strES


End Sub


Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        g.Font.Bold = True

30        FillG

40        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer



50        intEL = Erl
60        strES = Err.Description
70        LogError "frmEndPlausible", "Form_Load", intEL, strES


End Sub


Private Sub g_Click()

          Static SortOrder As Boolean

10        On Error GoTo g_Click_Error

20        If g.MouseRow = 0 Then
30            If SortOrder Then
40                g.Sort = flexSortGenericAscending
50            Else
60                g.Sort = flexSortGenericDescending
70            End If
80            SortOrder = Not SortOrder
90            Exit Sub
100       End If

110       If g.Col = 0 Then
120           Exit Sub
130       ElseIf g.Col = 1 Then
140           g.Enabled = False
150           g = iBOX(g.TextMatrix(g.Row, 0) & " Plausible Low?", , g)
160           SaveG
170           g.Enabled = True
180       Else
190           g.Enabled = False
200           g = iBOX(g.TextMatrix(g.Row, 0) & " Plausible High?", , g)
210           SaveG
220           g.Enabled = True
230       End If

240       Exit Sub

g_Click_Error:

          Dim strES As String
          Dim intEL As Integer



250       intEL = Erl
260       strES = Err.Description
270       LogError "frmEndPlausible", "g_Click", intEL, strES


End Sub


