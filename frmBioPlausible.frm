VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmBioPlausible 
   Caption         =   "NetAcquire - Biochemistry Plausible Ranges"
   ClientHeight    =   7155
   ClientLeft      =   1260
   ClientTop       =   1020
   ClientWidth     =   5415
   Icon            =   "frmBioPlausible.frx":0000
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
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   675
      Left            =   4185
      Picture         =   "frmBioPlausible.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6345
      Width           =   1035
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   675
      Left            =   4185
      Picture         =   "frmBioPlausible.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5490
      Width           =   1035
   End
   Begin MSFlexGridLib.MSFlexGrid grdPla 
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
Attribute VB_Name = "frmBioPlausible"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

10        Unload Me

End Sub

Private Sub cmdPrint_Click()

          Dim Y As Long

10        On Error GoTo cmdPrint_Click_Error

20        Printer.Print
30        Printer.Font.Name = "Courier New"
40        Printer.Font.Size = 12

50        Printer.Print "List of Biochemistry Plausible Ranges."
60        Printer.Print

70        For Y = 0 To grdPla.Rows - 1
80            grdPla.Row = Y
90            grdPla.Col = 0
100           Printer.Print grdPla; Tab(20);
110           grdPla.Col = 1
120           Printer.Print grdPla; Tab(30);
130           grdPla.Col = 2
140           Printer.Print grdPla
150       Next

160       Printer.EndDoc

170       Exit Sub

cmdPrint_Click_Error:

          Dim strES As String
          Dim intEL As Integer

180       intEL = Erl
190       strES = Err.Description
200       LogError "frmBioPlausible", "cmdPrint_Click", intEL, strES

End Sub

Private Sub FillG()

          Dim tb As New Recordset
          Dim sql As String



10        On Error GoTo FillG_Error

20        ClearFGrid grdPla

30        sql = "SELECT distinct ShortName, PrintPriority, " & _
                "PlausibleLow, PlausibleHigh " & _
                "from BioTestDefinitions " & _
                "Order by PrintPriority"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        Do While Not tb.EOF
70            grdPla.AddItem tb!ShortName & vbTab & _
                             Format$(tb!PlausibleLow & "") & vbTab & _
                             Format$(tb!PlausibleHigh & "")
80            tb.MoveNext
90        Loop

100       FixG grdPla




110       Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

120       intEL = Erl
130       strES = Err.Description
140       LogError "frmBioPlausible", "FillG", intEL, strES


End Sub

Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        grdPla.Font.Bold = True

30        FillG

40        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmBioPlausible", "Form_Load", intEL, strES


End Sub

Private Sub grdPla_Click()

          Static SortOrder As Boolean
          Dim temp As String


10        On Error GoTo grdPla_Click_Error

20        If grdPla.MouseRow = 0 Then
30            If SortOrder Then
40                grdPla.Sort = flexSortGenericAscending
50            Else
60                grdPla.Sort = flexSortGenericDescending
70            End If
80            SortOrder = Not SortOrder
90            Exit Sub
100       End If
110       temp = ""
120       If grdPla.Col = 0 Then
130           Exit Sub
140       ElseIf grdPla.Col = 1 Then
150           grdPla.Enabled = False
160           temp = grdPla
170           grdPla = iBOX(grdPla.TextMatrix(grdPla.Row, 0) & " Plausible Low?", , grdPla)
180           If grdPla = "" Then grdPla = temp
190           SaveG
200           grdPla.Enabled = True
210       Else
220           grdPla.Enabled = False
230           temp = grdPla
240           grdPla = iBOX(grdPla.TextMatrix(grdPla.Row, 0) & " Plausible High?", , grdPla)
250           If grdPla = "" Then grdPla = temp
260           SaveG
270           grdPla.Enabled = True
280       End If




290       Exit Sub

grdPla_Click_Error:

          Dim strES As String
          Dim intEL As Integer

300       intEL = Erl
310       strES = Err.Description
320       LogError "frmBioPlausible", "grdPla_Click", intEL, strES


End Sub

Private Sub SaveG()

          Dim sql As String



10        On Error GoTo SaveG_Error

20        sql = "UPDATE BioTestDefinitions " & _
                "Set PlausibleLow = " & Val(grdPla.TextMatrix(grdPla.Row, 1)) & ", " & _
                "PlausibleHigh = " & Val(grdPla.TextMatrix(grdPla.Row, 2)) & " " & _
                "WHERE ShortName = '" & grdPla.TextMatrix(grdPla.Row, 0) & "'"
30        Cnxn(0).Execute sql



40        Exit Sub

SaveG_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmBioPlausible", "SaveG", intEL, strES, sql


End Sub

