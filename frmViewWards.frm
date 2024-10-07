VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmViewWards 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Ward Look Up"
   ClientHeight    =   6720
   ClientLeft      =   180
   ClientTop       =   420
   ClientWidth     =   14070
   Icon            =   "frmViewWards.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   14070
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox optDept 
      Caption         =   "Urine"
      Height          =   255
      Index           =   8
      Left            =   6975
      TabIndex        =   14
      Top             =   630
      Width           =   825
   End
   Begin VB.CheckBox optDept 
      Caption         =   "Faeces"
      Height          =   255
      Index           =   7
      Left            =   6975
      TabIndex        =   13
      Top             =   360
      Width           =   825
   End
   Begin VB.CheckBox optDept 
      Caption         =   "Imm"
      Height          =   255
      Index           =   6
      Left            =   6975
      TabIndex        =   12
      Top             =   90
      Width           =   825
   End
   Begin VB.CheckBox optDept 
      Caption         =   "End"
      Height          =   255
      Index           =   5
      Left            =   5670
      TabIndex        =   11
      Top             =   585
      Width           =   825
   End
   Begin VB.CheckBox optDept 
      Alignment       =   1  'Right Justify
      Caption         =   "All"
      Height          =   255
      Index           =   4
      Left            =   5040
      TabIndex        =   10
      Top             =   90
      Width           =   555
   End
   Begin VB.CheckBox optDept 
      Caption         =   "Log In/Out"
      Height          =   285
      Index           =   3
      Left            =   5670
      TabIndex        =   6
      Top             =   90
      Width           =   1185
   End
   Begin VB.CheckBox optDept 
      Caption         =   "Coag"
      Height          =   255
      Index           =   2
      Left            =   5670
      TabIndex        =   9
      Top             =   360
      Width           =   825
   End
   Begin VB.CheckBox optDept 
      Alignment       =   1  'Right Justify
      Caption         =   "Bio"
      Height          =   255
      Index           =   1
      Left            =   5025
      TabIndex        =   8
      Top             =   585
      Width           =   570
   End
   Begin VB.CheckBox optDept 
      Alignment       =   1  'Right Justify
      Caption         =   "Haem"
      Height          =   285
      Index           =   0
      Left            =   4830
      TabIndex        =   7
      Top             =   315
      Width           =   765
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   5565
      Left            =   210
      TabIndex        =   4
      Top             =   990
      Width           =   13740
      _ExtentX        =   24236
      _ExtentY        =   9816
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      FormatString    =   $"frmViewWards.frx":030A
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
      Height          =   750
      Left            =   8160
      Picture         =   "frmViewWards.frx":03D3
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   90
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      Caption         =   "Between Dates"
      Height          =   945
      Left            =   225
      TabIndex        =   0
      Top             =   0
      Width           =   4440
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   675
         Left            =   3150
         Picture         =   "frmViewWards.frx":06DD
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   180
         Width           =   1020
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   315
         Left            =   1620
         TabIndex        =   1
         Top             =   300
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         Format          =   59047937
         CurrentDate     =   37679
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   315
         Left            =   150
         TabIndex        =   2
         Top             =   300
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         Format          =   59047937
         CurrentDate     =   37679
      End
   End
End
Attribute VB_Name = "frmViewWards"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Activated As Boolean

Private Sub FillG()

          Dim sql As String
          Dim tb As New Recordset
          Dim s As String

10        On Error GoTo FillG_Error

20        ClearFGrid g

30        If optDept(0) <> 1 And optDept(1) <> 1 And optDept(2) <> 1 And optDept(3) <> 1 And optDept(4) <> 1 Then
40            FixG g
50            Exit Sub
60        End If

70        sql = "SELECT * from ViewedReports WHERE " & _
                "DateTime between '" & Format$(dtFrom, "dd/mmm/yyyy") & "' " & _
                "and '" & Format$(dtTo, "dd/mmm/yyyy") & " 23:59:59' " & _
                "and ("

80        If optDept(4) Then
90            sql = Mid(sql, 1, Len(sql) - 5)
100       Else
110           If optDept(0) Or optDept(1) Or optDept(2) Then
120               sql = sql & "Discipline = 'A' or "
130           End If
140           If optDept(0) Then    'Haem
150               sql = sql & "Discipline = 'F' or " & _
                        "Discipline = 'G' or " & _
                        "Discipline = 'H' or " & _
                        "Discipline = 'R' or "
160           End If
170           If optDept(1) Then    'Bio
180               sql = sql & "Discipline = 'B' or " & _
                        "Discipline = 'D' or "
190           End If
200           If optDept(2) Then    'Coag
210               sql = sql & "Discipline = 'C' or " & _
                        "Discipline = 'E' or "
220           End If
230           If optDept(3) Then    'LogIn/Out
240               sql = sql & "Discipline = 'L' or " & _
                        "Discipline = 'M' or " & _
                        "Discipline = 'O' or " & _
                        "Discipline = 'X' or "
250           End If
260           If optDept(5) Then    'End
270               sql = sql & "Discipline = 'U' or " & _
                        "Discipline = 'V' or "
280           End If
290           If optDept(6) Then    'Imm
300               sql = sql & " Discipline = 'I' or " & _
                        "Discipline = 'J' or "
310           End If
320           If optDept(7) Then    'Faeces
330               sql = sql & " Discipline = 'Q' or"
340           End If
350           If optDept(8) Then    'Urine
360               sql = sql & " Discipline = 'P' or "
370           End If
380           If InStr(Right(sql, 3), "or") > 0 Then sql = Left$(sql, Len(sql) - 3) & ") "
390       End If

400       sql = sql & "order by DateTime desc"
410       Set tb = New Recordset
420       RecOpenClient 0, tb, sql
430       Do While Not tb.EOF
440           s = tb!Chart & vbTab & _
                  Format$(tb!Datetime, "dd/mm/yy hh:mm:ss") & vbTab & _
                  tb!SampleID & vbTab & tb!Name & "" & vbTab & tb!Dob & "" & vbTab & _
                  tb!Viewer & vbTab
450           Select Case tb!Discipline & ""
              Case "A": s = s & "Results Overview"
460           Case "B": s = s & "Biochemistry Result"
470           Case "C": s = s & "Coagulation Result"
480           Case "D": s = s & "Biochemistry History"
490           Case "E": s = s & "Coagulation History"
500           Case "F": s = s & "Haematology History"
510           Case "G": s = s & "Haematology Graphs"
520           Case "H": s = s & "Cumulative Haematology"
530           Case "I": s = s & "Immunology History"
540           Case "J": s = s & "Immunology Graphs"
550           Case "L": s = s & "Log On"
560           Case "M": s = s & "Manual Log Off"
570           Case "O": s = s & "Auto Log Off"
580           Case "R": s = s & "Haematology Result"
590           Case "S": s = s & "Blood Gas Result"
600           Case "T": s = s & "Blood Gas History"
610           Case "X": s = s & "Close Program"

                  'I Imm Result
                  'J Imm History

620           End Select
630           g.AddItem s
640           tb.MoveNext
650       Loop

660       FixG g

670       Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

680       intEL = Erl
690       strES = Err.Description
700       LogError "frmViewWards", "FillG", intEL, strES, sql

End Sub

Private Sub bcancel_Click()

10        Unload Me

End Sub


Private Sub cmdRefresh_Click()

10        On Error GoTo cmdRefresh_Click_Error

20        FillG

30        Exit Sub

cmdRefresh_Click_Error:

          Dim strES As String
          Dim intEL As Integer


40        intEL = Erl
50        strES = Err.Description
60        LogError "frmViewWards", "cmdRefresh_Click", intEL, strES


End Sub

Private Sub Form_Activate()

10        On Error GoTo Form_Activate_Error

20        Activated = True

30        Exit Sub

Form_Activate_Error:

          Dim strES As String
          Dim intEL As Integer


40        intEL = Erl
50        strES = Err.Description
60        LogError "frmViewWards", "Form_Activate", intEL, strES


End Sub

Private Sub Form_Load()

          Dim Checks As String
          Dim n As Long

10        On Error GoTo Form_Load_Error

20        Activated = False

          'If SysWard(0) = False Then
          '  iMsg "Not In Use"
          '  cmdRefresh.Enabled = False
          '  dtTo.Enabled = False
          '  dtFrom.Enabled = False
          '  For n = 0 To 3
          '    optDept(n).Enabled = False
          '  Next
          '  Exit Sub
          'End If

30        Checks = GetSetting("NetAcquire", "StartUp", "Check", "0000")
40        For n = 0 To 3
50            optDept(n) = IIf(Mid$(Checks, n + 1, 1) = "1", 1, 0)
60        Next

70        dtTo = Format$(Now, "dd/mm/yyyy")
80        dtFrom = Format$(Now, "dd/mm/yyyy")

90        FillG

100       Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer


110       intEL = Erl
120       strES = Err.Description
130       LogError "frmViewWards", "Form_Load", intEL, strES


End Sub


Private Sub Form_Unload(Cancel As Integer)

          Dim Checks As String
          Dim n As Long

10        On Error GoTo Form_Unload_Error

20        Activated = False

30        Checks = ""
40        For n = 0 To 3
50            Checks = Checks & IIf(optDept(n), "1", "0")
60        Next
70        SaveSetting "NetAcquire", "StartUp", "Check", Checks

80        Activated = False

90        Exit Sub

Form_Unload_Error:

          Dim strES As String
          Dim intEL As Integer


100       intEL = Erl
110       strES = Err.Description
120       LogError "frmViewWards", "Form_Unload", intEL, strES


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
90        End If

100       Exit Sub

g_Click_Error:

          Dim strES As String
          Dim intEL As Integer


110       intEL = Erl
120       strES = Err.Description
130       LogError "frmViewWards", "g_Click", intEL, strES


End Sub


Private Sub optDept_Click(Index As Integer)

10        On Error GoTo optDept_Click_Error

20        If Not Activated Then Exit Sub

30        FillG

40        Exit Sub

optDept_Click_Error:

          Dim strES As String
          Dim intEL As Integer


50        intEL = Erl
60        strES = Err.Description
70        LogError "frmViewWards", "optDept_Click", intEL, strES


End Sub


