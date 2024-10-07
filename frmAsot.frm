VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAsot 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Add Haematology Result"
   ClientHeight    =   7065
   ClientLeft      =   465
   ClientTop       =   495
   ClientWidth     =   10470
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "frmAsot.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "frmAsot.frx":030A
   ScaleHeight     =   7065
   ScaleWidth      =   10470
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   9495
      Picture         =   "frmAsot.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6075
      Width           =   885
   End
   Begin VB.Frame Frame1 
      Caption         =   "Show"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   3255
      Begin VB.OptionButton optView 
         Caption         =   "&All"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   135
         TabIndex        =   7
         Top             =   210
         Width           =   675
      End
      Begin VB.OptionButton optView 
         Caption         =   "&Incomplete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   930
         TabIndex        =   6
         Top             =   210
         Value           =   -1  'True
         Width           =   1245
      End
      Begin VB.OptionButton optView 
         Caption         =   "&Ordered"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   2190
         TabIndex        =   5
         Top             =   210
         Width           =   1035
      End
   End
   Begin MSComCtl2.DTPicker dtDate 
      Height          =   315
      Left            =   3510
      TabIndex        =   3
      Top             =   180
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      _Version        =   393216
      Format          =   220069889
      CurrentDate     =   37082
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   9495
      Picture         =   "frmAsot.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4140
      Width           =   885
   End
   Begin MSFlexGridLib.MSFlexGrid grdHaem 
      Height          =   6135
      Left            =   270
      TabIndex        =   1
      Top             =   765
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   10821
      _Version        =   393216
      Cols            =   10
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      GridLines       =   3
      GridLinesFixed  =   3
      FormatString    =   "<Specimen  |<Chart         |<Name                      |<ESR   |<Retic |<IM   |<Asot  |<Sickledex  |<Malaria  |<RF  "
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   9495
      Picture         =   "frmAsot.frx":0C28
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5130
      Width           =   885
   End
End
Attribute VB_Name = "frmAsot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Activated As Boolean

Private Sub cmdCancel_Click()

10        Unload Me

End Sub

Private Sub cmdPrint_Click()

          Dim Num As Long



10        On Error GoTo cmdPrint_Click_Error

20        If grdHaem.Rows = 2 And grdHaem.TextMatrix(1, 0) = "" Then Exit Sub

30        For Num = 0 To grdHaem.Rows - 1
40            Printer.Print grdHaem.TextMatrix(Num, 0);
50            Printer.Print Tab(10); grdHaem.TextMatrix(Num, 1);    'chart
60            Printer.Print Tab(18); Left$(grdHaem.TextMatrix(Num, 2), 25);    'name
70            Printer.Print Tab(44); grdHaem.TextMatrix(Num, 3);    'esr
80            Printer.Print Tab(50); grdHaem.TextMatrix(Num, 4);    'retic
90            Printer.Print Tab(68); grdHaem.TextMatrix(Num, 5)    'im
100       Next

110       Printer.EndDoc




120       Exit Sub

cmdPrint_Click_Error:

          Dim strES As String
          Dim intEL As Integer

130       intEL = Erl
140       strES = Err.Description
150       LogError "frmAsot", "cmdPrint_Click", intEL, strES


End Sub

Private Sub cmdSave_Click()

          Dim tb As New Recordset
          Dim sql As String
          Dim Num As Long

10        On Error GoTo cmdSave_Click_Error

20        If grdHaem.Rows = 2 And grdHaem.TextMatrix(1, 0) = "" Then Exit Sub

30        For Num = 1 To grdHaem.Rows - 1
40            sql = "SELECT * from HaemResults WHERE " & _
                    "SampleID = '" & grdHaem.TextMatrix(Num, 0) & "'"
50            Set tb = New Recordset
60            RecOpenServer 0, tb, sql
70            If Not tb.EOF Then
80                tb!ESR = Left$(Trim$(grdHaem.TextMatrix(Num, 3)), 5)
90                tb!retics = Left$(Trim$(grdHaem.TextMatrix(Num, 4)), 5)
100               tb!MonoSpot = Left$(Trim$(grdHaem.TextMatrix(Num, 5)), 1)
110               tb!tASOt = Left$(Trim$(grdHaem.TextMatrix(Num, 6)), 1)
120               tb.Update
130           End If
140       Next

150       Unload Me

160       Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

170       intEL = Erl
180       strES = Err.Description
190       LogError "frmAsot", "cmdsave_Click", intEL, strES

End Sub

Private Sub dtDate_CloseUp()

10        On Error GoTo dtDate_CloseUp_Error

20        FillGrid

30        Exit Sub

dtDate_CloseUp_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmAsot", "dtDate_CloseUp", intEL, strES


End Sub

Private Sub FillGrid()

      Dim sn As New Recordset
      Dim sql As String
      Dim Str As String

10    On Error GoTo FillGrid_Error

20    ClearFGrid grdHaem

30    grdHaem.Visible = False

40    sql = "SELECT Demographics.Chart, Demographics.PatName, haemresults.cmonospot, " & _
            "HaemResults.ESR, haemresults.cesr, HaemResults.Reta, HaemResults.tasot, HaemResults.casot, "
50    sql = sql & "HaemResults.MonoSpot, haemresults.cretics,haemresults.Sickledex,haemresults.cSickledex,haemresults.Malaria,HaemResults.cMalaria,haemresults.ra,HaemResults.cra, haemresults.sampleid  from Demographics, HaemResults WHERE " & _
            "demographics.RunDate = '" & Format$(dtDate, "dd/mmm/yyyy") & "' " & _
            "and Demographics.SampleID = HaemResults.SampleID "
60    If optView(1) Then
70        sql = sql & "and ( (cesr = 1  and ((ESR is null) or esr = '?')) "
80        sql = sql & "or (cretics = 1 and ((Reta is null) or reta = '?')) " & _
                "or (cmonospot = 1 and ((MonoSpot is null) or monospot = '?')) "
90        sql = sql & " or  (cSickledex = 1  and ((Sickledex is null) or Sickledex = '?')) "
100       sql = sql & " or  (cMalaria = 1  and ((Malaria is null) or Malaria = '?')) "
110       sql = sql & " or  (cRa = 1  and ((ra is null) or ra = '?'))) "
120   ElseIf optView(2) Then
130       sql = sql & "and ( cesr = 1 " & _
                "or cretics = 1 " & _
                "or cmonospot = 1 " & _
                "or cSickledex = 1 " & _
                "or cMalaria = 1 " & _
                "Or Cra=1)"

140   End If
150   sql = sql & " order by Demographics.SampleID"

160   Set sn = New Recordset
170   RecOpenServer 0, sn, sql
180   Do While Not sn.EOF
190       Str = Trim$(sn!SampleID) & vbTab & Trim$(sn!Chart & "") & vbTab
200       Str = Str & sn!PatName & vbTab
210       If sn!cESR = True And sn!ESR <> "" Then
220           Str = Str & sn!ESR
230       ElseIf sn!cESR = True Then
240           Str = Str & "X"
250       End If
260       Str = Str & vbTab
270       If sn!cRetics = True And sn!reta <> "" Then
280           Str = Str & sn!reta
290       ElseIf sn!cRetics = True Then
300           Str = Str & "X"
310       End If
320       Str = Str & vbTab
330       If sn!cMonospot = True And sn!MonoSpot <> "" Then
340           Str = Str & sn!MonoSpot
350       ElseIf sn!cMonospot = True Then
360           Str = Str & "X"
370       End If
380       Str = Str & vbTab
390       If sn!cASot = True And sn!tASOt <> "" Then
400           Str = Str & sn!tASOt
410       ElseIf sn!cASot = True Then
420           Str = Str & "X"
430       End If
440       Str = Str & vbTab
450       If sn!csickledex = True And sn!sickledex <> "" Then
460           Str = Str & sn!sickledex
470       ElseIf sn!csickledex = True Then
480           Str = Str & "X"
490       End If
500       Str = Str & vbTab
510       If sn!cMalaria = True And sn!Malaria <> "" Then
520           Str = Str & sn!Malaria
530       ElseIf sn!cMalaria = True Then
540           Str = Str & "X"
550       End If
560       Str = Str & vbTab
570       If sn!cRA = True And sn!Ra <> "" Then
580           Str = Str & sn!Ra
590       ElseIf sn!cRA = True Then
600           Str = Str & "X"
610       End If
620       Str = Str & vbTab
630       grdHaem.AddItem Str
640       sn.MoveNext
650   Loop

660   FixG grdHaem

670   Exit Sub

FillGrid_Error:

      Dim strES As String
      Dim intEL As Integer

680   intEL = Erl
690   strES = Err.Description
700   LogError "frmAsot", "FillGrid", intEL, strES, sql

End Sub

Private Sub Form_Activate()

10        On Error GoTo Form_Activate_Error

20        If Not Activated Then
30            Activated = True
40            FillGrid
50        End If

60        Set_Font Me

70        Exit Sub

Form_Activate_Error:

          Dim strES As String
          Dim intEL As Integer

80        intEL = Erl
90        strES = Err.Description
100       LogError "frmAsot", "Form_Activate", intEL, strES


End Sub

Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        Activated = False

30        dtDate = Format$(Now, "dd/mm/yyyy")

40        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmAsot", "Form_Load", intEL, strES


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
60        LogError "frmAsot", "Form_Unload", intEL, strES


End Sub

Private Sub grdHaem_Click()

          Dim Str As String
          Dim sql As String
          Dim tb As Recordset


10        On Error GoTo grdHaem_Click_Error

20        If grdHaem.MouseRow = 0 Or grdHaem.Col < 3 Then Exit Sub

30        If grdHaem.TextMatrix(grdHaem.RowSel, 0) = "" Then Exit Sub

40        sql = "Select * from haemresults where sampleid = " & grdHaem.TextMatrix(grdHaem.RowSel, 0) & ""
50        Set tb = New Recordset
60        RecOpenServer 0, tb, sql
70        If tb.EOF Or tb!Valid = 1 Then
80            iMsg "Result Valid. Can not be changed!"
90            Exit Sub
100       End If

110       Str = iBOX("Enter " & grdHaem.TextMatrix(0, grdHaem.Col) & " for " & vbCrLf & grdHaem.TextMatrix(grdHaem.row, 2), , grdHaem.TextMatrix(grdHaem.row, grdHaem.Col))

120       Str = UCase(Trim(Str))

130       If grdHaem.Col = 5 Or grdHaem.Col = 6 Then    '
140           If Str <> "P" Or Str <> "N" Or Str <> "I" Then
150               iMsg "Incorrect Result. Resullt either P,N or I"
160               Str = ""
170           End If
180       ElseIf grdHaem.Col = 3 Or grdHaem.Col = 4 Then    '
190           If IsNumeric(Str) Or Str = "?" Then
200           Else
210               iMsg "Result must be numeric!"
220               Str = ""
230           End If
240       End If

250       grdHaem = Str




260       Exit Sub

grdHaem_Click_Error:

          Dim strES As String
          Dim intEL As Integer

270       intEL = Erl
280       strES = Err.Description
290       LogError "frmAsot", "grdHaem_Click", intEL, strES, sql


End Sub

Private Sub optView_Click(Index As Integer)

10        On Error GoTo optView_Click_Error

20        FillGrid

30        Exit Sub

optView_Click_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmAsot", "optView_Click", intEL, strES


End Sub
