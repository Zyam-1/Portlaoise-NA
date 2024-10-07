VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmSemenHistory 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Semen Analysis History"
   ClientHeight    =   3900
   ClientLeft      =   915
   ClientTop       =   4440
   ClientWidth     =   12675
   Icon            =   "frmSemenHistory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   12675
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton bCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   720
      Left            =   11580
      Picture         =   "frmSemenHistory.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2970
      Width           =   930
   End
   Begin MSFlexGridLib.MSFlexGrid grdHistory 
      Height          =   3315
      Left            =   60
      TabIndex        =   0
      Top             =   390
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   5847
      _Version        =   393216
      Cols            =   10
      RowHeightMin    =   600
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      BackColorSel    =   -2147483624
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      AllowUserResizing=   1
      FormatString    =   $"frmSemenHistory.frx":0614
   End
   Begin VB.Label lblChart 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6510
      TabIndex        =   4
      Top             =   60
      Width           =   1515
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Chart"
      Height          =   195
      Left            =   6090
      TabIndex        =   3
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblName 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   750
      TabIndex        =   2
      Top             =   60
      Width           =   3045
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      Height          =   195
      Left            =   300
      TabIndex        =   1
      Top             =   120
      Width           =   420
   End
End
Attribute VB_Name = "frmSemenHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub FillG()

          Dim sql As String
          Dim tb As New Recordset
          Dim s As String

10        On Error GoTo FillG_Error

20        With grdHistory
30            .Rows = 2
40            .AddItem ""
50            .RemoveItem 1
60        End With

70        sql = "SELECT D.PatName, D.SampleDate, S.* " & _
                "from Demographics as D, SemenResults as S WHERE " & _
                "D.PatName = '" & lblName & "' " & _
                "and D.SampleID = S.SampleID " & _
                "order by D.SampleID desc"
80        Set tb = New Recordset
90        RecOpenServer 0, tb, sql
100       Do While Not tb.EOF
110           If Trim$(lblName) = "" Then
120               lblName = tb!PatName & ""
130           End If
140           s = Trim$(tb!SampleID & "") - SysOptSemenOffset(0) & vbTab & _
                  Format(tb!SampleDate, "dd/mm/yy hh:mm") & vbTab & _
                  Trim$(tb!SemenCount & "") & vbTab & _
                  Trim$(tb!Volume & " ml") & vbTab & _
                  Trim$(tb!Consistency & "") & vbTab
150           If Trim$(tb!MotilityPro & "") <> "" Then
160               s = s & Trim$(tb!MotilityPro & "%")
170           End If
180           s = s & vbTab
190           If Trim$(tb!MotilityNonPro & "") <> "" Then
200               s = s & Trim$(tb!MotilityNonPro & "%")
210           End If
220           s = s & vbTab
230           If Trim$(tb!MotilityNonMotile & "") <> "" Then
240               s = s & Trim$(tb!MotilityNonMotile & "%")
250           End If
260           s = s & vbTab
270           If Trim$(tb!Motility & "") <> "" Then
280               s = s & Trim$(tb!Motility & "%")
290           End If
300           s = s & vbTab
310           s = s & LoadSemenMorphology(tb!SampleID)
320           grdHistory.AddItem s
330           tb.MoveNext
340       Loop

350       With grdHistory
360           If .Rows > 2 Then
370               .RemoveItem 1
380           End If
390       End With

400       Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

410       intEL = Erl
420       strES = Err.Description
430       LogError "frmSemenHistory", "FillG", intEL, strES, sql

End Sub

Private Sub bcancel_Click()

10        Unload Me

End Sub

Private Sub Form_Activate()

10        On Error GoTo Form_Activate_Error

20        grdHistory.TextMatrix(0, 5) = "Motile Progressive"
30        grdHistory.TextMatrix(0, 6) = "Motile Non Progressive"

40        FillG

50        Exit Sub

Form_Activate_Error:

          Dim strES As String
          Dim intEL As Integer

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmSemenHistory", "Form_Activate", intEL, strES

End Sub
Private Function LoadSemenMorphology(ByVal SampleIDWithOffset As Double) As String

          Dim sql As String
          Dim tb As Recordset
          Dim s As String

10        On Error GoTo LoadSemenMorphology_Error

20        LoadSemenMorphology = ""

30        sql = "SELECT Result FROM GenericResults WHERE " & _
                "SampleID = " & SampleIDWithOffset & " " & _
                "AND TestName = 'SemenMorphResult'"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        If Not tb.EOF Then
70            s = tb!Result & " "
80            sql = "SELECT * FROM GenericResults WHERE " & _
                    "SampleID = " & SampleIDWithOffset & " " & _
                    "AND TestName = 'SemenMorphDescription'"
90            Set tb = New Recordset
100           RecOpenServer 0, tb, sql
110           If Not tb.EOF Then
120               s = s & tb!Result & ""
130           End If
140           LoadSemenMorphology = s
150       End If

160       Exit Function

LoadSemenMorphology_Error:

          Dim strES As String
          Dim intEL As Integer

170       intEL = Erl
180       strES = Err.Description
190       LogError "frmSemenHistory", "LoadSemenMorphology", intEL, strES, sql

End Function


