VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmFormOpt 
   Caption         =   "Form Tabbing Options for "
   ClientHeight    =   7560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9690
   Icon            =   "frmFormOpt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   9690
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Tool Tip Text"
      Height          =   555
      Left            =   4545
      TabIndex        =   18
      Top             =   4680
      Width           =   3705
      Begin VB.OptionButton optTip 
         Caption         =   "On"
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
         Index           =   1
         Left            =   1980
         TabIndex        =   20
         Top             =   225
         Width           =   1185
      End
      Begin VB.OptionButton optTip 
         Caption         =   "Off"
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
         Index           =   0
         Left            =   225
         TabIndex        =   19
         Top             =   225
         Width           =   1185
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Down"
      Height          =   735
      Left            =   8640
      Picture         =   "frmFormOpt.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3465
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "Up"
      Height          =   690
      Left            =   8595
      Picture         =   "frmFormOpt.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   495
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Exit"
      Height          =   780
      Left            =   8415
      Picture         =   "frmFormOpt.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6660
      Width           =   1095
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update Tab Order"
      Height          =   825
      Left            =   270
      Picture         =   "frmFormOpt.frx":0C28
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4275
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Caption         =   "Font Selection"
      Height          =   2130
      Left            =   225
      TabIndex        =   10
      Top             =   5310
      Width           =   4290
      Begin VB.TextBox txtTest 
         Height          =   1590
         Left            =   90
         TabIndex        =   13
         Top             =   315
         Width           =   2580
      End
      Begin VB.CommandButton cmdFont 
         Caption         =   "Font Selection"
         Height          =   915
         Left            =   2790
         Picture         =   "frmFormOpt.frx":0F32
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   225
         Width           =   1320
      End
      Begin VB.CommandButton cmdUpdFont 
         Caption         =   "Set Font"
         Height          =   825
         Left            =   2790
         Picture         =   "frmFormOpt.frx":123C
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1215
         Width           =   1320
      End
      Begin MSComDlg.CommonDialog dlgFonts 
         Left            =   135
         Top             =   1305
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Range Colors"
      Height          =   2130
      Left            =   4635
      TabIndex        =   3
      Top             =   5310
      Width           =   3615
      Begin MSComDlg.CommonDialog dlgColor 
         Left            =   45
         Top             =   1530
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdUpdateCol 
         Caption         =   "Update Color"
         Height          =   870
         Left            =   1620
         Picture         =   "frmFormOpt.frx":1546
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   855
         Width           =   1185
      End
      Begin VB.OptionButton optCol 
         Caption         =   "Back"
         CausesValidation=   0   'False
         Height          =   240
         Index           =   1
         Left            =   2385
         TabIndex        =   8
         Top             =   360
         Width           =   690
      End
      Begin VB.OptionButton optCol 
         Caption         =   "Fore"
         CausesValidation=   0   'False
         Height          =   240
         Index           =   0
         Left            =   1485
         TabIndex        =   7
         Top             =   360
         Value           =   -1  'True
         Width           =   690
      End
      Begin VB.Label lblHigh 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "High"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   330
         Left            =   90
         TabIndex        =   6
         Top             =   1125
         Width           =   1230
      End
      Begin VB.Label lblPlas 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         Caption         =   "Inplausible"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   90
         TabIndex        =   5
         Top             =   720
         Width           =   1230
      End
      Begin VB.Label lblLow 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "Low"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   330
         Left            =   90
         TabIndex        =   4
         Top             =   315
         Width           =   1230
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdForm 
      Height          =   3705
      Left            =   270
      TabIndex        =   2
      Top             =   495
      Width           =   8340
      _ExtentX        =   14711
      _ExtentY        =   6535
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      FormatString    =   $"frmFormOpt.frx":1850
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.OptionButton optForm 
      Caption         =   "Histology"
      Height          =   330
      Index           =   1
      Left            =   3870
      TabIndex        =   1
      Top             =   90
      Visible         =   0   'False
      Width           =   2445
   End
   Begin VB.OptionButton optForm 
      Caption         =   "General Chemistry"
      Height          =   330
      Index           =   0
      Left            =   1530
      TabIndex        =   0
      Top             =   90
      Visible         =   0   'False
      Width           =   2445
   End
End
Attribute VB_Name = "frmFormOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Ind As Integer


Private Sub cmdCancel_Click()
10        On Error GoTo cmdCancel_Click_Error

20        LoadUserOpts

30        Unload frmEditAll
40        Unload Me

50        Exit Sub

cmdCancel_Click_Error:

          Dim strES As String
          Dim intEL As Integer



60        intEL = Erl
70        strES = Err.Description
80        LogError "frmFormOpt", "cmdCancel_Click", intEL, strES


End Sub

Private Sub cmdFont_Click()


10        On Error GoTo cmdFont_Click_Error

20        dlgFonts.Flags = cdlCFScreenFonts
30        dlgFonts.FontName = txtTest.Font
40        dlgFonts.ShowFont
50        With txtTest.Font
60            .Name = dlgFonts.FontName
70        End With


80        Exit Sub

cmdFont_Click_Error:

          Dim strES As String
          Dim intEL As Integer



90        intEL = Erl
100       strES = Err.Description
110       LogError "frmFormOpt", "cmdFont_Click", intEL, strES


End Sub


Private Sub cmdUp_Click()


          Dim n As Long
          Dim s As String
          Dim X As Long

10        On Error GoTo cmdUp_Click_Error

20        If grdForm.Row = 1 Then Exit Sub

30        n = grdForm.Row

40        s = ""
50        For X = 0 To grdForm.Cols - 1
60            s = s & grdForm.TextMatrix(n, X) & vbTab
70        Next
80        s = Left$(s, Len(s) - 1)

90        grdForm.RemoveItem n
100       grdForm.AddItem s, n - 1

110       grdForm.Row = n - 1
120       For X = 0 To grdForm.Cols - 1
130           grdForm.Col = X
140           grdForm.CellBackColor = vbYellow
150       Next

160       grdForm.TopRow = grdForm.Row

170       Exit Sub

cmdUp_Click_Error:

          Dim strES As String
          Dim intEL As Integer



180       intEL = Erl
190       strES = Err.Description
200       LogError "frmFormOpt", "cmdUp_Click", intEL, strES


End Sub

Private Sub cmdUpdate_Click()

10        On Error GoTo cmdUpdate_Click_Error

20        UPDATE_Tab

30        optForm_Click Ind


40        Exit Sub

cmdUpdate_Click_Error:

          Dim strES As String
          Dim intEL As Integer



50        intEL = Erl
60        strES = Err.Description
70        LogError "frmFormOpt", "cmdUpdate_Click", intEL, strES


End Sub

Private Sub cmdUPDATECol_Click()
          Dim sql As String
          Dim tb As New Recordset



10        On Error GoTo cmdUPDATECol_Click_Error

20        sql = "SELECT * from options WHERE description = 'LOWFORE' and username = '" & Username & "'"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If tb.EOF Then tb.AddNew
60        tb!Username = Username
70        tb!Description = "LOWFORE"
80        tb!Contents = lblLow.ForeColor
90        tb.Update

100       sql = "SELECT * from options WHERE description = 'LOWBACK' and username = '" & Username & "'"
110       Set tb = New Recordset
120       RecOpenServer 0, tb, sql
130       If tb.EOF Then tb.AddNew
140       tb!Username = Username
150       tb!Description = "LOWBACK"
160       tb!Contents = lblLow.BackColor
170       tb.Update

180       sql = "SELECT * from options WHERE description = 'PLASFORE' and username = '" & Username & "'"
190       Set tb = New Recordset
200       RecOpenServer 0, tb, sql
210       If tb.EOF Then tb.AddNew
220       tb!Username = Username
230       tb!Description = "PLASFORE"
240       tb!Contents = lblPlas.ForeColor
250       tb.Update

260       sql = "SELECT * from options WHERE description = 'PLASBACK' and username = '" & Username & "'"
270       Set tb = New Recordset
280       RecOpenServer 0, tb, sql
290       If tb.EOF Then tb.AddNew
300       tb!Username = Username
310       tb!Description = "PLASBACK"
320       tb!Contents = lblPlas.BackColor
330       tb.Update

340       sql = "SELECT * from options WHERE description = 'HIGHFORE' and username = '" & Username & "'"
350       Set tb = New Recordset
360       RecOpenServer 0, tb, sql
370       If tb.EOF Then tb.AddNew
380       tb!Username = Username
390       tb!Description = "HIGHFORE"
400       tb!Contents = lblHigh.ForeColor
410       tb.Update

420       sql = "SELECT * from options WHERE description = 'HIGHBACK' and username = '" & Username & "'"
430       Set tb = New Recordset
440       RecOpenServer 0, tb, sql
450       If tb.EOF Then tb.AddNew
460       tb!Username = Username
470       tb!Description = "HIGHBACK"
480       tb!Contents = lblHigh.BackColor
490       tb.Update

500       LoadUserOpts

510       Exit Sub

cmdUPDATECol_Click_Error:

          Dim strES As String
          Dim intEL As Integer



520       intEL = Erl
530       strES = Err.Description
540       LogError "frmFormOpt", "cmdUPDATECol_Click", intEL, strES, sql


End Sub

Private Sub cmdUpdFont_Click()
          Dim sql As String
          Dim tb As New Recordset


10        On Error GoTo cmdUpdFont_Click_Error

20        sql = "SELECT * from options WHERE description = 'FONT' and username = '" & Username & "'"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If tb.EOF Then tb.AddNew
60        tb!Username = Username
70        tb!Description = "FONT"
80        tb!Contents = txtTest.Font.Name
90        tb.Update

100       LoadUserOpts

110       Set_Font Me

120       Exit Sub

cmdUpdFont_Click_Error:

          Dim strES As String
          Dim intEL As Integer



130       intEL = Erl
140       strES = Err.Description
150       LogError "frmFormOpt", "cmdUpdFont_Click", intEL, strES, sql


End Sub

Private Sub Command2_Click()
          Dim n As Long
          Dim s As String
          Dim X As Long

10        On Error GoTo Command2_Click_Error

20        If grdForm.Row = grdForm.Rows - 1 Then Exit Sub

30        n = grdForm.Row

40        s = ""
50        For X = 0 To grdForm.Cols - 1
60            s = s & grdForm.TextMatrix(n, X) & vbTab
70        Next
80        s = Left$(s, Len(s) - 1)

90        grdForm.RemoveItem n
100       grdForm.AddItem s, n + 1

110       grdForm.Row = n + 1
120       For X = 0 To grdForm.Cols - 1
130           grdForm.Col = X
140           grdForm.CellBackColor = vbYellow
150       Next

160       grdForm.TopRow = grdForm.Row

170       Exit Sub

Command2_Click_Error:

          Dim strES As String
          Dim intEL As Integer



180       intEL = Erl
190       strES = Err.Description
200       LogError "frmFormOpt", "Command2_Click", intEL, strES

End Sub

Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        Me.Caption = Me.Caption & Username

30        Set_Font Me
40        lblLow.ForeColor = SysOptLowFore(0)
50        lblLow.BackColor = SysOptLowBack(0)
60        lblPlas.ForeColor = SysOptPlasFore(0)
70        lblPlas.BackColor = SysOptPlasBack(0)
80        lblHigh.ForeColor = SysOptHighFore(0)
90        lblHigh.BackColor = SysOptHighBack(0)
100       If SysOptToolTip(0) = True Then
110           optTip(1).Value = True
120       Else
130           optTip(0).Value = True
140       End If

150       Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer



160       intEL = Erl
170       strES = Err.Description
180       LogError "frmFormOpt", "Form_Load", intEL, strES


End Sub

Private Sub grdForm_Click()
          Dim n As Long

10        On Error GoTo grdForm_Click_Error

20        If grdForm.ColSel = 3 Then
30            grdForm.TextMatrix(grdForm.RowSel, 3) = iBOX("Set TabIndex", , grdForm.TextMatrix(grdForm.RowSel, 3), False)
40        End If

50        For n = 0 To 4
60            grdForm.Col = n
70            grdForm.CellBackColor = vbRed
80        Next

90        Exit Sub

grdForm_Click_Error:

          Dim strES As String
          Dim intEL As Integer



100       intEL = Erl
110       strES = Err.Description
120       LogError "frmFormOpt", "grdForm_Click", intEL, strES


End Sub

Private Sub lblHigh_Click()

10        On Error GoTo lblHigh_Click_Error

20        If optCol(0) Then
30            dlgColor.Color = lblHigh.ForeColor
40        Else
50            dlgColor.Color = lblHigh.BackColor
60        End If

70        dlgColor.ShowColor

80        If optCol(0) Then
90            lblHigh.ForeColor = dlgColor.Color
100       Else
110           lblHigh.BackColor = dlgColor.Color
120       End If

130       Exit Sub

lblHigh_Click_Error:

          Dim strES As String
          Dim intEL As Integer



140       intEL = Erl
150       strES = Err.Description
160       LogError "frmFormOpt", "lblHigh_Click", intEL, strES


End Sub

Private Sub lblLow_Click()

10        On Error GoTo lblLow_Click_Error

20        If optCol(0) Then
30            dlgColor.Color = lblLow.ForeColor
40        Else
50            dlgColor.Color = lblLow.BackColor
60        End If

70        dlgColor.ShowColor

80        If optCol(0) Then
90            lblLow.ForeColor = dlgColor.Color
100       Else
110           lblLow.BackColor = dlgColor.Color
120       End If


130       Exit Sub

lblLow_Click_Error:

          Dim strES As String
          Dim intEL As Integer



140       intEL = Erl
150       strES = Err.Description
160       LogError "frmFormOpt", "lblLow_Click", intEL, strES


End Sub

Private Sub lblPlas_Click()

10        On Error GoTo lblPlas_Click_Error

20        If optCol(0) Then
30            dlgColor.Color = lblPlas.ForeColor
40        Else
50            dlgColor.Color = lblPlas.BackColor
60        End If

70        dlgColor.ShowColor

80        If optCol(0) Then
90            lblPlas.ForeColor = dlgColor.Color
100       Else
110           lblPlas.BackColor = dlgColor.Color
120       End If

130       Exit Sub

lblPlas_Click_Error:

          Dim strES As String
          Dim intEL As Integer



140       intEL = Erl
150       strES = Err.Description
160       LogError "frmFormOpt", "lblPlas_Click", intEL, strES

End Sub

Private Sub optForm_Click(Index As Integer)
          Dim tx As Control
          Dim tb As New Recordset
          Dim sql As String


10        On Error GoTo optForm_Click_Error

20        Ind = Index

30        ClearFGrid grdForm



40        If optForm(0) Then
50            sql = "SELECT * from options WHERE description like 'frmEditAll.%' and username = '" & Username & "'"
60            Set tb = New Recordset
70            RecOpenServer 0, tb, sql
80            If Not tb.EOF Then
90                Do While Not tb.EOF
100                   For Each tx In frmEditAll.Controls
110                       If tx.Name = Link(tb!Description, 2) Then
120                           grdForm.AddItem Link(tb!Description, 1) & vbTab & Link(tb!Description, 2) & vbTab & tx.Tag & vbTab & tb!Contents
130                       End If
140                   Next
150                   tb.MoveNext
160               Loop
170           Else
180               For Each tx In frmEditAll.Controls
190                   If TypeOf tx Is TextBox Then
200                       grdForm.AddItem "frmEditAll" & vbTab & tx.Name & vbTab & tx.Tag & vbTab & tx.TabIndex & vbTab & tx.Visible
210                   ElseIf TypeOf tx Is CommandButton Then
220                       grdForm.AddItem "frmEditAll" & vbTab & tx.Name & vbTab & tx.Tag & vbTab & tx.TabIndex & vbTab & tx.Visible
230                   End If
240               Next
250           End If
260       Else
270           For Each tx In frmEditHisto.Controls
280               If TypeOf tx Is TextBox Then
290                   grdForm.AddItem "frmEditHisto" & vbTab & tx.Name & vbTab & tx.Tag & vbTab & Val(tx.TabIndex) & vbTab & tx.Visible
300               ElseIf TypeOf tx Is CommandButton Then
310                   grdForm.AddItem "frmEditHisto" & vbTab & tx.Name & vbTab & tx.Tag & vbTab & Val(tx.TabIndex) & vbTab & tx.Visible
320               End If
330           Next
340       End If


350       Sort 3

360       FixG grdForm

370       Exit Sub

optForm_Click_Error:

          Dim strES As String
          Dim intEL As Integer



380       intEL = Erl
390       strES = Err.Description
400       LogError "frmFormOpt", "optForm_Click", intEL, strES, sql


End Sub

Private Sub Sort(ByVal Col As Long)

10        On Error GoTo Sort_Error

20        grdForm.Col = Col
30        grdForm.Sort = 3

40        Exit Sub

Sort_Error:

          Dim strES As String
          Dim intEL As Integer



50        intEL = Erl
60        strES = Err.Description
70        LogError "frmFormOpt", "Sort", intEL, strES


End Sub

Private Sub UPDATE_Tab()
          Dim sql As String
          Dim tb As New Recordset

          Dim n As Long



10        On Error GoTo UPDATE_Tab_Error

20        For n = 1 To grdForm.Rows - 1
30            sql = "DELETE from options WHERE description = '" & grdForm.TextMatrix(n, 0) & "." & grdForm.TextMatrix(n, 1) & "' and username = '" & Username & "'"
40            Cnxn(0).Execute sql
50            sql = "SELECT * from options WHERE description = '" & grdForm.TextMatrix(n, 0) & "." & grdForm.TextMatrix(n, 1) & "' and username = '" & Username & "'"
60            Set tb = New Recordset
70            RecOpenServer 0, tb, sql
80            If tb.EOF Then tb.AddNew
90            tb!Description = grdForm.TextMatrix(n, 0) & "." & grdForm.TextMatrix(n, 1)
100           tb!Contents = n - 1
110           tb!Username = Username
120           tb.Update
130       Next

140       sql = "DELETE from options WHERE description = 'SETFOC' and username = '" & Username & "'"
150       Cnxn(0).Execute sql

160       sql = "INSERT into options ( description, contents, username) values " & _
                "('SETFOC' , '" & grdForm.TextMatrix(1, 1) & "', '" & Username & "')"
170       Cnxn(0).Execute sql

180       Exit Sub

UPDATE_Tab_Error:

          Dim strES As String
          Dim intEL As Integer



190       intEL = Erl
200       strES = Err.Description
210       LogError "frmFormOpt", "UPDATE_Tab", intEL, strES, sql


End Sub


Private Function Link(ByVal Desc As String, ByVal Li As Long) As String
          Dim n As Long
          Dim s As String

10        On Error GoTo Link_Error

20        Desc = Trim(Desc)

30        If Li = 1 Then
40            For n = 1 To Len(Desc)
50                If Mid(Desc, n, 1) = "." Then Exit For
60                s = s & Mid(Desc, n, 1)
70            Next
80        ElseIf Li = 2 Then
90            s = Mid(Desc, 10, Len(Desc) - 9)
100       End If

110       Link = s

120       Exit Function

Link_Error:

          Dim strES As String
          Dim intEL As Integer



130       intEL = Erl
140       strES = Err.Description
150       LogError "frmFormOpt", "Link", intEL, strES

End Function

Private Sub optTip_Click(Index As Integer)
          Dim OnOff As Integer
          Dim sql As String

10        On Error GoTo optTip_Click_Error

20        If optTip(0).Value = True Then
30            OnOff = 0
40        Else
50            OnOff = 1
60        End If


70        sql = "DELETE from options WHERE description = 'TOOLTIP' and username = '" & Username & "'"
80        Cnxn(0).Execute sql

90        sql = "INSERT into options ( description, contents, username) values " & _
                "('TOOLTIP' , '" & OnOff & "', '" & Username & "')"
100       Cnxn(0).Execute sql



110       Exit Sub

optTip_Click_Error:

          Dim strES As String
          Dim intEL As Integer



120       intEL = Erl
130       strES = Err.Description
140       LogError "frmFormOpt", "optTip_Click", intEL, strES, sql


End Sub
