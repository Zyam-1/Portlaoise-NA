VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmClinicians 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Clinician List"
   ClientHeight    =   7935
   ClientLeft      =   510
   ClientTop       =   1170
   ClientWidth     =   9285
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
   Icon            =   "frmClinicians.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7935
   ScaleWidth      =   9285
   Begin VB.CommandButton cmdExcel 
      Caption         =   "Export to Excel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7785
      Picture         =   "frmClinicians.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4080
      Width           =   1380
   End
   Begin VB.CommandButton cmdMoveUp 
      Enabled         =   0   'False
      Height          =   555
      Left            =   7650
      Picture         =   "frmClinicians.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5520
      Width           =   465
   End
   Begin VB.CommandButton cmdMoveDown 
      Enabled         =   0   'False
      Height          =   525
      Left            =   7650
      Picture         =   "frmClinicians.frx":0A56
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6090
      Width           =   465
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Save"
      Height          =   705
      Left            =   7800
      Picture         =   "frmClinicians.frx":0E98
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2430
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Frame Frame1 
      Caption         =   "Add New Clinician"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1875
      Left            =   150
      TabIndex        =   12
      Top             =   180
      Width           =   8970
      Begin VB.ComboBox cmbHospital 
         Height          =   315
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   360
         Width           =   3585
      End
      Begin VB.TextBox txtSurname 
         Height          =   285
         Left            =   3810
         MaxLength       =   30
         TabIndex        =   4
         Top             =   1050
         Width           =   2625
      End
      Begin VB.TextBox txtForeName 
         Height          =   285
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   3
         Top             =   1050
         Width           =   1875
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Default         =   -1  'True
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   7515
         Picture         =   "frmClinicians.frx":11A2
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   675
         Width           =   1155
      End
      Begin VB.ComboBox cmbWard 
         Height          =   315
         Left            =   9000
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1665
         Visible         =   0   'False
         Width           =   2625
      End
      Begin VB.TextBox txtCode 
         Height          =   285
         Left            =   900
         MaxLength       =   12
         TabIndex        =   0
         Top             =   360
         Width           =   1545
      End
      Begin VB.TextBox txtTitle 
         Height          =   285
         Left            =   930
         MaxLength       =   30
         TabIndex        =   2
         Top             =   1050
         Width           =   975
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Hospital"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2610
         TabIndex        =   21
         Top             =   405
         Width           =   570
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Surname"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3810
         TabIndex        =   19
         Top             =   870
         Width           =   630
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "ForeName"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1950
         TabIndex        =   18
         Top             =   870
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Title"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   930
         TabIndex        =   17
         Top             =   870
         Width           =   300
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Default Ward"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   9675
         TabIndex        =   15
         Top             =   1575
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   450
         TabIndex        =   14
         Top             =   390
         Width           =   375
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Clinician"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   300
         TabIndex        =   13
         Top             =   1110
         Width           =   585
      End
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
      Height          =   735
      Left            =   7785
      Picture         =   "frmClinicians.frx":14AC
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3240
      Width           =   1380
   End
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
      Height          =   705
      Left            =   7785
      Picture         =   "frmClinicians.frx":17B6
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7155
      Width           =   1425
   End
   Begin MSFlexGridLib.MSFlexGrid grdClin 
      Height          =   5625
      Left            =   150
      TabIndex        =   11
      Top             =   2070
      Width           =   7485
      _ExtentX        =   13203
      _ExtentY        =   9922
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
      ScrollBars      =   2
      AllowUserResizing=   1
      FormatString    =   $"frmClinicians.frx":1AC0
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   225
      Left            =   7785
      TabIndex        =   16
      Top             =   2115
      Visible         =   0   'False
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      Height          =   345
      Left            =   7823
      TabIndex        =   23
      Top             =   4860
      Visible         =   0   'False
      Width           =   1305
   End
End
Attribute VB_Name = "frmClinicians"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Activated As Boolean

Private Sub cmbHospital_Click()

10        On Error GoTo cmbHospital_Click_Error

20        FillWards
30        FillG

40        Exit Sub

cmbHospital_Click_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmClinicians", "cmbHospital_Click", intEL, strES


End Sub

Private Sub cmbWard_Change()

10        On Error GoTo cmbWard_Change_Error

20        If Trim$(txtCode) = "" Or Trim$(txtSurname) = "" Then
30            cmdAdd.Enabled = False
40        Else
50            cmdAdd.Enabled = True
60        End If

70        Exit Sub

cmbWard_Change_Error:

          Dim strES As String
          Dim intEL As Integer

80        intEL = Erl
90        strES = Err.Description
100       LogError "frmClinicians", "cmbWard_Change", intEL, strES


End Sub

Private Sub cmdadd_Click()

          Dim ClinFullName As String
          Dim n As Long



10        On Error GoTo cmdadd_Click_Error

20        txtCode = UCase$(Trim$(txtCode))
30        If txtCode = "" Then
40            iMsg "Enter Code.", vbCritical
50            Exit Sub
60        End If

70        If Len(txtCode) <> 3 Then
80            iMsg "Code Length muxt be 3 Chars.", vbCritical
90            Exit Sub
100       End If

110       txtSurname = Trim$(txtSurname)
120       If txtSurname = "" Then
130           iMsg "Enter Clinicians Surname.", vbCritical
140           Exit Sub
150       End If


160       For n = 1 To grdClin.Rows - 1
170           If grdClin.TextMatrix(n, 1) = txtCode Then
180               iMsg "Code already used"
190               Exit Sub
200           End If
210       Next


220       ClinFullName = Trim$(txtTitle) & " " & Trim$(txtForeName) & " " & Trim$(txtSurname)

230       grdClin.AddItem "Yes" & vbTab & _
                          txtCode & vbTab & _
                          txtTitle & vbTab & _
                          txtForeName & vbTab & _
                          txtSurname & vbTab & _
                          ClinFullName & vbTab & _
                          cmbWard

240       txtCode = ""
250       txtTitle = ""
260       txtForeName = ""
270       txtSurname = ""
280       txtCode.Locked = False
290       cmbWard.ListIndex = -1

300       If grdClin.Rows > 2 And grdClin.TextMatrix(1, 0) = "" Then grdClin.RemoveItem 1

310       cmdSave.Visible = True

320       txtCode.SetFocus




330       Exit Sub

cmdadd_Click_Error:

          Dim strES As String
          Dim intEL As Integer

340       intEL = Erl
350       strES = Err.Description
360       LogError "frmClinicians", "cmdAdd_Click", intEL, strES


End Sub

Private Sub cmdCancel_Click()

10        Unload Me

End Sub

Private Sub cmdExcel_Click()
          Dim strTitle As String
10        strTitle = "List of Clinicians" & vbCr
20        ExportFlexGrid grdClin, Me, strTitle
End Sub

Private Sub cmdMoveDown_Click()

          Dim n As Long
          Dim s As String
          Dim X As Long



10        On Error GoTo cmdMoveDown_Click_Error

20        If grdClin.Row = grdClin.Rows - 1 Then Exit Sub
30        n = grdClin.Row

40        s = ""
50        For X = 0 To grdClin.Cols - 1
60            s = s & grdClin.TextMatrix(n, X) & vbTab
70        Next
80        s = Left$(s, Len(s) - 1)

90        grdClin.RemoveItem n
100       If n < grdClin.Rows Then
110           grdClin.AddItem s, n + 1
120           grdClin.Row = n + 1
130       Else
140           grdClin.AddItem s
150           grdClin.Row = grdClin.Rows - 1
160       End If

170       For X = 0 To grdClin.Cols - 1
180           grdClin.Col = X
190           grdClin.CellBackColor = vbYellow
200       Next

210       cmdSave.Visible = True



220       Exit Sub

cmdMoveDown_Click_Error:

          Dim strES As String
          Dim intEL As Integer

230       intEL = Erl
240       strES = Err.Description
250       LogError "frmClinicians", "cmdMoveDown_Click", intEL, strES


End Sub

Private Sub cmdMoveUp_Click()

          Dim n As Long
          Dim s As String
          Dim X As Long


10        On Error GoTo cmdMoveUp_Click_Error

20        If grdClin.Row = 1 Then Exit Sub

30        n = grdClin.Row

40        s = ""
50        For X = 0 To grdClin.Cols - 1
60            s = s & grdClin.TextMatrix(n, X) & vbTab
70        Next
80        s = Left$(s, Len(s) - 1)

90        grdClin.RemoveItem n
100       grdClin.AddItem s, n - 1

110       grdClin.Row = n - 1
120       For X = 0 To grdClin.Cols - 1
130           grdClin.Col = X
140           grdClin.CellBackColor = vbYellow
150       Next

160       cmdSave.Visible = True



170       Exit Sub

cmdMoveUp_Click_Error:

          Dim strES As String
          Dim intEL As Integer

180       intEL = Erl
190       strES = Err.Description
200       LogError "frmClinicians", "cmdMoveUp_Click", intEL, strES


End Sub

Private Sub cmdPrint_Click()

          Dim Y As Long

10        On Error GoTo cmdPrint_Click_Error

20        Printer.Print
30        Printer.Font.Name = "Courier New"
40        Printer.Font.Size = 12

50        Printer.Print "List of Clinicians."
60        Printer.Print

70        For Y = 0 To grdClin.Rows - 1
80            grdClin.Row = Y
90            grdClin.Col = 0    'Inuse
100           Printer.Print grdClin; Tab(8);
110           grdClin.Col = 1    'Code
120           Printer.Print grdClin; Tab(16);
130           grdClin.Col = 5    'Clinician
140           Printer.Print Left$(grdClin, 25); Tab(43);
150           grdClin.Col = 6    'Ward
160           Printer.Print grdClin
170       Next

180       Printer.EndDoc

190       Exit Sub

cmdPrint_Click_Error:

          Dim strES As String
          Dim intEL As Integer

200       intEL = Erl
210       strES = Err.Description
220       LogError "frmClinicians", "cmdPrint_Click", intEL, strES

End Sub

Private Sub cmdSave_Click()

          Dim HospCode As String
          Dim Y As Long
          Dim tb As New Recordset
          Dim sql As String



10        On Error GoTo cmdSave_Click_Error

20        HospCode = ListCodeFor("HO", cmbHospital)

30        pb.Max = grdClin.Rows - 1
40        pb.Visible = True
50        cmdSave.Caption = "SavingrdClin..."



60        For Y = 1 To grdClin.Rows - 1
70            pb = Y
80            sql = "SELECT * from Clinicians WHERE " & _
                    "Code = '" & grdClin.TextMatrix(Y, 1) & "' " & _
                    "and HospitalCode = '" & HospCode & "'"
90            Set tb = New Recordset
100           RecOpenClient 0, tb, sql
110           If tb.EOF Then
120               tb.AddNew
130           End If
140           tb!Code = UCase$(grdClin.TextMatrix(Y, 1))
150           tb!Text = initial2upper(grdClin.TextMatrix(Y, 5))
160           tb!HospitalCode = HospCode
170           If grdClin.TextMatrix(Y, 0) = "Yes" Then tb!InUse = 1 Else tb!InUse = 0
180           tb!Ward = grdClin.TextMatrix(Y, 6)
190           tb!Title = initial2upper(grdClin.TextMatrix(Y, 2))
200           tb!ForeName = initial2upper(grdClin.TextMatrix(Y, 3))
210           tb!SurName = initial2upper(grdClin.TextMatrix(Y, 4))
220           tb!ListOrder = Y
230           tb.Update
240       Next

250       pb.Visible = False
260       cmdSave.Visible = False
270       cmdSave.Caption = "Save"



280       Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

290       intEL = Erl
300       strES = Err.Description
310       LogError "frmClinicians", "cmdsave_Click", intEL, strES


End Sub

Private Sub FillG()

          Dim tb As New Recordset
          Dim sql As String
          Dim s As String
          Dim Hosp As String


10        On Error GoTo FillG_Error

20        ClearFGrid grdClin

30        Hosp = ListCodeFor("HO", cmbHospital)


40        sql = "SELECT * from Clinicians WHERE " & _
                "HospitalCode = '" & Hosp & "' ORDER BY LISTORDER"
50        Set tb = New Recordset
60        RecOpenServer 0, tb, sql

70        Do While Not tb.EOF
80            If tb!InUse = 1 Then s = "Yes" Else s = "No"
90            s = s & vbTab & _
                  tb!Code & vbTab & _
                  tb!Title & vbTab & _
                  tb!ForeName & vbTab & _
                  tb!SurName & vbTab & _
                  tb!Text
100           grdClin.AddItem s
110           tb.MoveNext
120       Loop

130       FixG grdClin




140       Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

150       intEL = Erl
160       strES = Err.Description
170       LogError "frmClinicians", "FillG", intEL, strES


End Sub

Sub FillWards()

          Dim HospCode As String
          Dim sql As String
          Dim tb As New Recordset



10        On Error GoTo FillWards_Error

20        cmbWard.Clear

30        HospCode = ListCodeFor("HO", cmbHospital)


40        sql = "SELECT * from wards WHERE hospitalcode = '" & HospCode & "'"
50        Set tb = New Recordset
60        RecOpenServer 0, tb, sql
70        Do While Not tb.EOF
80            cmbWard.AddItem Trim(tb!Text)
90            tb.MoveNext
100       Loop

110       cmbWard.ListIndex = -1




120       Exit Sub

FillWards_Error:

          Dim strES As String
          Dim intEL As Integer

130       intEL = Erl
140       strES = Err.Description
150       LogError "frmClinicians", "FillWards", intEL, strES


End Sub

Private Sub Form_Activate()

10        On Error GoTo Form_Activate_Error

20        If Activated Then
30            Exit Sub
40        End If

50        Activated = True

60        FillG

70        Exit Sub

Form_Activate_Error:

          Dim strES As String
          Dim intEL As Integer

80        intEL = Erl
90        strES = Err.Description
100       LogError "frmClinicians", "Form_Activate", intEL, strES


End Sub

Private Sub Form_Load()
          Dim tb As New Recordset
          Dim sql As String
          Dim n As Long


10        On Error GoTo Form_Load_Error

20        sql = "SELECT * from lists WHERE listtype = 'HO'"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql

50        Do While Not tb.EOF
60            cmbHospital.AddItem Trim(tb!Text)
70            tb.MoveNext
80        Loop

90        For n = 0 To cmbHospital.ListCount
100           If UCase(cmbHospital.List(n)) = UCase(HospName(0)) Then
110               cmbHospital.ListIndex = n
120           End If
130       Next



140       FillWards



150       Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

160       intEL = Erl
170       strES = Err.Description
180       LogError "frmClinicians", "Form_Load", intEL, strES, sql


End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

10        On Error GoTo Form_QueryUnload_Error

20        If cmdSave.Visible Then
30            If iMsg("Cancel without saving?", vbQuestion + vbYesNo) = vbNo Then
40                Cancel = True
50            End If
60        End If

70        Exit Sub

Form_QueryUnload_Error:

          Dim strES As String
          Dim intEL As Integer

80        intEL = Erl
90        strES = Err.Description
100       LogError "frmClinicians", "Form_QueryUnload", intEL, strES


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
60        LogError "frmClinicians", "Form_Unload", intEL, strES


End Sub

Private Sub grdClin_Click()

          Static SortOrder As Boolean
          Dim X As Long
          Dim Y As Long
          Dim ySave As Long


10        On Error GoTo grdClin_Click_Error

20        ySave = grdClin.Row

30        If grdClin.MouseRow = 0 Then
40            If SortOrder Then
50                grdClin.Sort = flexSortGenericAscending
60            Else
70                grdClin.Sort = flexSortGenericDescending
80            End If
90            SortOrder = Not SortOrder
100           cmdMoveUp.Enabled = False
110           cmdMoveDown.Enabled = False
120           cmdSave.Visible = True
130           Exit Sub
140       End If

150       If grdClin.Col = 0 Then
160           grdClin = IIf(grdClin = "Yes", "No", "Yes")
170           cmdSave.Visible = True
180           Exit Sub
190       End If

200       If grdClin.Col = 1 Then
210           grdClin.Enabled = False
220           If iMsg("Edit this line?", vbQuestion + vbYesNo) = vbYes Then
230               txtCode = grdClin.TextMatrix(grdClin.Row, 1)
240               txtCode.Locked = True
250               txtTitle = grdClin.TextMatrix(grdClin.Row, 2)
260               txtForeName = grdClin.TextMatrix(grdClin.Row, 3)
270               txtSurname = grdClin.TextMatrix(grdClin.Row, 4)
280               grdClin.RemoveItem grdClin.Row
290               cmdSave.Visible = True
300           End If
310           grdClin.Enabled = True
320           Exit Sub
330       End If

340       grdClin.Visible = False
350       grdClin.Col = 0
360       For Y = 1 To grdClin.Rows - 1
370           grdClin.Row = Y
380           If grdClin.CellBackColor = vbYellow Then
390               For X = 0 To grdClin.Cols - 1
400                   grdClin.Col = X
410                   grdClin.CellBackColor = 0
420               Next
430               Exit For
440           End If
450       Next
460       grdClin.Row = ySave
470       grdClin.Visible = True

480       For X = 0 To grdClin.Cols - 1
490           grdClin.Col = X
500           grdClin.CellBackColor = vbYellow
510       Next

520       cmdMoveUp.Enabled = True
530       cmdMoveDown.Enabled = True




540       Exit Sub

grdClin_Click_Error:

          Dim strES As String
          Dim intEL As Integer

550       intEL = Erl
560       strES = Err.Description
570       LogError "frmClinicians", "grdClin_Click", intEL, strES


End Sub

Private Sub grdClin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo grdClin_MouseMove_Error

20        If grdClin.MouseRow = 0 Then
30            grdClin.ToolTipText = ""
40        ElseIf grdClin.MouseCol = 0 Then
50            grdClin.ToolTipText = "Click to Toggle"
60        ElseIf grdClin.MouseCol = 1 Then
70            grdClin.ToolTipText = "Click to Edit"
80        Else
90            grdClin.ToolTipText = "Click to Move"
100       End If

110       Exit Sub

grdClin_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer

120       intEL = Erl
130       strES = Err.Description
140       LogError "frmClinicians", "grdClin_MouseMove", intEL, strES


End Sub

Private Sub txtCode_Change()

10        On Error GoTo txtCode_Change_Error

20        If Trim$(txtCode) = "" Or Trim$(txtSurname) = "" Then
30            cmdAdd.Enabled = False
40        Else
50            cmdAdd.Enabled = True
60        End If

70        Exit Sub

txtCode_Change_Error:

          Dim strES As String
          Dim intEL As Integer

80        intEL = Erl
90        strES = Err.Description
100       LogError "frmClinicians", "txtCode_Change", intEL, strES


End Sub

Private Sub txtForeName_Change()

10        On Error GoTo txtForeName_Change_Error

20        If Trim$(txtCode) = "" Or Trim$(txtSurname) = "" Then
30            cmdAdd.Enabled = False
40        Else
50            cmdAdd.Enabled = True
60        End If

70        Exit Sub

txtForeName_Change_Error:

          Dim strES As String
          Dim intEL As Integer

80        intEL = Erl
90        strES = Err.Description
100       LogError "frmClinicians", "txtForeName_Change", intEL, strES


End Sub

Private Sub txtForename_LostFocus()

10        On Error GoTo txtForename_LostFocus_Error

20        txtForeName = initial2upper(txtForeName)

30        Exit Sub

txtForename_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmClinicians", "txtForename_LostFocus", intEL, strES


End Sub

Private Sub txtSurname_Change()
10        On Error GoTo txtSurname_Change_Error

20        If Trim$(txtCode) = "" Or Trim$(txtSurname) = "" Then
30            cmdAdd.Enabled = False
40        Else
50            cmdAdd.Enabled = True
60        End If

70        Exit Sub

txtSurname_Change_Error:

          Dim strES As String
          Dim intEL As Integer

80        intEL = Erl
90        strES = Err.Description
100       LogError "frmClinicians", "txtSurname_Change", intEL, strES

End Sub

Private Sub txtSurname_LostFocus()

10        On Error GoTo txtSurname_LostFocus_Error

20        txtSurname = initial2upper(txtSurname)

30        Exit Sub

txtSurname_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmClinicians", "txtSurname_LostFocus", intEL, strES


End Sub

Private Sub txtTitle_Change()

10        On Error GoTo txtTitle_Change_Error

20        If Trim$(txtCode) = "" Or Trim$(txtSurname) = "" Then
30            cmdAdd.Enabled = False
40        Else
50            cmdAdd.Enabled = True
60        End If

70        Exit Sub

txtTitle_Change_Error:

          Dim strES As String
          Dim intEL As Integer

80        intEL = Erl
90        strES = Err.Description
100       LogError "frmClinicians", "txtTitle_Change", intEL, strES


End Sub

Private Sub txtTitle_LostFocus()

10        On Error GoTo txtTitle_LostFocus_Error

20        txtTitle = initial2upper(txtTitle)

30        Exit Sub

txtTitle_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmClinicians", "txtTitle_LostFocus", intEL, strES


End Sub
