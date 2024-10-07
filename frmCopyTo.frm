VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCopyTo 
   Caption         =   "NetAcquire"
   ClientHeight    =   4890
   ClientLeft      =   75
   ClientTop       =   1560
   ClientWidth     =   9360
   Icon            =   "frmCopyTo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4890
   ScaleWidth      =   9360
   Begin VB.ComboBox cmbPrinter 
      Height          =   315
      Left            =   7140
      TabIndex        =   12
      Text            =   "cmbPrinter"
      Top             =   4920
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.CommandButton bCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Exit"
      Height          =   795
      Left            =   7860
      Picture         =   "frmCopyTo.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3930
      Width           =   1245
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   795
      Left            =   6510
      Picture         =   "frmCopyTo.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3930
      Width           =   1245
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   795
      Index           =   2
      Left            =   5340
      Picture         =   "frmCopyTo.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3660
      Width           =   825
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   795
      Index           =   1
      Left            =   2760
      Picture         =   "frmCopyTo.frx":0D60
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3660
      Width           =   825
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   795
      Index           =   0
      Left            =   360
      Picture         =   "frmCopyTo.frx":11A2
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3660
      Width           =   825
   End
   Begin VB.ComboBox cmbGP 
      Height          =   315
      Left            =   4110
      TabIndex        =   1
      Text            =   "cmbGP"
      Top             =   4860
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox cmbWard 
      Height          =   315
      Left            =   390
      TabIndex        =   2
      Text            =   "cmbWard"
      Top             =   4860
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox cmbClinician 
      Height          =   315
      Left            =   2250
      TabIndex        =   0
      Top             =   4860
      Visible         =   0   'False
      Width           =   1275
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   2385
      Left            =   240
      TabIndex        =   13
      Top             =   1200
      Width           =   8865
      _ExtentX        =   15637
      _ExtentY        =   4207
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      RowHeightMin    =   315
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      FormatString    =   $"frmCopyTo.frx":15E4
   End
   Begin VB.Label lblSampleID 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7620
      TabIndex        =   14
      Top             =   510
      Width           =   1485
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Sample ID"
      Height          =   195
      Left            =   7620
      TabIndex        =   10
      Top             =   300
      Width           =   735
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Send Copy To"
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   990
      Width           =   1020
   End
   Begin VB.Label lblOriginal 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Top             =   510
      Width           =   7155
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Send Original To"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   300
      Width           =   1185
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuWards 
         Caption         =   "&Ward List Length"
      End
      Begin VB.Menu mnuClinicians 
         Caption         =   "&Clinician List Length"
      End
      Begin VB.Menu mnuGPs 
         Caption         =   "&GP List Length"
      End
      Begin VB.Menu mnuNull 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmCopyTo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mEditScreen As Form

Private DataChanged As Boolean

Private Sub AdjustBlanks()

          Dim Y As Long
          Dim x As Long
          Dim IsBlank As Boolean

          'remove all blank lines
10        On Error GoTo AdjustBlanks_Error

20        For Y = g.Rows - 1 To 1 Step -1
30            IsBlank = True
40            For x = 0 To 2
50                If g.TextMatrix(Y, x) <> "" Then
60                    IsBlank = False
70                    Exit For
80                End If
90            Next
100           If IsBlank Then
110               If g.Rows > 2 Then
120                   g.RemoveItem Y
130               Else
140                   g.AddItem ""
150                   g.RemoveItem 1
160               End If
170           End If
180       Next

          'add a blank line to the bottom
190       Y = g.Rows - 1
200       IsBlank = True
210       For x = 0 To 2
220           If g.TextMatrix(Y, x) <> "" Then
230               IsBlank = False
240           End If
250       Next
260       If Not IsBlank Then
270           g.AddItem ""
280       End If

290       Exit Sub

AdjustBlanks_Error:

          Dim strES As String
          Dim intEL As Integer

300       intEL = Erl
310       strES = Err.Description
320       LogError "frmCopyTo", "AdjustBlanks", intEL, strES

End Sub

Private Sub bcancel_Click()

10        Unload Me

End Sub

Private Sub cmbClinician_Click()

10        On Error GoTo cmbClinician_Click_Error

20        DataChanged = True

30        g.SetFocus

40        Exit Sub

cmbClinician_Click_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmCopyTo", "cmbClinician_Click", intEL, strES

End Sub


Private Sub cmbClinician_LostFocus()

10        On Error GoTo cmbClinician_LostFocus_Error

20        cmbClinician = QueryKnown("Clin", cmbClinician, mEditScreen.cmbHospital)

30        If g.Enabled Then g.SetFocus

40        Exit Sub

cmbClinician_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmCopyTo", "cmbClinician_LostFocus", intEL, strES

End Sub

Private Sub cmbGP_LostFocus()

10        On Error GoTo cmbGP_LostFocus_Error

20        cmbGP = QueryKnown("GP", cmbGP, mEditScreen.cmbHospital)

30        If g.Enabled Then g.SetFocus

40        Exit Sub

cmbGP_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmCopyTo", "cmbGP_LostFocus", intEL, strES

End Sub

Private Sub cmbPrinter_Click()

10        On Error GoTo cmbPrinter_Click_Error

20        DataChanged = True

30        g.SetFocus

40        Exit Sub

cmbPrinter_Click_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmCopyTo", "cmbPrinter_Click", intEL, strES

End Sub


Private Sub cmbGP_Click()

10        On Error GoTo cmbGP_Click_Error

20        DataChanged = True

30        g.SetFocus

40        Exit Sub

cmbGP_Click_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmCopyTo", "cmbGP_Click", intEL, strES

End Sub


Private Sub cmbWard_Click()

10        On Error GoTo cmbWard_Click_Error

20        DataChanged = True

30        If g.Enabled Then g.SetFocus

40        Exit Sub

cmbWard_Click_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmCopyTo", "cmbWard_Click", intEL, strES

End Sub



Private Sub cmbWard_LostFocus()

          Dim Found As Boolean
          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo cmbWard_LostFocus_Error

20        If Trim$(cmbWard) = "" Then
30            cmbWard = "GP"
40            Exit Sub
50        End If

60        Found = False

70        sql = "SELECT * from wards WHERE (text = '" & AddTicks(cmbWard) & "' or code = '" & AddTicks(cmbWard) & "') and hospitalcode = '" & ListCodeFor("HO", mEditScreen.cmbHospital) & "'"
80        Set tb = New Recordset
90        RecOpenServer 0, tb, sql
100       If Not tb.EOF Then
110           cmbWard = Trim(tb!Text)
120           Found = True
130       End If

140       If Found = False Then
150           cmbWard = ""
160       End If
170       If g.Enabled = True Then g.SetFocus
180       Exit Sub

cmbWard_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

190       intEL = Erl
200       strES = Err.Description
210       LogError "frmCopyTo", "cmbWard_LostFocus", intEL, strES, sql

End Sub

Private Sub cmdClear_Click(Index As Integer)

          Dim Y As Long

10        On Error GoTo cmdClear_Click_Error

20        For Y = 1 To g.Rows - 1
30            g.TextMatrix(Y, Index) = ""
40        Next

50        Select Case Index
          Case 0: cmbWard.Visible = False
60        Case 1: cmbClinician.Visible = False
70        Case 2: cmbGP.Visible = False
80        End Select

90        AdjustBlanks

100       Exit Sub

cmdClear_Click_Error:

          Dim strES As String
          Dim intEL As Integer

110       intEL = Erl
120       strES = Err.Description
130       LogError "frmCopyTo", "cmdClear_Click", intEL, strES

End Sub

Private Sub cmdSave_Click()

          Dim sql As String
          Dim Y As Long

10        On Error GoTo cmdSave_Click_Error

20        sql = "Delete from SendCopyTo where " & _
                "SampleID = '" & lblSampleID & "'"
30        Cnxn(0).Execute sql

40        For Y = 1 To g.Rows - 2
50            If g.TextMatrix(Y, 4) = "Use Default" Then
60                g.TextMatrix(Y, 4) = ""
70            End If
80            sql = "Insert into SendCopyTo " & _
                    "(SampleID, Ward, Clinician, GP, Device, Destination) VALUES " & _
                    "('" & lblSampleID & "', " & _
                    " '" & AddTicks(g.TextMatrix(Y, 0)) & "', " & _
                    " '" & AddTicks(g.TextMatrix(Y, 1)) & "', " & _
                    " '" & AddTicks(g.TextMatrix(Y, 2)) & "', " & _
                    " '" & AddTicks(g.TextMatrix(Y, 3)) & "', " & _
                    " '" & AddTicks(g.TextMatrix(Y, 4)) & "')"

90            Cnxn(0).Execute sql
100       Next

110       Unload Me

120       Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

130       intEL = Erl
140       strES = Err.Description
150       LogError "frmCopyTo", "cmdsave_Click", intEL, strES, sql

End Sub

Private Sub Form_Activate()

          Dim sql As String
          Dim tb As Recordset
          Dim s As String

10        On Error GoTo Form_Activate_Error

20        g.Rows = 2
30        g.AddItem ""
40        g.RemoveItem 1

50        sql = "Select * from SendCopyTo where " & _
                "SampleID = '" & lblSampleID & "'"
60        Set tb = New Recordset
70        RecOpenServer 0, tb, sql
80        Do While Not tb.EOF
90            s = tb!Ward & vbTab & _
                  tb!Clinician & vbTab & _
                  tb!GP & vbTab & _
                  tb!Device & vbTab & _
                  tb!Destination & ""
100           g.AddItem s
110           tb.MoveNext
120       Loop

130       If g.Rows > 2 Then
140           g.RemoveItem 1
150       End If

160       Call ChangeComboHeight(cmbClinician, GetOptionSetting("ClinicianCCListLength", 8))
170       Call ChangeComboHeight(cmbGP, GetOptionSetting("GPCCListLength", 8))
180       Call ChangeComboHeight(cmbWard, GetOptionSetting("WardCCListLength", 8))
190       g.ColWidth(4) = 0


200       Exit Sub

Form_Activate_Error:

          Dim strES As String
          Dim intEL As Integer

210       intEL = Erl
220       strES = Err.Description
230       LogError "frmCopyTo", "Form_Activate", intEL, strES, sql

End Sub

Public Property Let EditScreen(ByVal f As Form)

10        On Error GoTo EditScreen_Error

20        Set mEditScreen = f

30        Exit Property

EditScreen_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmCopyTo", "EditScreen", intEL, strES

End Property
Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        FillGPsClinWard Me, mEditScreen.cmbHospital

30        cmbWard.AddItem "", 0
40        cmbClinician.AddItem "", 0
50        cmbGP.AddItem "", 0

60        FillPrinterList

70        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

80        intEL = Erl
90        strES = Err.Description
100       LogError "frmCopyTo", "Form_Load", intEL, strES

End Sub

Private Sub g_Click()

          Dim Found As Boolean
          Dim FaxNumber As String
          Dim WCG As String

10        On Error GoTo g_Click_Error

20        If g.MouseRow = 0 Then Exit Sub

30        Select Case g.Col
          Case 0:
40            If g.TextMatrix(g.RowSel, 2) <> "" Then
50                Exit Sub
60            End If
70            cmbWard.Move g.Left + g.CellLeft, g.Top + g.CellTop, g.CellWidth
80            cmbWard.Visible = True
90            cmbWard.SetFocus
100       Case 1:
110           If g.TextMatrix(g.RowSel, 2) <> "" Then
120               Exit Sub
130           End If
140           cmbClinician.Move g.Left + g.CellLeft, g.Top + g.CellTop, g.CellWidth
150           cmbClinician.Visible = True
160           cmbClinician.SetFocus
170       Case 2:
180           If g.TextMatrix(g.RowSel, 1) <> "" Or g.TextMatrix(g.RowSel, 0) <> "" Then
190               Exit Sub
200           End If
210           cmbGP.Move g.Left + g.CellLeft, g.Top + g.CellTop, g.CellWidth
220           cmbGP.Visible = True
230           cmbGP.SetFocus
240       Case 3:
250           Found = False
260           FaxNumber = ""
270           WCG = ""
280           If g.TextMatrix(g.Row, 0) <> "" Then
290               Found = True
300               FaxNumber = IsFaxable("Wards", g.TextMatrix(g.Row, 0))
310               If FaxNumber <> "" Then
320                   WCG = "W"
330               End If
340           End If
350           If g.TextMatrix(g.Row, 1) <> "" Then
360               Found = True
                  'FaxNumber = IsFaxable("Clinicians", g.TextMatrix(g.Row, 1))
                  'If FaxNumber <> "" Then
                  '  WCG = WCG & "C"
                  'End If
370           End If
380           If g.TextMatrix(g.Row, 2) <> "" Then
390               Found = True
400               FaxNumber = IsFaxable("GPs", g.TextMatrix(g.Row, 2))
410               If FaxNumber <> "" Then
420                   WCG = "G"
430               End If
440           End If

450           If Found Then
460               If g.TextMatrix(g.Row, 3) = "Printer" Then
470                   If FaxNumber <> "" Then
480                       g.TextMatrix(g.Row, 3) = "FAX"
490                       g.TextMatrix(g.Row, 4) = FaxNumber
500                   End If
510               Else
520                   g.TextMatrix(g.Row, 3) = "Printer"
530                   g.TextMatrix(g.Row, 4) = ""
540               End If
550           End If

560       Case 4:
570           If g.TextMatrix(g.Row, 3) = "Printer" Then
580               cmbPrinter.Move g.Left + g.CellLeft, g.Top + g.CellTop, g.CellWidth
590               cmbPrinter.Visible = True
600               cmbPrinter.SetFocus
610           End If
620       End Select

630       DataChanged = True

640       Exit Sub

g_Click_Error:

          Dim strES As String
          Dim intEL As Integer

650       intEL = Erl
660       strES = Err.Description
670       LogError "frmCopyTo", "g_Click", intEL, strES

End Sub

Private Sub FillPrinterList()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo FillPrinterList_Error

20        cmbPrinter.Clear

30        cmbPrinter.AddItem "Use Default"
40        cmbPrinter.AddItem ""

50        sql = "SELECT * FROM InstalledPrinters WHERE " & _
                "Location = 'Lab'"
60        Set tb = New Recordset
70        RecOpenServer 0, tb, sql

80        Do While Not tb.EOF
90            cmbPrinter.AddItem tb!PrinterName & ""
100           tb.MoveNext
110       Loop

120       Exit Sub

FillPrinterList_Error:

          Dim strES As String
          Dim intEL As Integer

130       intEL = Erl
140       strES = Err.Description
150       LogError "frmCopyTo", "FillPrinterList", intEL, strES, sql

End Sub

Private Sub g_GotFocus()

10        On Error GoTo g_GotFocus_Error

20        Select Case g.Col
          Case 0:
30            If cmbWard.Visible Then
40                g = cmbWard
50                cmbWard.Visible = False
60                g.TextMatrix(g.Row, 2) = ""
70            End If
80        Case 1:
90            If cmbClinician.Visible Then
100               g = cmbClinician
110               cmbClinician.Visible = False
120               g.TextMatrix(g.Row, 2) = ""
130           End If
140       Case 2:
150           If cmbGP.Visible Then
160               g = cmbGP
170               cmbGP.Visible = False
180               g.TextMatrix(g.Row, 0) = ""
190               g.TextMatrix(g.Row, 1) = ""
200           End If
210       Case 4:
220           If cmbPrinter.Visible Then
230               g = cmbPrinter
240               cmbPrinter.Visible = False
250           End If
260       End Select

270       If g.TextMatrix(g.Row, 3) = "" Then
280           g.TextMatrix(g.Row, 3) = "Printer"
290       End If

300       cmbWard.Visible = False
310       cmbClinician.Visible = False
320       cmbGP.Visible = False
330       cmbPrinter.Visible = False

340       AdjustBlanks

350       Exit Sub

g_GotFocus_Error:

          Dim strES As String
          Dim intEL As Integer

360       intEL = Erl
370       strES = Err.Description
380       LogError "frmCopyTo", "g_GotFocus", intEL, strES

End Sub

Private Sub mnuClinicians_Click()

          Dim f As Form
          Dim Options(0 To 5) As String
          Dim n As Integer
          Dim Sel As Integer

10        For n = 0 To 5
20            Options(n) = Choose(n + 1, "8", "10", "15", "20", "25", "30")
30        Next
40        Set f = New fcdrDBox
50        With f
60            .Options = Options
70            .Prompt = "Enter number of items visible in 'Clinicians'"
80            .Show 1
90            Sel = .ReturnValue
100       End With
110       Unload f
120       Set f = Nothing

130       If Sel <> 0 Then
140           SaveOptionSetting "ClinicianCCListLength", Sel
150       End If

End Sub

Private Sub mnuExit_Click()

10        Unload Me

End Sub


Private Sub mnuGPs_Click()

          Dim f As Form
          Dim Options(0 To 5) As String
          Dim n As Integer
          Dim Sel As Integer

10        For n = 0 To 5
20            Options(n) = Choose(n + 1, "8", "10", "15", "20", "25", "30")
30        Next
40        Set f = New fcdrDBox
50        With f
60            .Options = Options
70            .Prompt = "Enter number of items visible in 'GPs'"
80            .Show 1
90            Sel = .ReturnValue
100       End With
110       Unload f
120       Set f = Nothing

130       If Sel <> 0 Then
140           SaveOptionSetting "GPCCListLength", Sel
150       End If

End Sub


Private Sub mnuWards_Click()

          Dim f As Form
          Dim Options(0 To 5) As String
          Dim n As Integer
          Dim Sel As Integer

10        For n = 0 To 5
20            Options(n) = Choose(n + 1, "8", "10", "15", "20", "25", "30")
30        Next
40        Set f = New fcdrDBox
50        With f
60            .Options = Options
70            .Prompt = "Enter number of items visible in 'Wards'"
80            .Show 1
90            Sel = .ReturnValue
100       End With
110       Unload f
120       Set f = Nothing

130       If Sel <> 0 Then
140           SaveOptionSetting "WardCCListLength", Sel
150       End If

End Sub


