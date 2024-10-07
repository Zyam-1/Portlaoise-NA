VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmEditMicroExternals 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Microbiology External Requests"
   ClientHeight    =   8535
   ClientLeft      =   2355
   ClientTop       =   5325
   ClientWidth     =   13470
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   13470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdMedibridgeResults 
      Appearance      =   0  'Flat
      Caption         =   "NVRL           St. James"
      Height          =   1000
      Left            =   11970
      Picture         =   "frmEditMicroExternals.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Order External Tests"
      Top             =   60
      Width           =   1200
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   1000
      Left            =   11970
      Picture         =   "frmEditMicroExternals.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   17
      Tag             =   "bprint"
      ToolTipText     =   "Print Result"
      Top             =   3375
      Width           =   1200
   End
   Begin VB.CommandButton cmdSetPrinter 
      Height          =   1000
      Left            =   11970
      Picture         =   "frmEditMicroExternals.frx":1794
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Back Colour Normal:Automatic Printer Selection.Back Colour Red:-Forced"
      Top             =   2370
      Width           =   1200
   End
   Begin MSComCtl2.UpDown udSampleID 
      Height          =   225
      Left            =   360
      TabIndex        =   7
      Top             =   510
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   397
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txtSampleID"
      BuddyDispid     =   196613
      OrigLeft        =   2430
      OrigTop         =   540
      OrigRight       =   2715
      OrigBottom      =   780
      Max             =   9999999
      Min             =   1
      Orientation     =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtSampleID 
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
      Left            =   270
      TabIndex        =   6
      Text            =   "12345678"
      Top             =   210
      Width           =   1425
   End
   Begin VB.CommandButton cmdOrderTests 
      Caption         =   "Order Tests"
      Height          =   1000
      Left            =   11970
      Picture         =   "frmEditMicroExternals.frx":265E
      Style           =   1  'Graphical
      TabIndex        =   4
      Tag             =   "bOrder"
      ToolTipText     =   "Order Tests for Sample"
      Top             =   5385
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   1000
      Left            =   11970
      Picture         =   "frmEditMicroExternals.frx":3528
      Style           =   1  'Graphical
      TabIndex        =   3
      Tag             =   "101"
      ToolTipText     =   "Exit Screen"
      Top             =   7395
      Width           =   1200
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Enabled         =   0   'False
      Height          =   1000
      Left            =   11970
      Picture         =   "frmEditMicroExternals.frx":43F2
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Delete Test"
      Top             =   4380
      Width           =   1200
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Details"
      Height          =   1000
      Left            =   11970
      Picture         =   "frmEditMicroExternals.frx":52BC
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Save Changes"
      Top             =   6390
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "Export to Excel"
      Height          =   1000
      Left            =   11970
      Picture         =   "frmEditMicroExternals.frx":6186
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1080
      Width           =   1200
   End
   Begin MSFlexGridLib.MSFlexGrid grdExt 
      Height          =   7455
      Left            =   210
      TabIndex        =   25
      Top             =   780
      Width           =   11490
      _ExtentX        =   20267
      _ExtentY        =   13150
      _Version        =   393216
      Cols            =   13
      RowHeightMin    =   400
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   2
      AllowUserResizing=   3
      FormatString    =   $"frmEditMicroExternals.frx":BD98
   End
   Begin VB.Label lblSampleTime 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   11100
      TabIndex        =   24
      Top             =   390
      Width           =   585
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Sample Time"
      Height          =   195
      Left            =   10110
      TabIndex        =   23
      Top             =   420
      Width           =   915
   End
   Begin VB.Label lblSampleDate 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   8220
      TabIndex        =   22
      Top             =   390
      Width           =   1110
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "SampleDate"
      Height          =   195
      Left            =   7260
      TabIndex        =   21
      Top             =   420
      Width           =   870
   End
   Begin VB.Label lblClinDetails 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   8220
      TabIndex        =   20
      Top             =   60
      Width           =   3480
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Clin Details"
      Height          =   195
      Left            =   7350
      TabIndex        =   19
      Top             =   90
      Width           =   780
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
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
      Left            =   11940
      TabIndex        =   18
      Top             =   2085
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label lblSex 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
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
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   6030
      TabIndex        =   15
      Top             =   75
      Width           =   750
   End
   Begin VB.Label lblDoB 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
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
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   4530
      TabIndex        =   14
      Top             =   75
      Width           =   1155
   End
   Begin VB.Label lblName 
      BackColor       =   &H0080FFFF&
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
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   2610
      TabIndex        =   13
      Top             =   360
      Width           =   4170
   End
   Begin VB.Label lblChart 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
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
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   2610
      TabIndex        =   12
      Top             =   75
      Width           =   1515
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Sex"
      Height          =   195
      Left            =   5730
      TabIndex        =   11
      Top             =   90
      Width           =   270
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "DoB"
      Height          =   195
      Left            =   4200
      TabIndex        =   10
      Top             =   105
      Width           =   315
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      Height          =   195
      Left            =   2160
      TabIndex        =   9
      Top             =   360
      Width           =   420
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Chart"
      Height          =   195
      Left            =   2190
      TabIndex        =   8
      Top             =   105
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "SampleID"
      Height          =   195
      Left            =   660
      TabIndex        =   5
      Top             =   30
      Width           =   690
   End
End
Attribute VB_Name = "frmEditMicroExternals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Activated As Boolean
Private pPrintToPrinter As String

Private pDepartment As String

Private pSampleID As String

Private pClinDetails As String
Private pSampleDate As String
Private pSampleTime As String
Private pSex As String

Public Property Let ClinDetails(ByVal sNewValue As String)

10        pClinDetails = sNewValue
20        lblClinDetails = pClinDetails

End Property



Public Property Let Department(ByVal sNewValue As String)

10        pDepartment = sNewValue

End Property


Public Property Let SampleDate(ByVal sNewValue As String)

10        pSampleDate = sNewValue
20        lblSampleDate = pSampleDate

End Property

Public Property Let SampleID(ByVal sNewValue As String)

10        pSampleID = sNewValue
20        txtSampleID = pSampleID

End Property


Public Property Let SampleTime(ByVal sNewValue As String)

10        pSampleTime = sNewValue
20        lblSampleTime = pSampleTime

End Property


Public Property Let sex(ByVal sNewValue As String)

10        pSex = sNewValue
20        lblSex = pSex

End Property




Private Sub EnableSave(ByVal Enable As Boolean)

10        cmdSave.Visible = Enable
20        udSampleID.Enabled = Not Enable
30        txtSampleID.Enabled = Not Enable

End Sub

Private Sub LoadDemographics()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo LoadDemographics_Error

20        lblName = ""
30        lblDoB = ""
40        lblChart = ""
50        lblSex = ""

60        If txtSampleID <> "" Then

70            sql = "SELECT PatName, COALESCE(Sex, '') AS Sex, " & _
                    "COALESCE(DoB, 0) AS DoB, " & _
                    "AandE FROM Demographics WHERE " & _
                    "SampleID = '" & Val(txtSampleID) + SysOptMicroOffset(0) & "'"
80            Set tb = New Recordset
90            RecOpenServer 0, tb, sql
100           If Not tb.EOF Then
110               lblName = Trim$(tb!PatName & "")
120               If IsDate(tb!Dob) Then
130                   lblDoB = Format$(tb!Dob, "dd/MM/yyyy")
140               Else
150                   lblDoB = ""
160               End If
170               lblChart = tb!AandE & ""
180               Select Case Left$(tb!sex, 1)
                  Case "M": lblSex = "Male"
190               Case "F": lblSex = "Female"
200               End Select
210           End If

220       End If

230       Exit Sub

LoadDemographics_Error:

          Dim strES As String
          Dim intEL As Integer

240       intEL = Erl
250       strES = Err.Description
260       LogError "frmEditMicroExternals", "LoadDemographics", intEL, strES, sql

End Sub

Private Sub LoadExternals()

          Dim sql As String
          Dim tb As New Recordset
          Dim s As String
          Dim TestName As String

10        On Error GoTo LoadExternals_Error

20        If txtSampleID = "" Then Exit Sub


30        ClearFGrid grdExt

40        sql = "SELECT * FROM ExtResults WHERE " & _
                "SampleID = '" & Val(txtSampleID) + SysOptMicroOffset(0) & "' " & _
                "ORDER BY OrderList"
50        Set tb = New Recordset
60        RecOpenServer 0, tb, sql
70        Do While Not tb.EOF
              '80        If IsNumeric(tb!Analyte) Then
              '90          TestName = eNumber2Name(Trim(tb!Analyte & ""), "General")
              '100       Else
80            TestName = tb!Analyte & ""
              '120       End If
90            s = TestName & "" & vbTab & _
                  Trim(tb!Result) & "" & vbTab & _
                  tb!NormalRange & "" & vbTab & _
                  tb!Units & "" & vbTab & _
                  tb!SendTo & "" & vbTab
100           If Not IsNull(tb!SentDate) Then
110               s = s & Format(tb!SentDate, "dd/mmm/yyyy") & vbTab
120           Else
130               s = s & vbTab
140           End If

150           If Not IsNull(tb!RetDate) Then
160               s = s & Format(tb!RetDate, "dd/mmm/yyyy") & vbTab
170           Else
180               s = s & vbTab
190           End If
200           s = s & Trim(tb!SapCode & "") & vbTab

210           If Trim(tb!Valid & "") <> "" And tb!Valid & "" = 1 Then
220               s = s & "V" & vbTab
230           Else
240               s = s & vbTab
250           End If
260           s = s & tb!InterimReportDate & "" & vbTab & _
                  tb!InterimReportComment & "" & vbTab & _
                  tb!FinalReportDate & "" & vbTab & _
                  tb!FinalReportComment & ""
270           grdExt.AddItem s
280           tb.MoveNext
290       Loop
300       FixG grdExt

310       Exit Sub

LoadExternals_Error:

          Dim strES As String
          Dim intEL As Integer

320       intEL = Erl
330       strES = Err.Description
340       LogError "frmEditMicroExternals", "LoadExternals", intEL, strES, sql

End Sub


Private Sub cmdDelete_Click()

          Dim TestName As String
          Dim sql As String
          Dim n As Integer

10        On Error GoTo cmdDelete_Click_Error

20        cmdDelete.Enabled = False

30        grdExt.Col = 0
40        For n = 1 To grdExt.Rows - 1
50            grdExt.row = n
60            If grdExt.CellBackColor = vbYellow Then
70                TestName = grdExt.TextMatrix(n, 0)

80                If iMsg("Delete " & TestName & "?", vbQuestion + vbYesNo) = vbYes Then
90                    sql = "DELETE FROM ExtResults WHERE " & _
                            "SampleID = '" & Val(txtSampleID) + SysOptMicroOffset(0) & "' " & _
                            "AND Department = 'Micro' " & _
                            "AND Analyte = '" & TestName & "'"
100                   Cnxn(0).Execute sql
110               End If

120               If grdExt.Rows = 2 Then
130                   grdExt.AddItem ""
140               End If
150               grdExt.RemoveItem n
160               Exit For
170           End If
180       Next

190       Exit Sub

cmdDelete_Click_Error:

          Dim strES As String
          Dim intEL As Integer

200       intEL = Erl
210       strES = Err.Description
220       LogError "frmEditMicroExternals", "cmdDelete_Click", intEL, strES, sql

End Sub

Private Sub cmdExcel_Click()

10        cmdDelete.Enabled = False
20        ExportFlexGrid grdExt, Me

End Sub

Private Sub cmdMedibridgeResults_Click()

      Dim MediBridgePathToViewer As String

10    On Error GoTo cmdMedibridgeResults_Click_Error

20    If cmdMedibridgeResults.BackColor <> vbYellow Then Exit Sub
      'view external results (Path changed to app.path because custom path was creating trouble
30    MediBridgePathToViewer = App.Path & "\MediBridgeViewer.exe "             ' GetOptionSetting("MedibridgePathToViewer", "")
40    If MediBridgePathToViewer <> "" Then
50        Shell MediBridgePathToViewer & " /SampleID=" & txtSampleID + SysOptMicroOffset(0) & _
                " /UserName=""" & UserName & """" & _
                " /Password=""" & UserPass & """" & _
                " /Department=""Medibridge""", vbNormalFocus
60    End If

70    Exit Sub

cmdMedibridgeResults_Click_Error:

      Dim strES As String
      Dim intEL As Integer

80    intEL = Erl
90    strES = Err.Description
100   LogError "frmEditMicroExternals", "cmdMedibridgeResults_Click", intEL, strES

End Sub

Private Sub cmdOrderTests_Click()

10        cmdDelete.Enabled = False

20        With frmAddToTests
30            .sex = lblSex
40            .SampleID = Val(txtSampleID) + SysOptMicroOffset(0)
50            .ClinDetails = lblClinDetails
60            .SampleDate = lblSampleDate
70            .SampleTime = lblSampleTime
80            .Department = "Micro"
90            .Show 1
100       End With

110       LoadExternals

120       EnableSave True

End Sub

Private Sub cmdCancel_Click()

10        If cmdSave.Visible Then
20            If iMsg("Cancel without Saving?", vbQuestion + vbYesNo) = vbNo Then
30                Exit Sub
40            End If
50        End If

60        Unload Me

End Sub


Private Sub cmdPrint_Click()

          Dim tb As New Recordset
          Dim sql As String
          Dim SID As Double
          Dim Ward As String
          Dim Clinician As String
          Dim GP As String

10        On Error GoTo cmdPrint_Click_Error

20        txtSampleID = Format(Val(txtSampleID))
30        SID = Val(txtSampleID) + SysOptMicroOffset(0)

40        sql = "SELECT Ward, Clinician, GP FROM Demographics WHERE " & _
                "SampleID = " & SID
50        Set tb = New Recordset
60        RecOpenServer 0, tb, sql
70        If tb.EOF Then
80            Exit Sub
90        End If
100       Ward = tb!Ward & ""
110       Clinician = tb!Clinician & ""
120       GP = tb!GP & ""

130       LogTimeOfPrinting SID, "X"
140       sql = "SELECT * FROM PrintPending WHERE " & _
                "Department = 'V' " & _
                "AND SampleID = " & SID
150       Set tb = New Recordset
160       RecOpenServer 0, tb, sql
170       If tb.EOF Then
180           tb.AddNew
190       End If
200       tb!SampleID = SID
210       tb!Department = "V"
220       tb!Initiator = UserName
230       tb!Ward = Ward
240       tb!Clinician = Clinician
250       tb!GP = GP
260       tb!UsePrinter = pPrintToPrinter
270       tb!pTime = Now
280       tb.Update

290       Exit Sub

cmdPrint_Click_Error:

          Dim strES As String
          Dim intEL As Integer

300       intEL = Erl
310       strES = Err.Description
320       LogError "frmEditMicroExternals", "cmdPrint_Click", intEL, strES, sql

End Sub

Private Sub cmdSave_Click()

          Dim tb As New Recordset
          Dim sql As String
          Dim Y As Integer
          Dim TestName As String
          Dim TempDate As String

10        On Error GoTo cmdSave_Click_Error

20        cmdDelete.Enabled = False

30        EnableSave False

40        txtSampleID = Format(Val(txtSampleID))
50        If Val(txtSampleID) = 0 Then Exit Sub

60        For Y = 1 To grdExt.Rows - 1
70            If Trim(grdExt.TextMatrix(Y, 0)) <> "" Then
80                TestName = grdExt.TextMatrix(Y, 0)
90                sql = "SELECT * FROM ExtResults WHERE " & _
                        "SampleID = '" & Val(txtSampleID) + SysOptMicroOffset(0) & "' " & _
                        "AND Department = 'Micro' " & _
                        "AND Analyte = '" & TestName & "'"
100               Set tb = New Recordset
110               RecOpenServer 0, tb, sql
120               If tb.EOF Then
130                   tb.AddNew
140               End If
150               tb!Department = "Micro"
160               tb!SampleID = Val(txtSampleID) + SysOptMicroOffset(0)
170               tb!Analyte = TestName
180               tb!Result = grdExt.TextMatrix(Y, 1)
190               tb!Units = grdExt.TextMatrix(Y, 3)

200               tb!SendTo = grdExt.TextMatrix(Y, 4)
210               TempDate = grdExt.TextMatrix(Y, 5)
220               If IsDate(TempDate) Then
230                   tb!SentDate = Format$(TempDate, "dd/MMM/yyyy")
240               Else
250                   tb!SentDate = Null
260               End If
270               TempDate = grdExt.TextMatrix(Y, 6)
280               If IsDate(TempDate) Then
290                   tb!RetDate = Format$(TempDate, "dd/MMM/yyyy")
300               Else
310                   tb!RetDate = Null
320               End If

330               tb!SapCode = grdExt.TextMatrix(Y, 7)


340               TempDate = grdExt.TextMatrix(Y, 9)
350               If IsDate(TempDate) Then
360                   tb!InterimReportDate = Format$(TempDate, "dd/MMM/yyyy")
370               Else
380                   tb!InterimReportDate = Null
390               End If

400               tb!InterimReportComment = grdExt.TextMatrix(Y, 10)

410               TempDate = grdExt.TextMatrix(Y, 11)
420               If IsDate(TempDate) Then
430                   tb!FinalReportDate = Format$(TempDate, "dd/MMM/yyyy")
440               Else
450                   tb!FinalReportDate = Null
460               End If

470               tb!FinalReportComment = grdExt.TextMatrix(Y, 12)

480               tb!UserName = UserName
490               tb!DateTimeOfRecord = Format$(Now, "dd/MMM/yyyy HH:mm:ss")
500               tb.Update
510           End If
520       Next

530       Unload Me

540       Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

550       intEL = Erl
560       strES = Err.Description
570       LogError "frmEditMicroExternals", "cmdSave_Click", intEL, strES, sql

End Sub

Private Sub cmdSetPrinter_Click()

10        On Error GoTo cmdSetPrinter_Click_Error

20        Set frmForcePrinter.f = frmEditMicroExternals
30        frmForcePrinter.Show 1

40        If pPrintToPrinter = "Automatic Selection" Then
50            pPrintToPrinter = ""
60        End If

70        If pPrintToPrinter <> "" Then
80            cmdSetPrinter.BackColor = vbRed
90            cmdSetPrinter.ToolTipText = "Print Forced to " & pPrintToPrinter
100       Else
110           cmdSetPrinter.BackColor = vbButtonFace
120           pPrintToPrinter = ""
130           cmdSetPrinter.ToolTipText = "Printer Selected Automatically"
140       End If

150       Exit Sub

cmdSetPrinter_Click_Error:

          Dim strES As String
          Dim intEL As Integer

160       intEL = Erl
170       strES = Err.Description
180       LogError "frmEditMicroExternals", "cmdSetPrinter_Click", intEL, strES

End Sub

Private Sub Form_Activate()

10    If Activated Then Exit Sub

20    Activated = True

30    LoadAllDetails

40    LoadExt

End Sub

Public Property Get PrintToPrinter() As String

10        On Error GoTo PrintToPrinter_Error

20        PrintToPrinter = pPrintToPrinter

30        Exit Property

PrintToPrinter_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmEditMicroExternals", "PrintToPrinter", intEL, strES

End Property


Public Property Let PrintToPrinter(ByVal strNewValue As String)

10        On Error GoTo PrintToPrinter_Error

20        pPrintToPrinter = strNewValue

30        Exit Property

PrintToPrinter_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmEditMicroExternals", "PrintToPrinter", intEL, strES

End Property


Private Sub Form_Load()

10        Activated = False
20        grdExt.ColWidth(2) = 1125
30        grdExt.ColWidth(3) = 1125
40        grdExt.ColWidth(5) = 1125

End Sub
Private Sub LoadExt()

Dim sql As String
Dim tb As New Recordset
Dim Deltatb As Recordset
Dim s As String
Dim TestName As String
Dim PreviousDate As String
Dim PreviousRec As Long
Dim sn As New Recordset
Dim n As Integer
Dim i As Integer


On Error GoTo LoadExt_Error

If txtSampleID = "" Then Exit Sub

sql = "SELECT Count(*) AS Cnt FROM MediBridgeResults WHERE SampleID = " & txtSampleID + SysOptMicroOffset(0)
Set tb = New Recordset
RecOpenServer 0, tb, sql
If tb!Cnt > 0 Then
    cmdMedibridgeResults.BackColor = vbYellow
Else
    cmdMedibridgeResults.BackColor = &H8000000F
End If

Exit Sub

LoadExt_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "frmEditMicroExternals", "LoadExt", intEL, strES, sql

End Sub

Private Sub Form_Unload(Cancel As Integer)
10        pPrintToPrinter = ""
End Sub

Private Sub grdExt_Click()

          Dim f As Form
          Dim Prompt As String
          Dim Str As String

10        On Error GoTo grdExt_Click_Error

20        cmdDelete.Enabled = False

30        With grdExt
40            If .MouseRow = 0 Then Exit Sub
50            Select Case .MouseCol
              Case 0:
60                If .TextMatrix(.MouseRow, 0) <> "" Then
70                    .Col = 0
80                    If .CellBackColor = vbYellow Then
90                        .CellBackColor = vbDesktop
100                       .CellForeColor = &H80000018
110                   Else
120                       .CellBackColor = vbYellow
130                       .CellForeColor = 1
140                       cmdDelete.Enabled = True
150                   End If
160               End If
170           Case 1:
180               Prompt = "Enter result for " & grdExt.TextMatrix(grdExt.row, 0)
190               Str = iBOX(Prompt, , grdExt.TextMatrix(grdExt.row, 1))
200               If Str <> "" Then
210                   grdExt.TextMatrix(grdExt.row, 1) = Str
220                   grdExt.TextMatrix(grdExt.row, 6) = Format(Now, "dd/mmm/yyyy")
230               End If
240           Case 7:
250               Prompt = "Enter Sap Code for " & grdExt.TextMatrix(grdExt.row, 0)
260               Str = iBOX(Prompt, , grdExt.TextMatrix(grdExt.row, 1))
270               If Str <> "" Then
280                   grdExt.TextMatrix(grdExt.row, 7) = Str
290               End If
300           Case 5, 6, 9, 11:
310               Set f = New frmEnterDate
320               f.DateVal = .TextMatrix(.row, .Col)
330               f.Show 1
340               .TextMatrix(.row, .Col) = f.DateVal
350               Unload f
360               Set f = Nothing
370               EnableSave True

380           Case 10, 12:
390               Set f = New frmEnterComment
400               f.txtComment = .TextMatrix(.row, .Col)
410               f.Show 1
420               .TextMatrix(.row, .Col) = f.txtComment
430               Unload f
440               Set f = Nothing
450               EnableSave True

460           End Select
470       End With

480       Exit Sub

grdExt_Click_Error:

          Dim strES As String
          Dim intEL As Integer

490       intEL = Erl
500       strES = Err.Description
510       LogError "frmEditMicroExternals", "grdExt_Click", intEL, strES

End Sub

Private Sub LoadAllDetails()

10        LoadDemographics
20        LoadExternals
30        cmdSave.Visible = False
40        cmdDelete.Enabled = False

End Sub



Private Sub txtSampleID_LostFocus()

10        LoadAllDetails

End Sub

Private Sub UDSampleID_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        LoadAllDetails

End Sub


