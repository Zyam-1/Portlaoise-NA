VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frmHiCyReport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Histology/Cytology Report"
   ClientHeight    =   8820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14730
   Icon            =   "frmHiCyReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8820
   ScaleWidth      =   14730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtNoCopies 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   13440
      TabIndex        =   33
      Text            =   "3"
      Top             =   5520
      Width           =   375
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   870
      Left            =   13410
      Picture         =   "frmHiCyReport.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6330
      Width           =   1185
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Exit"
      Height          =   870
      Left            =   13410
      Picture         =   "frmHiCyReport.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7740
      Width           =   1185
   End
   Begin VB.TextBox txtReport 
      BackColor       =   &H00FFFFFF&
      Height          =   6690
      Left            =   795
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1920
      Width           =   12480
   End
   Begin ComCtl2.UpDown udNoCopies 
      Height          =   420
      Left            =   13860
      TabIndex        =   35
      Top             =   5520
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   741
      _Version        =   327681
      Value           =   1
      BuddyControl    =   "txtNoCopies"
      BuddyDispid     =   196609
      OrigLeft        =   420
      OrigTop         =   210
      OrigRight       =   675
      OrigBottom      =   585
      Max             =   9
      Min             =   1
      SyncBuddy       =   -1  'True
      Wrap            =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "No Of Copies"
      Height          =   195
      Left            =   13440
      TabIndex        =   34
      Top             =   5280
      Width           =   945
   End
   Begin VB.Label lblNatureOfSpecimen 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   810
      TabIndex        =   32
      Top             =   1215
      Width           =   8550
   End
   Begin VB.Label lblRundate 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   12660
      TabIndex        =   30
      Top             =   135
      Width           =   1770
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "Sample Date"
      Height          =   195
      Left            =   11670
      TabIndex        =   29
      Top             =   990
      Width           =   915
   End
   Begin VB.Label lblSampleDate 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   12660
      TabIndex        =   28
      Top             =   915
      Width           =   1770
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "Received Date"
      Height          =   195
      Left            =   11505
      TabIndex        =   27
      Top             =   600
      Width           =   1080
   End
   Begin VB.Label lblRecDate 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   12660
      TabIndex        =   26
      Top             =   525
      Width           =   1770
   End
   Begin VB.Label lblValDate 
      Caption         =   "hhh"
      Height          =   255
      Left            =   9540
      TabIndex        =   25
      Top             =   1305
      Width           =   3930
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Address"
      Height          =   195
      Left            =   180
      TabIndex        =   24
      Top             =   540
      Width           =   570
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Age"
      Height          =   195
      Left            =   8265
      TabIndex        =   23
      Top             =   540
      Width           =   285
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Sex"
      Height          =   195
      Left            =   9570
      TabIndex        =   22
      Top             =   540
      Width           =   270
   End
   Begin VB.Label lblSex 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9915
      TabIndex        =   21
      Top             =   465
      Width           =   1215
   End
   Begin VB.Label lblAge 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8610
      TabIndex        =   20
      Top             =   465
      Width           =   735
   End
   Begin VB.Label lblAddress 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   810
      TabIndex        =   19
      Top             =   465
      Width           =   5685
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Gp"
      Height          =   195
      Left            =   4950
      TabIndex        =   18
      Top             =   930
      Width           =   210
   End
   Begin VB.Label lblGp 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5205
      TabIndex        =   17
      Top             =   885
      Width           =   4155
   End
   Begin VB.Label Label7 
      Caption         =   "Nature Of Specimen"
      Height          =   420
      Left            =   45
      TabIndex        =   16
      Top             =   1290
      Width           =   735
   End
   Begin VB.Label lblClinician 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   810
      TabIndex        =   15
      Top             =   885
      Width           =   4065
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Clinician"
      Height          =   195
      Left            =   165
      TabIndex        =   14
      Top             =   900
      Width           =   585
   End
   Begin VB.Label lblValid 
      Caption         =   "Validated"
      Height          =   255
      Left            =   9540
      TabIndex        =   12
      Top             =   1590
      Visible         =   0   'False
      Width           =   3930
   End
   Begin VB.Label lblChart 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9900
      TabIndex        =   4
      Top             =   60
      Width           =   1230
   End
   Begin VB.Label lblDob 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "88/88/8888"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6960
      TabIndex        =   3
      Top             =   465
      Width           =   1215
   End
   Begin VB.Label lblName 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3210
      TabIndex        =   2
      Top             =   60
      Width           =   6135
   End
   Begin VB.Label lblSampleID 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   810
      TabIndex        =   1
      Top             =   67
      Width           =   1770
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      Height          =   195
      Left            =   2760
      TabIndex        =   9
      Top             =   135
      Width           =   420
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Chart"
      Height          =   195
      Left            =   9450
      TabIndex        =   8
      Top             =   135
      Width           =   375
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Dob"
      Height          =   195
      Left            =   6585
      TabIndex        =   10
      Top             =   540
      Width           =   300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Sample ID"
      Height          =   195
      Left            =   45
      TabIndex        =   7
      Top             =   135
      Width           =   735
   End
   Begin VB.Label lblYear 
      Height          =   195
      Left            =   2430
      TabIndex        =   11
      Top             =   3780
      Width           =   1005
   End
   Begin VB.Label lblNotValid 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Not Valid"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   9750
      TabIndex        =   13
      Top             =   870
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "Run Date"
      Height          =   195
      Left            =   11895
      TabIndex        =   31
      Top             =   210
      Width           =   690
   End
End
Attribute VB_Name = "frmHiCyReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pSampleID As String

Private pPrintToPrinter As String


Public Property Let PrintToPrinter(ByVal strNewValue As String)

10        On Error GoTo PrintToPrinter_Error

20        pPrintToPrinter = strNewValue

30        Exit Property

PrintToPrinter_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmHiCyReport", "PrintToPrinter", intEL, strES

End Property
Public Property Get PrintToPrinter() As String

10        On Error GoTo PrintToPrinter_Error

20        PrintToPrinter = pPrintToPrinter

30        Exit Property

PrintToPrinter_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmHiCyReport", "PrintToPrinter", intEL, strES

End Property

Private Sub cmdCancel_Click()

10        Unload Me

End Sub

Public Sub Load_Report(ByVal Dept As String, _
                       ByVal Samp As String, _
                       ByVal Year As String)

10        On Error GoTo Load_Report_Error

20        pSampleID = Samp
30        lblYear = Year

40        If Dept = "C" Then
50            Load_Cyto
60        Else
70            Load_Histo
80        End If

90        Exit Sub

Load_Report_Error:

          Dim strES As String
          Dim intEL As Integer

100       intEL = Erl
110       strES = Err.Description
120       LogError "frmHiCyReport", "Load_Report", intEL, strES

End Sub


Private Sub Load_Cyto()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo Load_Cyto_Error

20        If Val(pSampleID) = 0 Then Exit Sub


30        sql = "SELECT * FROM Demographics WHERE " & _
                "SampleID = '" & pSampleID & "' " & _
                "AND hYear = '" & lblYear & "'"

40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        If Not tb.EOF Then
70            lblClinician = tb!Clinician & ""
80            lblGP = tb!GP & ""
90            lblAddress = tb!Addr0 & ""
100           lblAge = tb!Age & ""
110           lblSex = tb!sex & ""
120           If Trim(tb!SampleDate & "") <> "" Then lblSampledate = Format(tb!SampleDate, "dd/mmm/yyyy")
130           If Trim(tb!Rundate & "") <> "" Then lblRundate = Format(tb!Rundate, "dd/mmm/yyyy")
140           If Trim(tb!RecDate & "") <> "" Then lblRecDate = Format(tb!RecDate, "dd/mmm/yyyy")
150           If (IsNull(tb!cytovalid) Or tb!cytovalid = 0) And UserMemberOf <> "Managers" And UserMemberOf <> "Users" Then
160               txtReport = "Not Validated"
170               lblNotValid.Visible = True


180           Else
190               lblValid.Visible = True
200               sql = "SELECT * FROM CytoResults WHERE " & _
                        "SampleID = '" & pSampleID & "' " & _
                        "AND hyear = '" & Trim$(lblYear) & "'"

210               Set tb = New Recordset
220               RecOpenServer 0, tb, sql
230               If Not tb.EOF Then
240                   lblNatureOfSpecimen = tb!NatureOfSpecimen & ""
250                   If Trim$(tb!natureofspecimenB & "") <> "" Then
260                       lblNatureOfSpecimen = lblNatureOfSpecimen & "  :  " & tb!natureofspecimenB
270                   End If
280                   If Trim$(tb!natureofspecimenC & "") <> "" Then
290                       lblNatureOfSpecimen = lblNatureOfSpecimen & "  :  " & tb!natureofspecimenC
300                   End If
310                   If Trim$(tb!natureofspecimenD & "") <> "" Then
320                       lblNatureOfSpecimen = lblNatureOfSpecimen & "  :  " & tb!natureofspecimenD
330                   End If

340                   txtReport = tb!cytoreport
350                   If Trim(tb!validdate & "") <> "" Then
360                       lblValDate = Format(tb!validdate, "dd/mmm/yyyy hh:mm")
370                   Else
380                       lblValDate = "No Date"
390                   End If
400               End If

410               sql = "SELECT * FROM Demographics WHERE " & _
                        "SampleID = '" & pSampleID & "' " & _
                        "AND hYear = '" & lblYear & "'"

420               Set tb = New Recordset
430               RecOpenServer 0, tb, sql
440               If Not tb.EOF Then
450                   lblClinician = tb!Clinician & ""
460                   lblGP = tb!GP & ""
470                   lblAddress = tb!Addr0 & ""
480                   lblAge = tb!Age & ""
490                   lblSex = tb!sex & ""
500                   If Trim(tb!SampleDate & "") <> "" Then lblSampledate = Format(tb!SampleDate, "dd/mmm/yyyy")
510                   If Trim(tb!Rundate & "") <> "" Then lblRundate = Format(tb!Rundate, "dd/mmm/yyyy")
520                   If Trim(tb!RecDate & "") <> "" Then lblRecDate = Format(tb!RecDate, "dd/mmm/yyyy")
530                   If tb!cytovalid = True Then
540                       lblValid.Visible = True
550                   Else
560                       lblNotValid.Visible = True
570                   End If
580               End If

590           End If
600       End If


610       txtReport.Locked = True

620       Exit Sub

Load_Cyto_Error:

          Dim strES As String
          Dim intEL As Integer

630       intEL = Erl
640       strES = Err.Description
650       LogError "frmHiCyReport", "Load_Cyto", intEL, strES, sql

End Sub

Private Sub Load_Histo()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo Load_Histo_Error

20        lblValid.Visible = False
30        lblValDate.Visible = False

40        sql = "SELECT Clinician, GP, Addr0, Age, Sex, SampleDate, RunDate, RecDate, " & _
                "COALESCE(HistoValid, 0) AS HistoValid FROM Demographics WHERE " & _
                "SampleID = '" & pSampleID & "' " & _
                "AND hYear = '" & lblYear & "'"

50        Set tb = New Recordset
60        RecOpenServer 0, tb, sql
70        If Not tb.EOF Then
80            lblClinician = tb!Clinician & ""
90            lblGP = tb!GP & ""
100           lblAddress = tb!Addr0 & ""
110           lblAge = tb!Age & ""
120           lblSex = tb!sex & ""
130           If Trim(tb!SampleDate & "") <> "" Then lblSampledate = Format(tb!SampleDate, "dd/mmm/yyyy")
140           If Trim(tb!Rundate & "") <> "" Then lblRundate = Format(tb!Rundate, "dd/mmm/yyyy")
150           If Trim(tb!RecDate & "") <> "" Then lblRecDate = Format(tb!RecDate, "dd/mmm/yyyy")
160           If tb!histovalid = 0 And UserMemberOf <> "Managers" And UserMemberOf <> "Users" Then
170               lblNotValid.Visible = True
180               txtReport = "Not Validated"
190           Else
200               lblValid.Visible = True

210               sql = "SELECT * FROM Historesults WHERE " & _
                        "SampleID = '" & pSampleID & "' " & _
                        "AND hyear = '" & lblYear & "'"

220               Set tb = New Recordset
230               RecOpenServer 0, tb, sql
240               If Not tb.EOF Then
250                   lblNatureOfSpecimen = "A: " & tb!NatureOfSpecimen & "" & "     "
260                   If tb!natureofspecimenB & "" <> "" Then lblNatureOfSpecimen = lblNatureOfSpecimen & "B: " & tb!natureofspecimenB & vbCrLf
270                   If tb!natureofspecimenC & "" <> "" Then lblNatureOfSpecimen = lblNatureOfSpecimen & "C: " & tb!natureofspecimenC & "     "
280                   If tb!natureofspecimenD & "" <> "" Then lblNatureOfSpecimen = lblNatureOfSpecimen & "D: " & tb!natureofspecimenD

290                   If Trim(tb!validdate & "") <> "" Then
300                       lblValDate = "Date : " & Format(tb!validdate, "dd/mmm/yyyy hh:mm")
310                       lblValDate.Visible = True
320                   Else
330                       lblValDate = "No Date"
340                   End If
350                   lblValid = "Validated by " & TechNameFor(tb!Username & "")
360                   txtReport = tb!historeport
370               End If
380           End If
390       End If

400       txtReport.Locked = True

410       Exit Sub

Load_Histo_Error:

          Dim strES As String
          Dim intEL As Integer

420       intEL = Erl
430       strES = Err.Description
440       LogError "frmHiCyReport", "Load_Histo", intEL, strES, sql

End Sub

Private Sub cmdPrint_Click()

          Dim sql As String
          Dim tb As Recordset
          Dim Dept As String
          Dim pTime As String

10        On Error GoTo cmdPrint_Click_Error

20        If GetSetting("Netacquire", "Histology", "Printer", "") <> "" Then
30            pPrintToPrinter = GetSetting("Netacquire", "Histology", "Printer", "")
40        End If

50        If pSampleID > 39999999 Then
60            Dept = "Y"
70        Else
80            Dept = "P"
90        End If

100       SaveOptionSetting "HistologyCopies", txtNoCopies

110       pTime = Format(Now, "dd/MMM/yyyy hh:mm:ss")
120       sql = "SELECT * FROM PrintPending WHERE " & _
                "Department = '" & Dept & "' " & _
                "AND SampleID = '" & pSampleID & "' " & _
                "AND pTime = '" & pTime & "'"
130       Set tb = New Recordset
140       RecOpenClient 0, tb, sql
150       If tb.EOF Then
160           tb.AddNew
170       End If
180       tb!SampleID = pSampleID
190       tb!Department = Dept
200       tb!Initiator = Username
210       tb!UsePrinter = pPrintToPrinter
220       tb!Hyear = Trim(lblYear)
230       tb!Clinician = lblClinician
240       tb!GP = lblGP
250       tb!pTime = pTime
260       tb!NoOfCopies = txtNoCopies
270       tb.Update

280       Exit Sub

cmdPrint_Click_Error:

          Dim strES As String
          Dim intEL As Integer

290       intEL = Erl
300       strES = Err.Description
310       LogError "frmHiCyReport", "cmdPrint_Click", intEL, strES

End Sub



Private Sub cmdPrint_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo cmdPrint_MouseDown_Error

20        If Button = 2 Then
30            Set frmForcePrinter.f = frmViewResults
40            frmForcePrinter.Show 1
50        End If

60        Exit Sub

cmdPrint_MouseDown_Error:

          Dim strES As String
          Dim intEL As Integer

70        intEL = Erl
80        strES = Err.Description
90        LogError "frmHiCyReport", "cmdPrint_MouseDown", intEL, strES

End Sub

Public Property Let SampleID(ByVal strNewValue As String)

10        pSampleID = strNewValue

End Property



