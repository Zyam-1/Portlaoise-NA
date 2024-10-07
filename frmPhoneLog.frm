VERSION 5.00
Begin VB.Form frmPhoneLog 
   Caption         =   "NetAcquire - Phone Log"
   ClientHeight    =   5850
   ClientLeft      =   285
   ClientTop       =   600
   ClientWidth     =   7335
   Icon            =   "frmPhoneLog.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5850
   ScaleWidth      =   7335
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   2100
      TabIndex        =   26
      Top             =   3465
      Width           =   3375
   End
   Begin VB.ComboBox cmbTitle 
      Height          =   315
      Left            =   1020
      TabIndex        =   25
      Text            =   "cmbTitle"
      Top             =   3450
      Width           =   1035
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   1000
      Left            =   1005
      Picture         =   "frmPhoneLog.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Save Changes"
      Top             =   4665
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   1000
      Left            =   5940
      Picture         =   "frmPhoneLog.frx":0316
      Style           =   1  'Graphical
      TabIndex        =   23
      Tag             =   "101"
      ToolTipText     =   "Exit Screen"
      Top             =   4680
      Width           =   1100
   End
   Begin VB.ComboBox cmbWard 
      Height          =   315
      Left            =   1005
      TabIndex        =   21
      Text            =   "cmbWard"
      Top             =   3825
      Width           =   3015
   End
   Begin VB.OptionButton optWard 
      Caption         =   "Wards"
      Height          =   195
      Left            =   2565
      TabIndex        =   20
      Top             =   2970
      Width           =   765
   End
   Begin VB.OptionButton optGP 
      Alignment       =   1  'Right Justify
      Caption         =   "GP's"
      Height          =   195
      Left            =   1905
      TabIndex        =   19
      Top             =   2970
      Value           =   -1  'True
      Width           =   645
   End
   Begin VB.CommandButton cmdHistory 
      Caption         =   "View Previous Details"
      Height          =   1000
      Left            =   2235
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4680
      Visible         =   0   'False
      Width           =   2000
   End
   Begin VB.Timer Timer1 
      Interval        =   30000
      Left            =   6480
      Top             =   1500
   End
   Begin VB.ComboBox cmbGP 
      Height          =   315
      Left            =   1005
      TabIndex        =   15
      Text            =   "cmbGP"
      Top             =   3840
      Width           =   3015
   End
   Begin VB.TextBox txtComment 
      Height          =   285
      Left            =   1005
      TabIndex        =   8
      Top             =   4260
      Width           =   6045
   End
   Begin VB.Frame Frame1 
      Caption         =   "Discipline"
      Height          =   2805
      Left            =   1080
      TabIndex        =   0
      Top             =   60
      Width           =   3015
      Begin VB.CheckBox chkDiscipline 
         Caption         =   "Endcrinology Results Phoned"
         Height          =   255
         Index           =   7
         Left            =   210
         TabIndex        =   22
         Top             =   2445
         Width           =   2415
      End
      Begin VB.CheckBox chkDiscipline 
         Caption         =   "Microbiology Results Phoned"
         Height          =   255
         Index           =   6
         Left            =   210
         TabIndex        =   18
         Top             =   2130
         Width           =   2415
      End
      Begin VB.CheckBox chkDiscipline 
         Caption         =   "External Results Phoned"
         Height          =   255
         Index           =   5
         Left            =   210
         TabIndex        =   6
         Top             =   1815
         Width           =   2085
      End
      Begin VB.CheckBox chkDiscipline 
         Caption         =   "Blood Gas Results Phoned"
         Height          =   255
         Index           =   4
         Left            =   210
         TabIndex        =   5
         Top             =   1500
         Width           =   2265
      End
      Begin VB.CheckBox chkDiscipline 
         Caption         =   "Immunology Results Phoned"
         Height          =   255
         Index           =   3
         Left            =   210
         TabIndex        =   4
         Top             =   1185
         Width           =   2355
      End
      Begin VB.CheckBox chkDiscipline 
         Caption         =   "Coagulation Results Phoned"
         Height          =   255
         Index           =   2
         Left            =   210
         TabIndex        =   3
         Top             =   870
         Width           =   2325
      End
      Begin VB.CheckBox chkDiscipline 
         Caption         =   "Biochemistry Results Phoned"
         Height          =   255
         Index           =   1
         Left            =   210
         TabIndex        =   2
         Top             =   555
         Width           =   2385
      End
      Begin VB.CheckBox chkDiscipline 
         Caption         =   "Haematology Results Phoned"
         Height          =   255
         Index           =   0
         Left            =   210
         TabIndex        =   1
         Top             =   240
         Width           =   2505
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Grade"
      Height          =   195
      Left            =   525
      TabIndex        =   27
      Top             =   3540
      Width           =   435
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Phone To"
      Height          =   195
      Left            =   255
      TabIndex        =   16
      Top             =   3900
      Width           =   705
   End
   Begin VB.Label lblDateTime 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5490
      TabIndex        =   14
      Top             =   1560
      Width           =   1635
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Date/Time"
      Height          =   195
      Left            =   4680
      TabIndex        =   13
      Top             =   1620
      Width           =   765
   End
   Begin VB.Label lblPhonedBy 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5490
      TabIndex        =   12
      Top             =   1050
      Width           =   1635
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Phoned By"
      Height          =   195
      Left            =   4665
      TabIndex        =   11
      Top             =   1095
      Width           =   780
   End
   Begin VB.Label lblSampleID 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   5490
      TabIndex        =   10
      Top             =   540
      Width           =   1635
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Sample ID"
      Height          =   195
      Left            =   4710
      TabIndex        =   9
      Top             =   570
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Comment"
      Height          =   195
      Left            =   300
      TabIndex        =   7
      Top             =   4290
      Width           =   660
   End
End
Attribute VB_Name = "frmPhoneLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pSampleID As String
Private pGP As String
Private pWardOrGp As String
Private m_Caller As String


Private Sub chkDiscipline_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo chkDiscipline_MouseUp_Error

20        cmdSave.Enabled = True

30        Exit Sub

chkDiscipline_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmPhoneLog", "chkDiscipline_MouseUp", intEL, strES

End Sub


Private Sub cmbGP_Click()

10        On Error GoTo cmbGP_Click_Error

20        cmdSave.Enabled = True

30        Exit Sub

cmbGP_Click_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmPhoneLog", "cmbGP_Click", intEL, strES

End Sub


Private Sub cmbTitle_Click()

10        On Error GoTo cmbTitle_Click_Error

20        cmdSave.Enabled = True

30        Exit Sub

cmbTitle_Click_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmPhoneLog", "cmbTitle_Click", intEL, strES

End Sub

Private Sub cmdCancel_Click()

10        Unload Me

End Sub


Private Sub cmdHistory_Click()

10        On Error GoTo cmdHistory_Click_Error

20        frmPhoneLogHistory.SampleID = pSampleID
30        frmPhoneLogHistory.Show 1
40        Exit Sub

cmdHistory_Click_Error:
          Dim strES As String
          Dim intEL As Integer
          
50        intEL = Erl
60        strES = Err.Description
70        LogError "frmPhoneLog", "cmdHistory_Click", intEL, strES

End Sub

Private Sub cmdSave_Click()

      Dim sql As String
      Dim n As Long
      Dim Disc As String
      Dim Obs As New Observations
      Dim Discipline As String
      Dim Title As String
      Dim Comment As String
      Dim PrintInitial As Boolean
      Dim PhonedTo As String

10    On Error GoTo cmdSave_Click_Error


20    PrintInitial = GetOptionSetting("PhoneLogPrintInitial", False)
30    Disc = ""
      'Check if at least one discipline is selected
40    For n = 0 To 7
50        If chkDiscipline(n).Value = 1 Then
60            Disc = Disc & Mid$("HBCIGXME", n + 1, 1)
70        Else
80            Disc = Disc & " "
90        End If
100   Next
110   If Trim$(Disc) = "" Then
120       iMsg "Select Discipline.", vbCritical
130       Exit Sub
140   End If

150   If Trim$(cmbGP) = "" And optGP Then
160       iMsg "Fill in 'Phone To'", vbCritical
170       Exit Sub
180   ElseIf Trim$(cmbWard) = "" And optWard Then
190       iMsg "Fill in 'Phone To'", vbCritical
200       Exit Sub
210   End If

      'check if GP or Ward
220   If optGP Then
230       PhonedTo = cmbGP.Text
240   Else
250       PhonedTo = cmbWard.Text
260   End If

270   sql = "INSERT into PhoneLog " & _
            "(DateTime,SampleID,PhonedTo, PhonedBy, Comment, Discipline, Title, PersonName) VALUES " & _
            "('" & Format$(Now, "yyyy/mm/dd hh:mm") & "', " & _
            "'" & pSampleID & "', " & _
            "'" & PhonedTo & "', " & _
            "'" & lblPhonedBy.Caption & "', " & _
            "'" & txtComment.Text & "', " & _
            "'" & Disc & "', " & _
            "'" & cmbTitle.Text & "', " & _
            "'" & txtName.Text & "')"
280   Cnxn(0).Execute sql

290   Title = Trim(cmbTitle & " " & txtName)
300   If Title <> "" Then Title = " (" & Title & ") "

310   Comment = txtComment & " phoned to " & IIf(optGP, cmbGP, cmbWard) & Title & " at " & _
                Format$(Now, "hh:mm") & " on " & Format$(Now, "dd/MM/yyyy") & _
                " by " & IIf(PrintInitial, UserCode, UserName) & "."
320   If chkDiscipline(0) Then
330       If InStr(frmEditAll.txtHaemComment, "phoned to") = 0 Then
340           If frmEditAll.txtHaemComment = "" Then
350               frmEditAll.txtHaemComment = Comment
360           Else
370               frmEditAll.txtHaemComment = frmEditAll.txtHaemComment & ". " & Comment
380           End If
390           frmEditAll.cmdSaveHaem.Enabled = False
400           Obs.Save pSampleID, False, "Haematology", Comment
410       End If
420   End If

430   If chkDiscipline(1) Then
440       If InStr(frmEditAll.txtBioComment, "phoned to") = 0 Then
450           If frmEditAll.txtBioComment = "" Then
460               frmEditAll.txtBioComment = Comment
470           Else
480               frmEditAll.txtBioComment = frmEditAll.txtBioComment & ". " & Comment
490           End If
500           frmEditAll.cmdSaveBio.Enabled = False
510           Obs.Save pSampleID, False, "Biochemistry", Comment
520       End If
530   End If

540   If chkDiscipline(2) Then
550       If InStr(frmEditAll.txtCoagComment, "phoned to") = 0 Then
560           If frmEditAll.txtCoagComment = "" Then
570               frmEditAll.txtCoagComment = Comment
580           Else
590               frmEditAll.txtCoagComment = frmEditAll.txtCoagComment & ". " & Comment
600           End If
610           frmEditAll.cmdSaveCoag.Enabled = False
620           Obs.Save pSampleID, False, "Coagulation", Comment
630       End If
640   End If

650   If chkDiscipline(3) Then
660       If InStr(frmEditAll.txtImmComment(1), "phoned to") = 0 Then
670           If frmEditAll.txtImmComment(1) = "" Then
680               frmEditAll.txtImmComment(1) = Comment
690           Else
700               frmEditAll.txtImmComment(1) = frmEditAll.txtImmComment(1) & ". " & Comment
710           End If
720           frmEditAll.cmdSaveImm(1).Enabled = False
730           Obs.Save pSampleID, False, "Immunology", Comment
740       End If

750   End If

760   If chkDiscipline(4) Then
770       If InStr(frmEditAll.txtBGaComment, "phoned to") = 0 Then
780           If frmEditAll.txtBGaComment = "" Then
790               frmEditAll.txtBGaComment = Comment
800           Else
810               frmEditAll.txtBGaComment = frmEditAll.txtBGaComment & ". " & Comment
820           End If
830           frmEditAll.cmdSaveBGa.Enabled = False
840           Obs.Save pSampleID, False, "BloodGas", Comment
850       End If
860   End If

870   If chkDiscipline(6) Then
880       If InStr(frmEditMicrobiologyNew.txtDemographicComment, "phoned to") = 0 Then
890           If frmEditMicrobiologyNew.txtDemographicComment = "" Then
900               frmEditMicrobiologyNew.txtDemographicComment = Comment
910           Else
920               frmEditMicrobiologyNew.txtDemographicComment = frmEditMicrobiologyNew.txtDemographicComment & ". " & Comment
930           End If
              'frmEditAll.cmdSaveImm(0).Enabled = False
940           Obs.Save pSampleID, False, "Demographic", Comment
950       End If
960   End If

970   If chkDiscipline(7) Then
980       If InStr(frmEditAll.txtImmComment(0), "phoned to") = 0 Then
990           If frmEditAll.txtImmComment(0) = "" Then
1000              frmEditAll.txtImmComment(0) = Comment
1010          Else
1020              frmEditAll.txtImmComment(0) = frmEditAll.txtImmComment(0) & ". " & Comment
1030          End If
1040          frmEditAll.cmdSaveImm(0).Enabled = False
1050          Obs.Save pSampleID, False, "Endocrinology", Comment
1060      End If

1070  End If

1080  Unload Me

1090  Exit Sub

cmdSave_Click_Error:

      Dim strES As String
      Dim intEL As Integer

1100  intEL = Erl
1110  strES = Err.Description
1120  LogError "frmPhoneLog", "cmdsave_Click", intEL, strES, sql

End Sub

Private Sub Form_Activate()


10        On Error GoTo Form_Activate_Error

20        If Caller = "Micro" Then
30            lblSampleID = pSampleID - SysOptMicroOffset(0)
40        Else
50            lblSampleID = pSampleID
60        End If
70        lblPhonedBy = UserName
80        lblDateTime = Format$(Now, "dd/mmm/yyyy hh:mm")
90        If pWardOrGp = "GP" Then
100           cmbGP = pGP
110           cmbGP.Visible = True
120           optGP.Value = True
130           cmbWard.Visible = False
140       Else
150           cmbWard = pGP
160           cmbGP.Visible = False
170           optWard.Value = True
180           cmbWard.Visible = True
190       End If
200       txtComment = ""
210       If CheckPhoneLog(pSampleID) Then
220           cmdHistory.Visible = True
230       Else
240           cmdHistory.Visible = False
250       End If
260       Select Case Caller
          Case "General"
270           chkDiscipline(0).Enabled = True
280           chkDiscipline(1).Enabled = True
290           chkDiscipline(2).Enabled = True
300           chkDiscipline(3).Enabled = True
310           chkDiscipline(4).Enabled = True
320           chkDiscipline(5).Enabled = False
330           chkDiscipline(6).Enabled = False
340           chkDiscipline(7).Enabled = True

350       Case "Micro"
360           chkDiscipline(0).Enabled = False
370           chkDiscipline(1).Enabled = False
380           chkDiscipline(2).Enabled = False
390           chkDiscipline(3).Enabled = False
400           chkDiscipline(4).Enabled = False
410           chkDiscipline(5).Enabled = False
420           chkDiscipline(6).Enabled = True
430           chkDiscipline(7).Enabled = False
440       Case "External"
450           chkDiscipline(0).Enabled = False
460           chkDiscipline(1).Enabled = False
470           chkDiscipline(2).Enabled = False
480           chkDiscipline(3).Enabled = False
490           chkDiscipline(4).Enabled = False
500           chkDiscipline(5).Enabled = True
510           chkDiscipline(6).Enabled = False
520           chkDiscipline(7).Enabled = False
530       Case Else
540           chkDiscipline(0).Enabled = True
550           chkDiscipline(1).Enabled = True
560           chkDiscipline(2).Enabled = True
570           chkDiscipline(3).Enabled = True
580           chkDiscipline(4).Enabled = True
590           chkDiscipline(5).Enabled = True
600           chkDiscipline(6).Enabled = True
610           chkDiscipline(7).Enabled = True
620       End Select
630       Exit Sub

Form_Activate_Error:

          Dim strES As String
          Dim intEL As Integer

640       intEL = Erl
650       strES = Err.Description
660       LogError "frmPhoneLog", "Form_Activate", intEL, strES

End Sub

Private Sub Form_Load()

          Dim n As Long

10        On Error GoTo Form_Load_Error

20        If pWardOrGp = "GP" Then
30            FillGPsWard Me, HospName(0)
40            cmbWard.Visible = False
50            cmbGP.Visible = True
60        Else
70            FillGPsWard Me, HospName(0)
80            cmbWard = ""
90            cmbWard.Visible = True
100           cmbGP.Visible = False
110       End If
120       For n = 0 To 6
130           chkDiscipline(n).Value = 0
140       Next
150       FillGenericList cmbTitle, "PersonTitles", True

160       Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

170       intEL = Erl
180       strES = Err.Description
190       LogError "frmPhoneLog", "Form_Load", intEL, strES

End Sub



Public Property Let SampleID(ByVal strNewValue As String)

10        On Error GoTo SampleID_Error

20        pSampleID = strNewValue

30        Exit Property

SampleID_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmPhoneLog", "SampleID", intEL, strES

End Property

Public Property Let GP(ByVal strNewValue As String)

10        On Error GoTo GP_Error

20        pGP = strNewValue

30        Exit Property

GP_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmPhoneLog", "GP", intEL, strES

End Property

Public Property Let WardOrGP(ByVal strNewValue As String)

10        On Error GoTo WardOrGP_Error

20        pWardOrGp = strNewValue

30        Exit Property

WardOrGP_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmPhoneLog", "WardOrGP", intEL, strES

End Property


Private Sub optGP_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo optGP_MouseUp_Error

20        FillGPsWard Me, HospName(0)
30        cmbWard.Visible = False
40        cmbGP.Visible = True

50        Exit Sub

optGP_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmPhoneLog", "optGP_MouseUp", intEL, strES

End Sub


Private Sub optWard_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo optWard_MouseUp_Error

20        FillGPsWard Me, HospName(0)
30        cmbWard = ""
40        cmbWard.Visible = True
50        cmbGP.Visible = False

60        Exit Sub

optWard_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer

70        intEL = Erl
80        strES = Err.Description
90        LogError "frmPhoneLog", "optWard_MouseUp", intEL, strES

End Sub


Private Sub Timer1_Timer()

10        On Error GoTo Timer1_Timer_Error

20        lblDateTime = Format$(Now, "dd/mmm/yyyy hh:mm")

30        Exit Sub

Timer1_Timer_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmPhoneLog", "Timer1_Timer", intEL, strES

End Sub

Private Sub txtComment_KeyPress(KeyAscii As Integer)

10        On Error GoTo txtComment_KeyPress_Error

20        cmdSave.Enabled = True

30        Exit Sub

txtComment_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmPhoneLog", "txtComment_KeyPress", intEL, strES

End Sub


Private Sub txtName_Change()

10        On Error GoTo txtName_Change_Error

20        cmdSave.Enabled = True

30        Exit Sub

txtName_Change_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmPhoneLog", "txtName_Change", intEL, strES

End Sub



Public Property Get Caller() As String

10        Caller = m_Caller

End Property

Public Property Let Caller(ByVal sCaller As String)

10        m_Caller = sCaller

End Property
