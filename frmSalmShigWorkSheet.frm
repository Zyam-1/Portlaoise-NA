VERSION 5.00
Begin VB.Form frmSalmShigWorkSheet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Salmonella/Shigella"
   ClientHeight    =   4305
   ClientLeft      =   210
   ClientTop       =   480
   ClientWidth     =   12210
   Icon            =   "frmSalmShigWorkSheet.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   12210
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      Height          =   1275
      Index           =   1
      Left            =   240
      TabIndex        =   56
      Top             =   120
      Width           =   7275
      Begin VB.Label lblSex 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5940
         TabIndex        =   66
         Top             =   330
         Width           =   705
      End
      Begin VB.Label lblAge 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4860
         TabIndex        =   65
         Top             =   330
         Width           =   585
      End
      Begin VB.Label lblDoB 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2850
         TabIndex        =   64
         Top             =   330
         Width           =   1515
      End
      Begin VB.Label lblName 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   690
         TabIndex        =   63
         Top             =   780
         Width           =   5955
      End
      Begin VB.Label lblChart 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   660
         TabIndex        =   62
         Top             =   300
         Width           =   1485
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Chart #"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   61
         Top             =   330
         Width           =   525
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   60
         Top             =   810
         Width           =   420
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "D.o.B"
         Height          =   195
         Index           =   0
         Left            =   2370
         TabIndex        =   59
         Top             =   360
         Width           =   405
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Age"
         Height          =   195
         Index           =   0
         Left            =   4530
         TabIndex        =   58
         Top             =   360
         Width           =   285
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Sex"
         Height          =   195
         Index           =   0
         Left            =   5610
         TabIndex        =   57
         Top             =   360
         Width           =   270
      End
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   735
      Left            =   10260
      Picture         =   "frmSalmShigWorkSheet.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   480
      Width           =   1155
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   735
      Left            =   8880
      Picture         =   "frmSalmShigWorkSheet.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   480
      Width           =   1155
   End
   Begin VB.Frame Frame3 
      Caption         =   "Salmonella"
      Height          =   2595
      Left            =   240
      TabIndex        =   24
      Top             =   1500
      Width           =   7275
      Begin VB.CheckBox chkColindale 
         Alignment       =   1  'Right Justify
         Caption         =   "Colindale"
         Height          =   195
         Left            =   3540
         TabIndex        =   39
         Top             =   2040
         Width           =   945
      End
      Begin VB.CheckBox chkB17 
         Caption         =   "Check4"
         Height          =   255
         Left            =   4590
         TabIndex        =   38
         Top             =   1560
         Width           =   255
      End
      Begin VB.CheckBox chkB15 
         Caption         =   "Check3"
         Height          =   255
         Left            =   4290
         TabIndex        =   37
         Top             =   1560
         Width           =   255
      End
      Begin VB.CheckBox chkB16 
         Caption         =   "Check2"
         Height          =   255
         Left            =   4590
         TabIndex        =   36
         Top             =   1290
         Width           =   255
      End
      Begin VB.CheckBox chkB12 
         Caption         =   "Check1"
         Height          =   255
         Left            =   4290
         TabIndex        =   35
         Top             =   1290
         Width           =   255
      End
      Begin VB.CheckBox chkOther 
         Caption         =   "Check1"
         Height          =   195
         Left            =   5670
         TabIndex        =   34
         Top             =   870
         Width           =   225
      End
      Begin VB.CheckBox chkLittleI 
         Caption         =   "Check1"
         Height          =   225
         Left            =   5310
         TabIndex        =   33
         Top             =   840
         Width           =   225
      End
      Begin VB.CommandButton bSens 
         Caption         =   "Sensitivity"
         Height          =   315
         Index           =   0
         Left            =   5970
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   1500
         Width           =   1125
      End
      Begin VB.CheckBox chkR 
         Alignment       =   1  'Right Justify
         Caption         =   "Rapid 1"
         Height          =   195
         Index           =   1
         Left            =   3930
         TabIndex        =   31
         Top             =   450
         Width           =   855
      End
      Begin VB.CheckBox chkR 
         Alignment       =   1  'Right Justify
         Caption         =   "Rapid 3"
         Height          =   195
         Index           =   3
         Left            =   3930
         TabIndex        =   30
         Top             =   870
         Width           =   855
      End
      Begin VB.CheckBox chkR 
         Alignment       =   1  'Right Justify
         Caption         =   "Rapid 2"
         Height          =   195
         Index           =   2
         Left            =   3930
         TabIndex        =   29
         Top             =   660
         Width           =   855
      End
      Begin VB.TextBox txtHAntigen 
         Height          =   285
         Left            =   5190
         TabIndex        =   28
         Top             =   450
         Width           =   1245
      End
      Begin VB.TextBox txtSalmType 
         Height          =   285
         Left            =   2040
         TabIndex        =   27
         Top             =   660
         Width           =   915
      End
      Begin VB.TextBox txtSalmID 
         BackColor       =   &H0000FFFF&
         Height          =   285
         Left            =   300
         TabIndex        =   26
         Top             =   2040
         Width           =   2655
      End
      Begin VB.TextBox txtColindaleResult 
         BackColor       =   &H0000FFFF&
         Height          =   285
         Left            =   4530
         MaxLength       =   40
         TabIndex        =   25
         Top             =   2010
         Width           =   2625
      End
      Begin VB.Label lblPolyH2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1680
         TabIndex        =   53
         Top             =   1590
         Width           =   1275
      End
      Begin VB.Label lblPolyH 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1680
         TabIndex        =   52
         Top             =   1200
         Width           =   1275
      End
      Begin VB.Label lblPolyO 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Positive"
         Height          =   255
         Left            =   2040
         TabIndex        =   51
         Top             =   360
         Width           =   915
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "H Antigen"
         Height          =   195
         Index           =   0
         Left            =   5220
         TabIndex        =   50
         Top             =   240
         Width           =   705
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "1,7"
         Height          =   195
         Left            =   4860
         TabIndex        =   49
         Top             =   1590
         Width           =   225
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "1,6"
         Height          =   195
         Left            =   4860
         TabIndex        =   48
         Top             =   1320
         Width           =   225
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "1,5"
         Height          =   195
         Left            =   3960
         TabIndex        =   47
         Top             =   1590
         Width           =   225
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "1,2"
         Height          =   195
         Left            =   3960
         TabIndex        =   46
         Top             =   1320
         Width           =   225
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Poly H Phase 2"
         Height          =   195
         Index           =   1
         Left            =   540
         TabIndex        =   45
         Top             =   1590
         Width           =   1095
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Poly H Phase 1 && 2"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   44
         Top             =   1230
         Width           =   1395
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Polyvalent-O Groups A-S"
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   43
         Top             =   390
         Width           =   1785
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Other"
         Height          =   195
         Left            =   5910
         TabIndex        =   42
         Top             =   870
         Width           =   390
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "i"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   5190
         TabIndex        =   41
         Top             =   840
         Width           =   60
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Type"
         Height          =   195
         Index           =   0
         Left            =   1590
         TabIndex        =   40
         Top             =   690
         Width           =   360
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Shigella"
      Height          =   2595
      Index           =   0
      Left            =   7590
      TabIndex        =   0
      Top             =   1500
      Width           =   4485
      Begin VB.CheckBox chkSonn 
         Alignment       =   1  'Right Justify
         Caption         =   "Phase 2"
         Height          =   195
         Index           =   1
         Left            =   3390
         TabIndex        =   17
         Top             =   1710
         Width           =   885
      End
      Begin VB.CheckBox chkSonn 
         Alignment       =   1  'Right Justify
         Caption         =   "Phase 1"
         Height          =   195
         Index           =   0
         Left            =   3390
         TabIndex        =   16
         Top             =   1470
         Width           =   885
      End
      Begin VB.CommandButton bSens 
         Caption         =   "Sensitivity"
         Height          =   315
         Index           =   1
         Left            =   1680
         TabIndex        =   15
         Top             =   1500
         Width           =   975
      End
      Begin VB.CheckBox chkFlex 
         Alignment       =   1  'Right Justify
         Caption         =   "X"
         Height          =   225
         Index           =   7
         Left            =   240
         TabIndex        =   14
         Top             =   1500
         Width           =   405
      End
      Begin VB.CheckBox chkFlex 
         Caption         =   "Y"
         Height          =   225
         Index           =   8
         Left            =   720
         TabIndex        =   13
         Top             =   1500
         Width           =   465
      End
      Begin VB.CheckBox chkFlex 
         Caption         =   "6"
         Height          =   225
         Index           =   6
         Left            =   720
         TabIndex        =   12
         Top             =   1230
         Width           =   465
      End
      Begin VB.CheckBox chkFlex 
         Caption         =   "5"
         Height          =   225
         Index           =   5
         Left            =   720
         TabIndex        =   11
         Top             =   960
         Width           =   465
      End
      Begin VB.CheckBox chkFlex 
         Caption         =   "4"
         Height          =   225
         Index           =   4
         Left            =   720
         TabIndex        =   10
         Top             =   690
         Width           =   465
      End
      Begin VB.CheckBox chkFlex 
         Alignment       =   1  'Right Justify
         Caption         =   "3"
         Height          =   225
         Index           =   3
         Left            =   240
         TabIndex        =   9
         Top             =   1230
         Width           =   405
      End
      Begin VB.CheckBox chkFlex 
         Alignment       =   1  'Right Justify
         Caption         =   "2"
         Height          =   225
         Index           =   2
         Left            =   240
         TabIndex        =   8
         Top             =   960
         Width           =   405
      End
      Begin VB.CheckBox chkFlex 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   225
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   690
         Width           =   405
      End
      Begin VB.CheckBox chkFlex 
         Alignment       =   1  'Right Justify
         Caption         =   "Poly"
         Height          =   225
         Index           =   0
         Left            =   300
         TabIndex        =   6
         Top             =   420
         Width           =   615
      End
      Begin VB.CheckBox chkDys 
         Alignment       =   1  'Right Justify
         Caption         =   "Poly 3-10"
         Height          =   195
         Index           =   1
         Left            =   3300
         TabIndex        =   5
         Top             =   900
         Width           =   975
      End
      Begin VB.CheckBox chkDys 
         Alignment       =   1  'Right Justify
         Caption         =   "Poly 1-10"
         Height          =   255
         Index           =   0
         Left            =   3300
         TabIndex        =   4
         Top             =   600
         Width           =   975
      End
      Begin VB.CheckBox chkBoy 
         Alignment       =   1  'Right Justify
         Caption         =   "Poly 12-15"
         Height          =   195
         Index           =   2
         Left            =   1590
         TabIndex        =   3
         Top             =   1110
         Width           =   1065
      End
      Begin VB.CheckBox chkBoy 
         Alignment       =   1  'Right Justify
         Caption         =   "Poly 7-11"
         Height          =   195
         Index           =   1
         Left            =   1680
         TabIndex        =   2
         Top             =   870
         Width           =   975
      End
      Begin VB.CheckBox chkBoy 
         Alignment       =   1  'Right Justify
         Caption         =   "Poly 1-6"
         Height          =   195
         Index           =   0
         Left            =   1770
         TabIndex        =   1
         Top             =   630
         Width           =   885
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "S Dysenteriae"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   3090
         TabIndex        =   23
         Top             =   390
         Width           =   1200
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "S Sonnei"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   3480
         TabIndex        =   22
         Top             =   1260
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "S Boydii"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   1920
         TabIndex        =   21
         Top             =   390
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "S Flexneri"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   20
         Top             =   210
         Width           =   855
      End
      Begin VB.Label lblShigellaType 
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   270
         TabIndex        =   19
         Top             =   2040
         Width           =   4005
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Type"
         Height          =   195
         Index           =   1
         Left            =   270
         TabIndex        =   18
         Top             =   1800
         Width           =   435
      End
   End
   Begin VB.Label lblSampleID 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   9360
      TabIndex        =   68
      Top             =   60
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Sample ID"
      Height          =   195
      Left            =   8580
      TabIndex        =   67
      Top             =   90
      Width           =   735
   End
End
Attribute VB_Name = "frmSalmShigWorkSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub SaveSalmShig()

      Dim tb As Recordset
      Dim sql As String
      Dim n As Long
      Dim Counter As Long

10    On Error GoTo SaveSalmShig_Error

20    sql = "SELECT * from SalmShig WHERE " & _
            "SampleID = '" & SysOptMicroOffset(0) + lblSampleID & "'"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    If tb.EOF Then tb.AddNew

60    tb!SampleID = SysOptMicroOffset(0) + lblSampleID

70    tb!PolyO = Trim$(Left(lblPolyO & " ", 1))
80    tb!PolyH = Trim$(Left(lblPolyH & " ", 1))
90    tb!PolyH2 = Trim$(Left(lblPolyH2 & " ", 1))

100   tb!SalmType = txtSalmType


110   Counter = 0
120   If chkR(1) Then Counter = 1
130   If chkR(2) Then Counter = Counter + 2
140   If chkR(3) Then Counter = Counter + 4
150   tb!Rapid = Counter

160   tb!LittleI = chkLittleI
170   tb!Other = chkOther

180   Counter = 0
190   If chkB12 Then Counter = 1
200   If chkB15 Then Counter = Counter + 2
210   If chkB16 Then Counter = Counter + 4
220   If chkB17 Then Counter = Counter + 8
230   tb!b12 = Counter

240   tb!SalmIdent = txtSalmID
250   tb!Colindale = chkColindale
260   tb!ColindaleResult = txtColindaleResult

270   tb!ShigType = lblShigellaType

280   Counter = 0
290   For n = 0 To 8
300     If chkFlex(n) Then Counter = Counter + 2 ^ n
310   Next
320   tb!Flex = Counter

330   Counter = 0
340   For n = 0 To 2
350     If chkBoy(n) Then Counter = Counter + 2 ^ n
360   Next
370   tb!Boy = Counter

380   Counter = 0
390   For n = 0 To 1
400     If chkDys(n) Then Counter = Counter + 2 ^ n
410   Next
420   tb!Dys = Counter

430   Counter = 0
440   For n = 0 To 1
450     If chkSonn(n) Then Counter = Counter + 2 ^ n
460   Next
470   tb!Sonn = Counter

480   tb.Update

490   Exit Sub

SaveSalmShig_Error:

      Dim strES As String
      Dim intEL As Integer

500   Screen.MousePointer = 0

510   intEL = Erl
520   strES = Err.Description
530   LogError "frmSalmShigWorkSheet", "SaveSalmShig", intEL, strES, sql


End Sub

Private Sub chkB12_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10    On Error GoTo chkB12_MouseUp_Error

20    cmdSave.Enabled = True

30    Exit Sub

chkB12_MouseUp_Error:

      Dim strES As String
      Dim intEL As Integer

40    Screen.MousePointer = 0

50    intEL = Erl
60    strES = Err.Description
70    LogError "frmSalmShigWorkSheet", "chkB12_MouseUp", intEL, strES


End Sub


Private Sub chkB15_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10    On Error GoTo chkB15_MouseUp_Error

20    cmdSave.Enabled = True

30    Exit Sub

chkB15_MouseUp_Error:

      Dim strES As String
      Dim intEL As Integer

40    Screen.MousePointer = 0

50    intEL = Erl
60    strES = Err.Description
70    LogError "frmSalmShigWorkSheet", "chkB15_MouseUp", intEL, strES


End Sub


Private Sub chkB16_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10    On Error GoTo chkB16_MouseUp_Error

20    cmdSave.Enabled = True

30    Exit Sub

chkB16_MouseUp_Error:

      Dim strES As String
      Dim intEL As Integer

40    Screen.MousePointer = 0

50    intEL = Erl
60    strES = Err.Description
70    LogError "frmSalmShigWorkSheet", "chkB16_MouseUp", intEL, strES


End Sub


Private Sub chkB17_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10    On Error GoTo chkB17_MouseUp_Error

20    cmdSave.Enabled = True

30    Exit Sub

chkB17_MouseUp_Error:

      Dim strES As String
      Dim intEL As Integer

40    Screen.MousePointer = 0

50    intEL = Erl
60    strES = Err.Description
70    LogError "frmSalmShigWorkSheet", "chkB17_MouseUp", intEL, strES


End Sub


Private Sub chkColindale_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10    On Error GoTo chkColindale_MouseUp_Error

20    cmdSave.Enabled = True

30    Exit Sub

chkColindale_MouseUp_Error:

      Dim strES As String
      Dim intEL As Integer

40    Screen.MousePointer = 0

50    intEL = Erl
60    strES = Err.Description
70    LogError "frmSalmShigWorkSheet", "chkColindale_MouseUp", intEL, strES


End Sub


Private Sub chkLittleI_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10    On Error GoTo chkLittleI_MouseUp_Error

20    cmdSave.Enabled = True

30    Exit Sub

chkLittleI_MouseUp_Error:

      Dim strES As String
      Dim intEL As Integer

40    Screen.MousePointer = 0

50    intEL = Erl
60    strES = Err.Description
70    LogError "frmSalmShigWorkSheet", "chkLittleI_MouseUp", intEL, strES


End Sub


Private Sub chkOther_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10    On Error GoTo chkOther_MouseUp_Error

20    cmdSave.Enabled = True

30    Exit Sub

chkOther_MouseUp_Error:

      Dim strES As String
      Dim intEL As Integer

40    Screen.MousePointer = 0

50    intEL = Erl
60    strES = Err.Description
70    LogError "frmSalmShigWorkSheet", "chkOther_MouseUp", intEL, strES


End Sub


Private Sub chkR_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

      Dim Pattern As String

10    On Error GoTo chkR_MouseUp_Error

20    Pattern = IIf(chkR(1) = 1, "+", "-") & _
                IIf(chkR(2) = 1, "+", "-") & _
                IIf(chkR(3) = 1, "+", "-")

30    Select Case Pattern
        Case "++-": txtHAntigen = "b"
40      Case "+-+": txtHAntigen = "d"
50      Case "+++": txtHAntigen = "E Complex"
60      Case "--+": txtHAntigen = "G Complex"
70      Case "-++": txtHAntigen = "k"
80      Case "-+-": txtHAntigen = "L Complex"
90      Case "+--": txtHAntigen = "r"
100     Case "---": txtHAntigen = ""
110   End Select

120   cmdSave.Enabled = True

130   Exit Sub

chkR_MouseUp_Error:

      Dim strES As String
      Dim intEL As Integer

140   Screen.MousePointer = 0

150   intEL = Erl
160   strES = Err.Description
170   LogError "frmSalmShigWorkSheet", "chkR_MouseUp", intEL, strES


End Sub

Private Sub cmdCancel_Click()

10    On Error GoTo cmdCancel_Click_Error

20    If cmdSave.Enabled Then
30      If iMsg("Cancel without Saving?", vbQuestion + vbYesNo) = vbNo Then
40        Exit Sub
50      End If
60    End If

70    Unload Me

80    Exit Sub

cmdCancel_Click_Error:

      Dim strES As String
      Dim intEL As Integer

90    Screen.MousePointer = 0

100   intEL = Erl
110   strES = Err.Description
120   LogError "frmSalmShigWorkSheet", "cmdCancel_Click", intEL, strES


End Sub

Private Sub chkBoy_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

      Dim n As Long

10    On Error GoTo chkBoy_MouseUp_Error

20    lblShigellaType = ""

30    For n = 0 To 2
40      If chkBoy(n) = 1 Then lblShigellaType = "Shigella Boydii"
50    Next

60    cmdSave.Enabled = True

70    Exit Sub

chkBoy_MouseUp_Error:

      Dim strES As String
      Dim intEL As Integer

80    Screen.MousePointer = 0

90    intEL = Erl
100   strES = Err.Description
110   LogError "frmSalmShigWorkSheet", "chkBoy_MouseUp", intEL, strES


End Sub


Private Sub chkFlex_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

      Dim n As Long
      Dim Found As Long

10    On Error GoTo chkFlex_MouseUp_Error

20    If Index > 0 And Index < 7 Then
30      Found = Index
40      For n = 1 To 6
50        chkFlex(n) = 0
60      Next
70      chkFlex(Found) = 1
80    End If

90    lblShigellaType = ""
100   Found = False

110   If chkFlex(0) = 1 Then
120     For n = 1 To 6
130       If chkFlex(n) = 1 Then
140         lblShigellaType = "Shigella Flexneri Type " & n
150         Found = True
160       End If
170     Next
180     If Found Then
190       If chkFlex(7) = 1 And chkFlex(8) = 0 Then
200         lblShigellaType = lblShigellaType & " Variant X"
210       ElseIf chkFlex(8) = 1 And chkFlex(7) = 0 Then
220         lblShigellaType = lblShigellaType & " Variant Y"
230       End If
240     Else
250       lblShigellaType = ""
260     End If
270   End If

280   cmdSave.Enabled = True

290   Exit Sub

chkFlex_MouseUp_Error:

      Dim strES As String
      Dim intEL As Integer

300   Screen.MousePointer = 0

310   intEL = Erl
320   strES = Err.Description
330   LogError "frmSalmShigWorkSheet", "chkFlex_MouseUp", intEL, strES


End Sub


Private Sub chkdys_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

      Dim n As Long

10    On Error GoTo chkdys_MouseUp_Error

20    lblShigellaType = ""

30    For n = 0 To 1
40      If chkDys(n) = 1 Then lblShigellaType = "Shigella Dysenteriae"
50    Next

60    cmdSave.Enabled = True

70    Exit Sub

chkdys_MouseUp_Error:

      Dim strES As String
      Dim intEL As Integer

80    Screen.MousePointer = 0

90    intEL = Erl
100   strES = Err.Description
110   LogError "frmSalmShigWorkSheet", "chkdys_MouseUp", intEL, strES


End Sub


Private Sub LoadSalmShig()

      Dim tb As Recordset
      Dim sql As String
      Dim n As Long

10    On Error GoTo LoadSalmShig_Error

20    ClearSalmShig

30    sql = "SELECT * from SalmShig WHERE " & _
            "SampleID = '" & SysOptMicroOffset(0) + lblSampleID & "'"
40    Set tb = New Recordset
50    RecOpenServer 0, tb, sql
60    If Not tb.EOF Then
70      If tb!PolyO & "" = "N" Then
80        lblPolyO = "Negative"
90        lblPolyO.BackColor = vbGreen
100     ElseIf tb!PolyO & "" = "P" Then
110       lblPolyO = "Positive"
120       lblPolyO.BackColor = vbRed
130     End If
140     txtSalmType = tb!SalmType & ""
150     If tb!PolyH & "" = "N" Then
160       lblPolyH = "Negative"
170       lblPolyH.BackColor = vbGreen
180     ElseIf tb!PolyH & "" = "P" Then
190       lblPolyH = "Positive"
200       lblPolyH.BackColor = vbRed
210     ElseIf tb!PolyH & "" = "I" Then
220       lblPolyH = "Indeterminate"
230       lblPolyH.BackColor = vbYellow
240     End If
250     If tb!PolyH2 & "" = "N" Then
260       lblPolyH2 = "Negative"
270       lblPolyH2.BackColor = vbGreen
280     ElseIf tb!PolyH2 & "" = "P" Then
290       lblPolyH2 = "Positive"
300       lblPolyH2.BackColor = vbRed
310     ElseIf tb!PolyH2 & "" = "I" Then
320       lblPolyH2 = "Indeterminate"
330       lblPolyH2.BackColor = vbYellow
340     End If
  
350     If Not IsNull(tb!Rapid) Then
360       chkR(1) = IIf(tb!Rapid And 1, 1, 0)
370       chkR(2) = IIf(tb!Rapid And 2, 1, 0)
380       chkR(3) = IIf(tb!Rapid And 4, 1, 0)
390     End If
  
400     chkLittleI = IIf(tb!LittleI, 1, 0)
410     chkOther = IIf(tb!Other, 1, 0)
420     If Not IsNull(tb!b12) Then
430       chkB12 = IIf(tb!b12 And 1, 1, 0)
440       chkB15 = IIf(tb!b12 And 2, 1, 0)
450       chkB16 = IIf(tb!b12 And 4, 1, 0)
460       chkB17 = IIf(tb!b12 And 8, 1, 0)
470     End If
480     txtSalmID = tb!SalmIdent & ""
490     chkColindale = IIf(tb!Colindale, 1, 0)
500     txtColindaleResult = tb!ColindaleResult & ""

510     lblShigellaType = tb!ShigType & ""
520     For n = 0 To 8
530       chkFlex(n) = IIf(tb!Flex And 2 ^ n, 1, 0)
540     Next
550     For n = 0 To 2
560       chkBoy(n) = IIf(tb!Boy And 2 ^ n, 1, 0)
570     Next
580     For n = 0 To 1
590       chkDys(n) = IIf(tb!Dys And 2 ^ n, 1, 0)
600       chkSonn(n) = IIf(tb!Sonn And 2 ^ n, 1, 0)
610     Next
620   End If

630   Exit Sub

LoadSalmShig_Error:

      Dim strES As String
      Dim intEL As Integer

640   Screen.MousePointer = 0

650   intEL = Erl
660   strES = Err.Description
670   LogError "frmSalmShigWorkSheet", "LoadSalmShig", intEL, strES, sql


End Sub

Private Sub ClearSalmShig()
  
      Dim n As Long
  
10    On Error GoTo ClearSalmShig_Error

20    lblPolyO = ""
30    lblPolyO.BackColor = &H8000000F

40    txtSalmType = ""
50    lblPolyH = ""
60    lblPolyH.BackColor = &H8000000F

70    chkR(1) = False
80    chkR(2) = False
90    chkR(3) = False
100   chkLittleI = False
110   chkOther = False
120   chkB12 = 0
130   chkB15 = 0
140   chkB16 = 0
150   chkB17 = 0
160   chkColindale = False
170   txtColindaleResult = ""
180   txtSalmID = ""

190   lblShigellaType = ""
200   For n = 0 To 8
210     chkFlex(n) = False
220   Next
230   For n = 0 To 2
240     chkBoy(n) = False
250   Next
260   For n = 0 To 1
270     chkDys(n) = False
280     chkSonn(n) = False
290   Next

300   Exit Sub

ClearSalmShig_Error:

      Dim strES As String
      Dim intEL As Integer

310   Screen.MousePointer = 0

320   intEL = Erl
330   strES = Err.Description
340   LogError "frmSalmShigWorkSheet", "ClearSalmShig", intEL, strES


End Sub

Private Sub cmdsave_Click()

10    On Error GoTo cmdsave_Click_Error

20    SaveSalmShig

30    Unload Me

40    Exit Sub

cmdsave_Click_Error:

      Dim strES As String
      Dim intEL As Integer

50    Screen.MousePointer = 0

60    intEL = Erl
70    strES = Err.Description
80    LogError "frmSalmShigWorkSheet", "cmdsave_Click", intEL, strES


End Sub

Private Sub Form_Activate()

10    On Error GoTo Form_Activate_Error

20    LoadSalmShig

30    Exit Sub

Form_Activate_Error:

      Dim strES As String
      Dim intEL As Integer

40    Screen.MousePointer = 0

50    intEL = Erl
60    strES = Err.Description
70    LogError "frmSalmShigWorkSheet", "Form_Activate", intEL, strES


End Sub

Private Sub Form_Unload(Cancel As Integer)

10    On Error GoTo Form_Unload_Error

20    With frmEditMicrobiologyNew
30      .lblSalmonella = txtSalmID
40      .lblShigella = lblShigellaType
50      .lblColindale = txtColindaleResult
60    End With

70    Exit Sub

Form_Unload_Error:

      Dim strES As String
      Dim intEL As Integer

80    Screen.MousePointer = 0

90    intEL = Erl
100   strES = Err.Description
110   LogError "frmSalmShigWorkSheet", "Form_Unload", intEL, strES


End Sub


Private Sub lblPolyO_Click()

10    On Error GoTo lblPolyO_Click_Error

20    With lblPolyO
30      Select Case .Caption
        Case ""
40        .Caption = "Negative"
50        .BackColor = vbGreen
60      Case "Negative"
70        .Caption = "Positive"
80        .BackColor = vbRed
90      Case "Positive"
100       .Caption = ""
110       .BackColor = &H8000000F
120     End Select
130   End With

140   cmdSave.Enabled = True

150   Exit Sub

lblPolyO_Click_Error:

      Dim strES As String
      Dim intEL As Integer

160   Screen.MousePointer = 0

170   intEL = Erl
180   strES = Err.Description
190   LogError "frmSalmShigWorkSheet", "lblPolyO_Click", intEL, strES


End Sub


Private Sub lblPolyH_Click()

10    On Error GoTo lblPolyH_Click_Error

20    With lblPolyH
30      Select Case .Caption
        Case ""
40        .Caption = "Negative"
50        .BackColor = vbGreen
60      Case "Negative"
70        .Caption = "Positive"
80        .BackColor = vbRed
90      Case "Positive"
100       .Caption = "Indeterminate"
110       .BackColor = vbYellow
120     Case "Indeterminate"
130       .Caption = ""
140       .BackColor = &H8000000F
150     End Select
160   End With

170   cmdSave.Enabled = True

180   Exit Sub

lblPolyH_Click_Error:

      Dim strES As String
      Dim intEL As Integer

190   Screen.MousePointer = 0

200   intEL = Erl
210   strES = Err.Description
220   LogError "frmSalmShigWorkSheet", "lblPolyH_Click", intEL, strES


End Sub



Private Sub lblPolyH2_Click()

10    On Error GoTo lblPolyH2_Click_Error

20    With lblPolyH2
30      Select Case .Caption
        Case ""
40        .Caption = "Negative"
50        .BackColor = vbGreen
60      Case "Negative"
70        .Caption = "Positive"
80        .BackColor = vbRed
90      Case "Positive"
100       .Caption = "Indeterminate"
110       .BackColor = vbYellow
120     Case "Indeterminate"
130       .Caption = ""
140       .BackColor = &H8000000F
150     End Select
160   End With

170   cmdSave.Enabled = True

180   Exit Sub

lblPolyH2_Click_Error:

      Dim strES As String
      Dim intEL As Integer

190   Screen.MousePointer = 0

200   intEL = Erl
210   strES = Err.Description
220   LogError "frmSalmShigWorkSheet", "lblPolyH2_Click", intEL, strES


End Sub


Private Sub chkSonn_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

      Dim n As Long

10    On Error GoTo chkSonn_MouseUp_Error

20    lblShigellaType = ""

30    For n = 0 To 1
40      If chkSonn(n) = 1 Then lblShigellaType = "Shigella Sonnei"
50    Next

60    cmdSave.Enabled = True

70    Exit Sub

chkSonn_MouseUp_Error:

      Dim strES As String
      Dim intEL As Integer

80    Screen.MousePointer = 0

90    intEL = Erl
100   strES = Err.Description
110   LogError "frmSalmShigWorkSheet", "chkSonn_MouseUp", intEL, strES


End Sub


Private Sub lblShigellaType_DblClick()

10    On Error GoTo lblShigellaType_DblClick_Error

20    If lblShigellaType = "" Then
30      lblShigellaType = "No Shigella Isolated"
40    Else
50      lblShigellaType = ""
60    End If

70    cmdSave.Enabled = True

80    Exit Sub

lblShigellaType_DblClick_Error:

      Dim strES As String
      Dim intEL As Integer

90    Screen.MousePointer = 0

100   intEL = Erl
110   strES = Err.Description
120   LogError "frmSalmShigWorkSheet", "lblShigellaType_DblClick", intEL, strES


End Sub


Private Sub txtColindaleResult_KeyPress(KeyAscii As Integer)

10    On Error GoTo txtColindaleResult_KeyPress_Error

20    cmdSave.Enabled = True

30    Exit Sub

txtColindaleResult_KeyPress_Error:

      Dim strES As String
      Dim intEL As Integer

40    Screen.MousePointer = 0

50    intEL = Erl
60    strES = Err.Description
70    LogError "frmSalmShigWorkSheet", "txtColindaleResult_KeyPress", intEL, strES


End Sub


Private Sub txtSalmID_DblClick()

10    On Error GoTo txtSalmID_DblClick_Error

20    txtSalmID = "No Salmonella Isolated"

30    cmdSave.Enabled = True

40    Exit Sub

txtSalmID_DblClick_Error:

      Dim strES As String
      Dim intEL As Integer

50    Screen.MousePointer = 0

60    intEL = Erl
70    strES = Err.Description
80    LogError "frmSalmShigWorkSheet", "txtSalmID_DblClick", intEL, strES


End Sub


Private Sub txtSalmType_KeyPress(KeyAscii As Integer)

10    On Error GoTo txtSalmType_KeyPress_Error

20    cmdSave.Enabled = True

30    Exit Sub

txtSalmType_KeyPress_Error:

      Dim strES As String
      Dim intEL As Integer

40    Screen.MousePointer = 0

50    intEL = Erl
60    strES = Err.Description
70    LogError "frmSalmShigWorkSheet", "txtSalmType_KeyPress", intEL, strES


End Sub


