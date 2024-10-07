VERSION 5.00
Begin VB.Form frmMicroSetUp 
   Caption         =   "NetAcquire - Microbiology Set Up"
   ClientHeight    =   3525
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5550
   LinkTopic       =   "Form1"
   ScaleHeight     =   3525
   ScaleWidth      =   5550
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkFluids 
      Alignment       =   1  'Right Justify
      Caption         =   "Fluids"
      Height          =   195
      Left            =   1050
      TabIndex        =   15
      Top             =   3030
      Width           =   705
   End
   Begin VB.CheckBox chkBC 
      Alignment       =   1  'Right Justify
      Caption         =   "Blood Culture"
      Height          =   195
      Left            =   510
      TabIndex        =   14
      Top             =   2640
      Width           =   1245
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   825
      Left            =   4140
      Picture         =   "frmMicroSetUp.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1170
      Width           =   1155
   End
   Begin VB.CommandButton bcancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   825
      Left            =   4140
      Picture         =   "frmMicroSetUp.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2400
      Width           =   1155
   End
   Begin VB.CheckBox chkHPylori 
      Caption         =   "H.Pylori"
      Height          =   195
      Left            =   1860
      TabIndex        =   11
      Top             =   2640
      Width           =   1065
   End
   Begin VB.CheckBox chkOP 
      Caption         =   "O/P"
      Height          =   195
      Left            =   1875
      TabIndex        =   10
      Top             =   2250
      Width           =   885
   End
   Begin VB.CheckBox chkCdiff 
      Caption         =   "C Diff"
      Height          =   195
      Left            =   1875
      TabIndex        =   9
      Top             =   1890
      Width           =   915
   End
   Begin VB.CheckBox chkRSV 
      Caption         =   "RSV"
      Height          =   195
      Left            =   1875
      TabIndex        =   8
      Top             =   1530
      Width           =   795
   End
   Begin VB.CheckBox chkRedSub 
      Caption         =   "Red/Sub"
      Height          =   195
      Left            =   1875
      TabIndex        =   7
      Top             =   1170
      Width           =   1215
   End
   Begin VB.CheckBox chkRA 
      Caption         =   "Rota/Adeno"
      Height          =   195
      Left            =   1875
      TabIndex        =   6
      Top             =   810
      Width           =   1305
   End
   Begin VB.CheckBox chkFob 
      Alignment       =   1  'Right Justify
      Caption         =   "FOB"
      Height          =   195
      Left            =   1140
      TabIndex        =   5
      Top             =   2250
      Width           =   615
   End
   Begin VB.CheckBox chkCS 
      Alignment       =   1  'Right Justify
      Caption         =   "C && S"
      Height          =   195
      Left            =   1050
      TabIndex        =   4
      Top             =   1890
      Width           =   705
   End
   Begin VB.CheckBox chkFae 
      Alignment       =   1  'Right Justify
      Caption         =   "Faeces"
      Height          =   195
      Left            =   930
      TabIndex        =   3
      Top             =   1530
      Width           =   825
   End
   Begin VB.CheckBox chkUrIn 
      Alignment       =   1  'Right Justify
      Caption         =   "Indentification"
      Height          =   195
      Left            =   450
      TabIndex        =   2
      Top             =   1170
      Width           =   1305
   End
   Begin VB.CheckBox chkUrine 
      Alignment       =   1  'Right Justify
      Caption         =   "Urine"
      Height          =   195
      Left            =   1080
      TabIndex        =   1
      Top             =   810
      Width           =   675
   End
   Begin VB.ComboBox cmbSite 
      Height          =   315
      Left            =   915
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   315
      Width           =   1860
   End
End
Attribute VB_Name = "frmMicroSetUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private Sub FillSites()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo FillSites_Error

20        cmbSite.Clear

30        sql = "SELECT Text FROM Lists WHERE " & _
                "ListType = 'SI' " & _
                "ORDER BY ListOrder"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        Do While Not tb.EOF
70            cmbSite.AddItem tb!Text & ""
80            tb.MoveNext
90        Loop
100       FixComboWidth cmbSite
110       Exit Sub

FillSites_Error:

          Dim strES As String
          Dim intEL As Integer

120       intEL = Erl
130       strES = Err.Description
140       LogError "frmMicroSetUp", "FillSites", intEL, strES, sql

End Sub

Private Sub bcancel_Click()

10        Unload Me

End Sub

Private Sub chkBC_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        cmdSave.Enabled = True

End Sub


Private Sub chkCdiff_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        cmdSave.Enabled = True

End Sub


Private Sub chkCS_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        cmdSave.Enabled = True

End Sub


Private Sub chkFae_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        cmdSave.Enabled = True

End Sub


Private Sub chkFluids_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        cmdSave.Enabled = True

End Sub


Private Sub chkFob_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        cmdSave.Enabled = True

End Sub


Private Sub chkHPylori_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        cmdSave.Enabled = True

End Sub


Private Sub chkOP_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        cmdSave.Enabled = True

End Sub


Private Sub chkRA_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        cmdSave.Enabled = True

End Sub


Private Sub chkRedSub_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        cmdSave.Enabled = True

End Sub


Private Sub chkRSV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        cmdSave.Enabled = True

End Sub


Private Sub chkUrIn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        cmdSave.Enabled = True

End Sub


Private Sub chkUrine_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        cmdSave.Enabled = True

End Sub


Private Sub cmbSite_Click()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo cmbSite_Click_Error

20        chkUrine = 0
30        chkUrIn = 0
40        chkFae = 0
50        chkCS = 0
60        chkFob = 0
70        chkRA = 0
80        chkRedSub = 0
90        chkRSV = 0
100       chkCdiff = 0
110       chkOP = 0
120       chkHPylori = 0
130       chkBC = 0
140       chkFluids = 0

150       sql = "SELECT * FROM MicroSetup WHERE Site = '" & cmbSite & "'"
160       Set tb = New Recordset
170       RecOpenServer 0, tb, sql
180       If Not tb.EOF Then
190           chkUrine = tb!Urine
200           chkUrIn = tb!UrIdent
210           chkFae = tb!Faeces
220           chkCS = tb!cS
230           chkFob = tb!FOB
240           chkRA = tb!Rota
250           chkRedSub = tb!rs
260           chkRSV = tb!RSV
270           chkCdiff = tb!CDiff
280           chkOP = tb!OP
290           chkHPylori = tb!HPylori
300           chkBC = tb!BC
310           chkFluids = tb!Fluids
320           tb.MoveNext
330       End If

340       Exit Sub

cmbSite_Click_Error:

          Dim strES As String
          Dim intEL As Integer

350       intEL = Erl
360       strES = Err.Description
370       LogError "frmMicroSetUp", "cmbSite_Click", intEL, strES, sql

End Sub

Private Sub cmdSave_Click()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo cmdSave_Click_Error

20        If cmbSite = "" Then
30            iMsg "Please select a site first", vbInformation
40            Exit Sub
50        End If

60        sql = "SELECT * FROM MicroSetup WHERE " & _
                "Site = '" & cmbSite & "'"
70        Set tb = New Recordset
80        RecOpenServer 0, tb, sql
90        If tb.EOF Then tb.AddNew
100       tb!Site = cmbSite
110       tb!Urine = chkUrine
120       tb!UrIdent = chkUrIn
130       tb!Faeces = chkFae
140       tb!cS = chkCS
150       tb!FOB = chkFob
160       tb!Rota = chkRA
170       tb!rs = chkRedSub
180       tb!RSV = chkRSV
190       tb!CDiff = chkCdiff
200       tb!OP = chkOP
210       tb!BC = chkBC
220       tb!HPylori = chkHPylori
230       tb!Fluids = chkFluids
240       tb.Update

250       cmdSave.Enabled = False

260       Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

270       intEL = Erl
280       strES = Err.Description
290       LogError "frmMicroSetUp", "cmdsave_Click", intEL, strES, sql

End Sub

Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        FillSites

30        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmMicroSetUp", "Form_Load", intEL, strES

End Sub
