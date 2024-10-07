VERSION 5.00
Begin VB.Form frmBatchLogInFaeces 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Faeces Batch Sample Log In"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5490
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   5490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frFaeces 
      Caption         =   "Faecal Requests"
      Height          =   3825
      Left            =   1260
      TabIndex        =   21
      Top             =   1320
      Width           =   3015
      Begin VB.CheckBox chkFaecal 
         Caption         =   "Red Sub"
         Height          =   195
         Index           =   13
         Left            =   630
         TabIndex        =   22
         Top             =   3480
         Width           =   1005
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "Salmonella / Shigella"
         Height          =   195
         Index           =   1
         Left            =   630
         TabIndex        =   4
         Top             =   780
         Width           =   1815
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "Occult Blood"
         Height          =   195
         Index           =   9
         Left            =   1170
         TabIndex        =   12
         Top             =   2370
         Width           =   1245
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "      Occult Blood"
         Height          =   195
         Index           =   8
         Left            =   900
         TabIndex        =   11
         Top             =   2370
         Width           =   1575
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "            Occult Blood"
         Height          =   195
         Index           =   7
         Left            =   630
         TabIndex        =   10
         Top             =   2370
         Width           =   1785
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "H.Pylori"
         Height          =   195
         Index           =   12
         Left            =   630
         TabIndex        =   15
         Top             =   3240
         Width           =   855
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "Toxin A/B"
         Height          =   195
         Index           =   11
         Left            =   630
         TabIndex        =   14
         Top             =   3000
         Width           =   1035
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "Coli 0157"
         Height          =   195
         Index           =   3
         Left            =   630
         TabIndex        =   6
         Top             =   1260
         Width           =   975
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "Rota"
         Height          =   195
         Index           =   5
         Left            =   630
         TabIndex        =   8
         Top             =   1740
         Width           =   735
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "Cryptosporidium"
         Height          =   195
         Index           =   4
         Left            =   630
         TabIndex        =   7
         Top             =   1500
         Width           =   1425
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "C && S"
         Height          =   195
         Index           =   0
         Left            =   630
         TabIndex        =   2
         Top             =   390
         Width           =   705
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "Adeno"
         Height          =   195
         Index           =   6
         Left            =   630
         TabIndex        =   9
         Top             =   1980
         Width           =   885
      End
      Begin VB.CheckBox chkKCandS 
         Caption         =   "K - C && S"
         Height          =   195
         Left            =   1560
         TabIndex        =   3
         Top             =   390
         Width           =   975
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "Campylobacter"
         Height          =   195
         Index           =   2
         Left            =   630
         TabIndex        =   5
         Top             =   1020
         Width           =   1365
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "O / P"
         Height          =   195
         Index           =   10
         Left            =   630
         TabIndex        =   13
         Top             =   2760
         Width           =   705
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   795
      Left            =   3000
      Picture         =   "frmBatchLogInFaeces.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5370
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      Caption         =   "Numbers"
      Height          =   1065
      Left            =   600
      TabIndex        =   18
      Top             =   90
      Width           =   4245
      Begin VB.TextBox txtFromNumber 
         Height          =   285
         Left            =   780
         TabIndex        =   0
         Top             =   540
         Width           =   1305
      End
      Begin VB.TextBox txtToNumber 
         Height          =   285
         Left            =   2280
         TabIndex        =   1
         Top             =   570
         Width           =   1305
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Start Number"
         Height          =   195
         Left            =   930
         TabIndex        =   20
         Top             =   345
         Width           =   930
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Stop Number"
         Height          =   195
         Left            =   2310
         TabIndex        =   19
         Top             =   375
         Width           =   930
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   795
      Left            =   1320
      Picture         =   "frmBatchLogInFaeces.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5370
      Width           =   1245
   End
End
Attribute VB_Name = "frmBatchLogInFaeces"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FO As FaecalOrder

Private Sub chkFaecal_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

10        cmdSave.Enabled = True

20        If chkFaecal(0).Value = 1 And _
             chkFaecal(5).Value = 1 And _
             chkFaecal(6).Value = 1 Then
30            chkKCandS.Value = 1
40        Else
50            chkKCandS.Value = 0
60        End If

70        If chkFaecal(0).Value = 1 Then
80            chkFaecal(1).Value = 1
90            chkFaecal(2).Value = 1
100           chkFaecal(3).Value = 1
110           chkFaecal(4).Value = 1
120       End If

End Sub


Private Sub chkKCandS_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        If chkKCandS.Value = 1 Then
20            chkFaecal(0).Value = 1
30            chkFaecal(1).Value = 1
40            chkFaecal(2).Value = 1
50            chkFaecal(3).Value = 1
60            chkFaecal(4).Value = 1
70            chkFaecal(5).Value = 1
80            chkFaecal(6).Value = 1
90        End If

100       cmdSave.Enabled = True

End Sub


Private Sub cmdCancel_Click()

10        If cmdSave.Enabled Then
20            If iMsg("Cancel without Saving?", vbQuestion + vbYesNo) = vbNo Then
30                Exit Sub
40            End If
50        End If

60        Unload Me

End Sub


Private Sub cmdSave_Click()

10        SaveDetails

20        cmdSave.Enabled = False

End Sub

Private Sub SaveDetails()

          Dim TotalNumbers As Long
          Dim c As Long

10        On Error GoTo SaveDetails_Error

20        If Trim(txtFromNumber) = "" Then
30            iMsg "Enter Start Number"
40            Exit Sub
50        End If

60        If Trim(txtToNumber) = "" Then
70            iMsg "Enter Stop Number"
80            Exit Sub
90        End If

100       If Val(txtFromNumber) = 0 Then
110           iMsg "Invalid Start Number"
120           Exit Sub
130       End If

140       If Val(txtToNumber) = 0 Then
150           iMsg "Invalid Stop Number"
160           Exit Sub
170       End If

180       If Val(txtFromNumber) > Val(txtToNumber) Then
190           iMsg "Stop Number must be greater than Start Number"
200           Exit Sub
210       End If

220       TotalNumbers = Val(txtToNumber) - Val(txtFromNumber)
230       If TotalNumbers > 100 Then
240           iMsg "Only 100 numbers per batch!", vbExclamation
250           Exit Sub
260       End If

270       If TotalNumbers > 20 Then
280           If iMsg("Save " & Format(TotalNumbers) & " Numbers?", vbQuestion + vbYesNo) = vbNo Then
290               Exit Sub
300           End If
310       End If

320       FO.cS = chkFaecal(0) = 1
330       FO.ssScreen = chkFaecal(1) = 1
340       FO.Campylobacter = chkFaecal(2) = 1
350       FO.Coli0157 = chkFaecal(3) = 1
360       FO.Cryptosporidium = chkFaecal(4) = 1
370       FO.Rota = chkFaecal(5) = 1
380       FO.Adeno = chkFaecal(6) = 1
390       FO.OB0 = chkFaecal(7) = 1
400       FO.OB1 = chkFaecal(8) = 1
410       FO.OB2 = chkFaecal(9) = 1
420       FO.OP = chkFaecal(10) = 1
430       FO.ToxinAB = chkFaecal(11) = 1
440       FO.HPylori = chkFaecal(12) = 1
450       FO.RedSub = chkFaecal(13) = 1

460       For c = Val(txtFromNumber) To Val(txtToNumber)

470           SaveFaecalOrder c, FO

480           SaveInitialMicroSiteDetails "Faeces", c + SysOptMicroOffset(0), ""

490       Next

500       Exit Sub

SaveDetails_Error:

          Dim strES As String
          Dim intEL As Integer

510       intEL = Erl
520       strES = Err.Description
530       LogError "frmBatchLogInFaeces", "SaveDetails", intEL, strES

End Sub

Private Sub txtFromNumber_KeyPress(KeyAscii As Integer)

10        cmdSave.Enabled = True

End Sub


Private Sub txtToNumber_KeyPress(KeyAscii As Integer)

10        cmdSave.Enabled = True

End Sub


