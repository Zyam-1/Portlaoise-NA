VERSION 5.00
Begin VB.Form frmBatchLogInUrine 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Urine Batch Sample Log In"
   ClientHeight    =   3765
   ClientLeft      =   2145
   ClientTop       =   1545
   ClientWidth     =   6465
   ControlBox      =   0   'False
   Icon            =   "frmBatchLogInUrine.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   6465
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Urine Sample"
      Height          =   1245
      Left            =   4170
      TabIndex        =   10
      Top             =   1260
      Width           =   1950
      Begin VB.OptionButton optU 
         Caption         =   "SPA"
         Height          =   195
         Index           =   3
         Left            =   990
         TabIndex        =   16
         Top             =   600
         Width           =   645
      End
      Begin VB.OptionButton optU 
         Alignment       =   1  'Right Justify
         Caption         =   "BSU"
         Height          =   195
         Index           =   2
         Left            =   300
         TabIndex        =   15
         Top             =   600
         Width           =   645
      End
      Begin VB.OptionButton optU 
         Caption         =   "CSU"
         Height          =   195
         Index           =   1
         Left            =   990
         TabIndex        =   14
         Top             =   330
         Width           =   645
      End
      Begin VB.OptionButton optU 
         Alignment       =   1  'Right Justify
         Caption         =   "MSU"
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   13
         Top             =   330
         Value           =   -1  'True
         Width           =   675
      End
      Begin VB.OptionButton optU 
         Alignment       =   1  'Right Justify
         Caption         =   "FVU"
         Height          =   195
         Index           =   4
         Left            =   300
         TabIndex        =   12
         Top             =   870
         Width           =   645
      End
      Begin VB.OptionButton optU 
         Caption         =   "EMU"
         Height          =   195
         Index           =   5
         Left            =   990
         TabIndex        =   11
         Top             =   870
         Width           =   675
      End
   End
   Begin VB.Frame frUrine 
      Caption         =   "Urine Requests"
      Height          =   1245
      Left            =   510
      TabIndex        =   7
      Top             =   1260
      Width           =   1950
      Begin VB.CheckBox chkUrine 
         Caption         =   "Red Sub"
         Height          =   195
         Index           =   2
         Left            =   330
         TabIndex        =   17
         Top             =   870
         Width           =   1125
      End
      Begin VB.CheckBox chkUrine 
         Caption         =   "Pregnancy"
         Height          =   195
         Index           =   1
         Left            =   330
         TabIndex        =   9
         Top             =   600
         Width           =   1155
      End
      Begin VB.CheckBox chkUrine 
         Caption         =   "C && S"
         Height          =   195
         Index           =   0
         Left            =   330
         TabIndex        =   8
         Top             =   330
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   795
      Left            =   1965
      Picture         =   "frmBatchLogInUrine.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2670
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      Caption         =   "Numbers"
      Height          =   1065
      Left            =   510
      TabIndex        =   1
      Top             =   90
      Width           =   5595
      Begin VB.TextBox txtToNumber 
         Height          =   285
         Left            =   4110
         TabIndex        =   5
         Top             =   450
         Width           =   1095
      End
      Begin VB.TextBox txtFromNumber 
         Height          =   285
         Left            =   1260
         TabIndex        =   4
         Top             =   450
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Stop Number"
         Height          =   195
         Left            =   3060
         TabIndex        =   3
         Top             =   495
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Start Number"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   495
         Width           =   930
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   795
      Left            =   3765
      Picture         =   "frmBatchLogInUrine.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2670
      Width           =   1245
   End
End
Attribute VB_Name = "frmBatchLogInUrine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SiteDetails As String

Private Sub chkUrine_Click(Index As Integer)

10        On Error GoTo chkUrine_Click_Error

20        cmdSave.Enabled = True

30        Exit Sub

chkUrine_Click_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmBatchLogIn", "chkUrine_Click", intEL, strES


End Sub

Private Sub cmdCancel_Click()

10        Unload Me

End Sub



Private Sub cmdSave_Click()

10        SaveTestsRequested

End Sub

Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        SiteDetails = optU(0).Caption

30        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmBatchLogIn", "Form_Load", intEL, strES


End Sub

Private Sub optU_Click(Index As Integer)

10        On Error GoTo optU_Click_Error

20        SiteDetails = optU(Index).Caption
30        cmdSave.Enabled = True

40        Exit Sub

optU_Click_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmBatchLogIn", "optU_Click", intEL, strES


End Sub

Private Sub SaveTestsRequested()

          Dim sql As String
          Dim n As Long
          Dim TotalNumbers As Long
          Dim c As Double
          Dim SampleIDWithOffset As Double
          Dim Found As Boolean

10        On Error GoTo SaveTestsRequested_Error

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
230       If TotalNumbers > 150 Then
240           iMsg "Only 150 numbers per batch!", vbExclamation
250           Exit Sub
260       End If

270       If TotalNumbers > 20 Then
280           If iMsg("Save " & Format(TotalNumbers) & " Numbers?", vbQuestion + vbYesNo) = vbNo Then
290               Exit Sub
300           End If
310       End If

320       Found = False
330       For n = 0 To 2
340           If chkUrine(n) Then
350               Found = True
360           End If
370       Next
380       If Not Found Then
390           iMsg "No tests requested!", vbExclamation
400           Exit Sub
410       End If

420       For c = Val(txtFromNumber) To Val(txtToNumber)

430           SampleIDWithOffset = c + SysOptMicroOffset(0)

440           sql = "If Exists(Select 1 From UrineRequests " & _
                    "Where SampleID = @SampleID0 ) " & _
                    "Begin " & _
                    "Update UrineRequests Set " & _
                    "CS = @CS1, Pregnancy = @Pregnancy2, RedSub = @RedSub3, UserName = '@UserName6' " & _
                    "Where SampleID = @SampleID0  " & _
                    "End  " & _
                    "Else " & _
                    "Begin  " & _
                    "Insert Into UrineRequests (SampleID, CS, Pregnancy, RedSub, UserName) Values " & _
                    "(@SampleID0, @CS1, @Pregnancy2, @RedSub3, '@UserName6') " & _
                    "End"

450           sql = Replace(sql, "@SampleID0", SampleIDWithOffset)
460           sql = Replace(sql, "@CS1", IIf(chkUrine(0), 1, 0))
470           sql = Replace(sql, "@Pregnancy2", IIf(chkUrine(1), 1, 0))
480           sql = Replace(sql, "@RedSub3", IIf(chkUrine(2), 1, 0))
490           sql = Replace(sql, "@UserName6", UserName)

500           Cnxn(0).Execute sql


510           SaveInitialMicroSiteDetails "Urine", SampleIDWithOffset, SiteDetails

              'Created on 01/02/2011 15:49:30
              'Autogenerated by SQL Scripting

520           sql = "If Exists(Select 1 From Urine " & _
                    "Where SampleID = @SampleID0) " & _
                    "Begin " & _
                    "Update Urine Set " & _
                    "SampleID = @SampleID0, " & _
                    "UserName = '@UserName25' " & _
                    "Where SampleID = @SampleID0 " & _
                    "End  " & _
                    "Else " & _
                    "Begin  " & _
                    "Insert Into Urine (SampleID, UserName) Values " & _
                    "(@SampleID0, '@UserName25') " & _
                    "End"

530           sql = Replace(sql, "@SampleID0", SampleIDWithOffset)
540           sql = Replace(sql, "@UserName25", UserName)

550           Cnxn(0).Execute sql

560       Next

570       cmdSave.Enabled = False

580       Exit Sub

SaveTestsRequested_Error:

          Dim strES As String
          Dim intEL As Integer

590       intEL = Erl
600       strES = Err.Description
610       LogError "frmBatchLogInUrine", "SaveTestsRequested", intEL, strES, sql

End Sub

Private Sub txtFromNumber_Change()
10        cmdSave.Enabled = True
End Sub

Private Sub txtToNumber_Change()
10        cmdSave.Enabled = True
End Sub
