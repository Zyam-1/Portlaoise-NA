VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frmMicroOrders 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Microbiology Orders"
   ClientHeight    =   6180
   ClientLeft      =   3135
   ClientTop       =   960
   ClientWidth     =   8625
   Icon            =   "frmMicroOrders.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   8625
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   825
      Left            =   3300
      Picture         =   "frmMicroOrders.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "Save Changes"
      Top             =   5220
      Width           =   1050
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Caption         =   "Exit"
      Height          =   825
      Left            =   4410
      Picture         =   "frmMicroOrders.frx":0316
      Style           =   1  'Graphical
      TabIndex        =   39
      Tag             =   "101"
      ToolTipText     =   "Exit Screen"
      Top             =   5220
      Width           =   1050
   End
   Begin VB.Frame Frame3 
      Caption         =   "Urine Sample"
      Height          =   1245
      Left            =   3300
      TabIndex        =   31
      Top             =   3120
      Width           =   2160
      Begin VB.OptionButton optU 
         Caption         =   "EMU"
         Height          =   195
         Index           =   5
         Left            =   1020
         TabIndex        =   37
         Top             =   900
         Width           =   675
      End
      Begin VB.OptionButton optU 
         Alignment       =   1  'Right Justify
         Caption         =   "FVU"
         Height          =   195
         Index           =   4
         Left            =   330
         TabIndex        =   36
         Top             =   900
         Width           =   645
      End
      Begin VB.OptionButton optU 
         Alignment       =   1  'Right Justify
         Caption         =   "MSU"
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   35
         Top             =   360
         Value           =   -1  'True
         Width           =   675
      End
      Begin VB.OptionButton optU 
         Caption         =   "CSU"
         Height          =   195
         Index           =   1
         Left            =   1020
         TabIndex        =   34
         Top             =   360
         Width           =   645
      End
      Begin VB.OptionButton optU 
         Alignment       =   1  'Right Justify
         Caption         =   "BSU"
         Height          =   195
         Index           =   2
         Left            =   330
         TabIndex        =   33
         Top             =   630
         Width           =   645
      End
      Begin VB.OptionButton optU 
         Caption         =   "SPA"
         Height          =   195
         Index           =   3
         Left            =   1020
         TabIndex        =   32
         Top             =   630
         Width           =   645
      End
   End
   Begin VB.Frame frFaeces 
      Caption         =   "Faecal Requests"
      Height          =   4395
      Left            =   870
      TabIndex        =   15
      Top             =   1650
      Width           =   2325
      Begin VB.CheckBox chkFaecal 
         Caption         =   "Gardia lambila"
         Height          =   195
         Index           =   17
         Left            =   180
         TabIndex        =   43
         Top             =   4050
         Width           =   1365
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "GDH"
         Height          =   195
         Index           =   15
         Left            =   1560
         TabIndex        =   42
         Top             =   2730
         Width           =   735
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "PCR"
         Height          =   195
         Index           =   16
         Left            =   1560
         TabIndex        =   41
         Top             =   2985
         Width           =   645
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "C.Diff Culture"
         Height          =   195
         Index           =   14
         Left            =   180
         TabIndex        =   38
         Top             =   3270
         Width           =   1995
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "Red Sub"
         Height          =   195
         Index           =   13
         Left            =   180
         TabIndex        =   30
         Top             =   3780
         Width           =   1095
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "Salmonella / Shigella"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   29
         Top             =   780
         Width           =   1815
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "Occult Blood"
         Height          =   195
         Index           =   9
         Left            =   720
         TabIndex        =   28
         Top             =   2370
         Width           =   1245
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "Check10"
         Height          =   195
         Index           =   8
         Left            =   450
         TabIndex        =   27
         Top             =   2370
         Width           =   225
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "Check9"
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   26
         Top             =   2370
         Width           =   255
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "H.Pylori"
         Height          =   195
         Index           =   12
         Left            =   180
         TabIndex        =   25
         Top             =   3525
         Width           =   855
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "Toxin A/B"
         Height          =   195
         Index           =   11
         Left            =   180
         TabIndex        =   24
         Top             =   3030
         Width           =   1035
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "Coli 0157"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   23
         Top             =   1260
         Width           =   975
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "Rota"
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   22
         Top             =   1740
         Width           =   735
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "Cryptosporidium"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   21
         Top             =   1500
         Width           =   1425
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "C && S"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   20
         Top             =   390
         Width           =   705
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "Adeno"
         Height          =   195
         Index           =   6
         Left            =   180
         TabIndex        =   19
         Top             =   1980
         Width           =   885
      End
      Begin VB.CheckBox chkKCandS 
         Caption         =   "K - C && S"
         Height          =   195
         Left            =   1110
         TabIndex        =   18
         Top             =   390
         Width           =   975
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "Campylobacter"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   17
         Top             =   1020
         Width           =   1365
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "O / P"
         Height          =   195
         Index           =   10
         Left            =   180
         TabIndex        =   16
         Top             =   2760
         Width           =   705
      End
   End
   Begin VB.TextBox txtSex 
      Height          =   285
      Left            =   3900
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   10
      Top             =   780
      Width           =   1545
   End
   Begin VB.TextBox txtDoB 
      Height          =   285
      Left            =   3900
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   9
      Top             =   420
      Width           =   1545
   End
   Begin VB.TextBox txtName 
      Height          =   300
      Left            =   960
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   8
      Tag             =   "tName"
      Top             =   1110
      Width           =   4485
   End
   Begin VB.TextBox txtChart 
      Height          =   300
      Left            =   960
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   7
      Top             =   780
      Width           =   1545
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
      Height          =   300
      Left            =   960
      MaxLength       =   12
      TabIndex        =   4
      Top             =   420
      Width           =   1545
   End
   Begin VB.Frame Frame2 
      Caption         =   "Urine Requests"
      Height          =   1305
      Left            =   3300
      TabIndex        =   0
      Top             =   1650
      Width           =   2145
      Begin VB.CheckBox chkUrine 
         Caption         =   "Red Sub"
         Height          =   195
         Index           =   2
         Left            =   510
         TabIndex        =   3
         Top             =   870
         Width           =   1215
      End
      Begin VB.CheckBox chkUrine 
         Caption         =   "Pregnancy"
         Height          =   255
         Index           =   1
         Left            =   510
         TabIndex        =   2
         Top             =   570
         Width           =   1155
      End
      Begin VB.CheckBox chkUrine 
         Caption         =   "C && S"
         Height          =   255
         Index           =   0
         Left            =   510
         TabIndex        =   1
         Top             =   315
         Width           =   765
      End
   End
   Begin ComCtl2.UpDown UpDown1 
      Height          =   300
      Left            =   2520
      TabIndex        =   5
      Top             =   420
      Width           =   570
      _ExtentX        =   1005
      _ExtentY        =   529
      _Version        =   327681
      Value           =   1
      BuddyControl    =   "txtSampleID"
      BuddyDispid     =   196620
      OrigLeft        =   1920
      OrigTop         =   540
      OrigRight       =   2160
      OrigBottom      =   1020
      Max             =   99999999
      Min             =   1
      Orientation     =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin MSFlexGridLib.MSFlexGrid GrdMicroPanel 
      Height          =   5580
      Left            =   6075
      TabIndex        =   44
      Top             =   450
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   9843
      _Version        =   393216
      Rows            =   1
      Cols            =   3
      FormatString    =   "|Test                 |    "
   End
   Begin VB.Image imgGreenTick 
      Height          =   225
      Left            =   0
      Picture         =   "frmMicroOrders.frx":0620
      Top             =   210
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgRedCross 
      Height          =   225
      Left            =   0
      Picture         =   "frmMicroOrders.frx":08F6
      Top             =   0
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Chart"
      Height          =   195
      Left            =   540
      TabIndex        =   14
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label43 
      AutoSize        =   -1  'True
      Caption         =   "Sex"
      Height          =   195
      Left            =   3600
      TabIndex        =   13
      Top             =   810
      Width           =   270
   End
   Begin VB.Label Label40 
      AutoSize        =   -1  'True
      Caption         =   "D.o.B"
      Height          =   195
      Index           =   0
      Left            =   3450
      TabIndex        =   12
      Top             =   480
      Width           =   405
   End
   Begin VB.Label Label39 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      Height          =   195
      Index           =   0
      Left            =   480
      TabIndex        =   11
      Top             =   1140
      Width           =   420
   End
   Begin VB.Label Label42 
      AutoSize        =   -1  'True
      Caption         =   "Sample ID"
      Height          =   195
      Left            =   225
      TabIndex        =   6
      Top             =   495
      Width           =   735
   End
End
Attribute VB_Name = "frmMicroOrders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FO As FaecalOrder

Private Sub LoadDetails()

          Dim tb As Recordset
          Dim sql As String
          Dim SampleIDWithOffset As Double
          Dim n As Long

10        On Error GoTo LoadDetails_Error

20        SampleIDWithOffset = Val(txtSampleID) + SysOptMicroOffset(0)

30        GetFaecalOrder Val(txtSampleID), FO
40        With FO
50            chkFaecal(0) = IIf(.cS, 1, 0)
60            chkFaecal(1) = IIf(.ssScreen, 1, 0)
70            chkFaecal(2) = IIf(.Campylobacter, 1, 0)
80            chkFaecal(3) = IIf(.Coli0157, 1, 0)
90            chkFaecal(4) = IIf(.Cryptosporidium, 1, 0)
100           chkFaecal(5) = IIf(.Rota, 1, 0)
110           chkFaecal(6) = IIf(.Adeno, 1, 0)
120           chkFaecal(7) = IIf(.OB0, 1, 0)
130           chkFaecal(8) = IIf(.OB1, 1, 0)
140           chkFaecal(9) = IIf(.OB2, 1, 0)
150           chkFaecal(10) = IIf(.OP, 1, 0)
160           chkFaecal(11) = IIf(.ToxinAB, 1, 0)
170           chkFaecal(12) = IIf(.HPylori, 1, 0)
180           chkFaecal(13) = IIf(.RedSub, 1, 0)
190           chkFaecal(14) = IIf(.CDiffCulture, 1, 0)
200           chkFaecal(15) = IIf(.GDH, 1, 0)
210           chkFaecal(16) = IIf(.PCR, 1, 0)
220           chkFaecal(17) = IIf(.GL, 1, 0)
230           If .ssScreen And .Rota And .Adeno Then
240               chkKCandS.Value = 1
250           End If

260       End With

270       For n = 0 To 2
280           chkUrine(n) = 0
290       Next

300       sql = "SELECT Patname, DoB, Chart, Sex " & _
                "from Demographics WHERE " & _
                "SampleID = " & SampleIDWithOffset & " "
310       Set tb = New Recordset
320       RecOpenServer 0, tb, sql
330       If Not tb.EOF Then
340           txtDoB = tb!Dob & ""
350           txtChart = tb!Chart & ""
360           Select Case UCase$(Left$(tb!sex & "", 1))
              Case "M": txtSex = "Male"
370           Case "F": txtSex = "Female"
380           Case Else: txtSex = ""
390           End Select
400           txtName = tb!PatName & ""
410       Else
420           txtDoB = ""
430           txtChart = ""
440           txtSex = ""
450           txtName = ""
460       End If

470       sql = "SELECT " & _
                "COALESCE(CS, 0) CS, " & _
                "COALESCE(Pregnancy, 0) Pregnancy, " & _
                "COALESCE(RedSub, 0) RedSub " & _
                "FROM UrineRequests WHERE " & _
                "SampleID = " & SampleIDWithOffset
480       Set tb = New Recordset
490       RecOpenServer 0, tb, sql
500       If Not tb.EOF Then
510           If tb!cS Then chkUrine(0) = 1
520           If tb!Pregnancy Then chkUrine(1) = 1
530           If tb!RedSub Then chkUrine(2) = 1
540       End If

550       sql = "SELECT SiteDetails from MicroSiteDetails WHERE " & _
                "SampleID = " & SampleIDWithOffset
560       Set tb = New Recordset
570       RecOpenClient 0, tb, sql
580       If Not tb.EOF Then
590           For n = 0 To 5
600               If UCase$(tb!SiteDetails & "") = optU(n).Caption Then
610                   optU(n).Value = True
620                   Exit For
630               End If
640           Next
650       End If

660       cmdSave.Enabled = False

670       Exit Sub

LoadDetails_Error:

          Dim strES As String
          Dim intEL As Integer

680       intEL = Erl
690       strES = Err.Description
700       LogError "frmMicroOrders", "LoadDetails", intEL, strES, sql

End Sub

Private Sub SaveDetails()

          Dim n As Long
          Dim sql As String
          Dim SampleIDWithOffset As Double
          Dim SiteDetails As String
          Dim FoundU As Boolean
          Dim FoundF As Boolean

10        On Error GoTo SaveDetails_Error

20        For n = 0 To 5
30            If optU(n).Value = True Then
40                SiteDetails = optU(n).Caption
50                Exit For
60            End If
70        Next

80        FoundU = False
90        For n = 0 To 2
100           If chkUrine(n) Then
110               FoundU = True
120           End If
130       Next

140       If Not FoundU Then
150           FoundF = False
160           For n = 0 To 17
170               If chkFaecal(n).Value = 1 Then
180                   FoundF = True
190                   Exit For
200               End If
210           Next
220       End If

'230       If Not FoundU And Not FoundF Then
'240           iMsg "Nothing to Save!", vbExclamation
'250           cmdSave.Enabled = False
'260           Exit Sub
'270       End If

280       SampleIDWithOffset = Val(txtSampleID) + SysOptMicroOffset(0)

290       If FoundU Then
              'Created on 18/02/2011 11:41:08
              'Autogenerated by SQL Scripting

300           sql = "If Exists(Select 1 From UrineRequests " & _
                    "Where SampleID = @SampleID0 ) " & _
                    "Begin " & _
                    "Update UrineRequests Set " & _
                    "SampleID = @SampleID0, " & _
                    "CS = @CS1, " & _
                    "Pregnancy = @Pregnancy2, " & _
                    "RedSub = @RedSub3, " & _
                    "UserName = '@UserName6' " & _
                    "Where SampleID = @SampleID0  " & _
                    "End  " & _
                    "Else " & _
                    "Begin  " & _
                    "Insert Into UrineRequests (SampleID, CS, Pregnancy, RedSub, UserName) Values " & _
                    "(@SampleID0, @CS1, @Pregnancy2, @RedSub3, '@UserName6') " & _
                    "End"

310           sql = Replace(sql, "@SampleID0", SampleIDWithOffset)
320           sql = Replace(sql, "@CS1", IIf(chkUrine(0), 1, 0))
330           sql = Replace(sql, "@Pregnancy2", IIf(chkUrine(1), 1, 0))
340           sql = Replace(sql, "@RedSub3", IIf(chkUrine(2), 1, 0))
350           sql = Replace(sql, "@UserName6", UserName)

360           Cnxn(0).Execute sql
              '    sql = "SELECT * from UrineRequests WHERE " & _
                   '          "SampleID = '" & SampleIDWithOffset & "'"
              '    Set tb = New Recordset
              '    RecOpenServer 0, tb, sql
              '    If tb.EOF Then
              '        tb.AddNew
              '    End If
              '    tb!SampleID = SampleIDWithOffset
              '    tb!cS = IIf(chkUrine(0), 1, 0)
              '    tb!Pregnancy = IIf(chkUrine(1), 1, 0)
              '    tb!RedSub = IIf(chkUrine(2), 1, 0)
              '    tb!UserName = UserName
              '    tb.Update



              'Created on 18/02/2011 11:44:36
              'Autogenerated by SQL Scripting

370           sql = "If Exists(Select 1 From Urine " & _
                    "Where SampleID = @SampleID0 ) " & _
                    "Begin " & _
                    "Update Urine Set " & _
                    "SampleID = @SampleID0, " & _
                    "UserName = '@UserName25' " & _
                    "Where SampleID = @SampleID0  " & _
                    "End  " & _
                    "Else " & _
                    "Begin  " & _
                    "Insert Into Urine (SampleID, UserName) Values " & _
                    "(@SampleID0, '@UserName25') " & _
                    "End"

380           sql = Replace(sql, "@SampleID0", SampleIDWithOffset)
390           sql = Replace(sql, "@UserName25", UserName)

400           Cnxn(0).Execute sql

              '    sql = "SELECT * from urine WHERE sampleid = '" & SampleIDWithOffset & "'"
              '    Set tb = New Recordset
              '    RecOpenServer 0, tb, sql
              '    If tb.EOF Then
              '        tb.AddNew
              '        tb!SampleID = SampleIDWithOffset
              '        tb!UserName = UserName
              '        tb.Update
              '    End If

410           SaveInitialMicroSiteDetails "Urine", SampleIDWithOffset, SiteDetails
              '680     tb.Update
420       ElseIf FoundF Then
430           FO.cS = chkFaecal(0) = 1
440           FO.ssScreen = chkFaecal(1) = 1
450           FO.Campylobacter = chkFaecal(2) = 1
460           FO.Coli0157 = chkFaecal(3) = 1
470           FO.Cryptosporidium = chkFaecal(4) = 1
480           FO.Rota = chkFaecal(5) = 1
490           FO.Adeno = chkFaecal(6) = 1
500           FO.OB0 = chkFaecal(7) = 1
510           FO.OB1 = chkFaecal(8) = 1
520           FO.OB2 = chkFaecal(9) = 1
530           FO.OP = chkFaecal(10) = 1
540           FO.ToxinAB = chkFaecal(11) = 1
550           FO.HPylori = chkFaecal(12) = 1
560           FO.RedSub = chkFaecal(13) = 1
570           FO.CDiffCulture = chkFaecal(14) = 1
              
580           FO.GDH = chkFaecal(15) = 1
590           FO.PCR = chkFaecal(16) = 1
              
600           FO.GL = chkFaecal(17) = 1
610           SaveFaecalOrder Val(txtSampleID), FO

620           SaveInitialMicroSiteDetails "Faeces", SampleIDWithOffset, ""

              'Created on 18/02/2011 11:46:11
              'Autogenerated by SQL Scripting

630           sql = "If Exists(Select 1 From Faeces " & _
                    "Where SampleID = @SampleID0 ) " & _
                    "Begin " & _
                    "Update Faeces Set " & _
                    "SampleID = @SampleID0 " & _
                    "Where SampleID = @SampleID0  " & _
                    "End  " & _
                    "Else " & _
                    "Begin  " & _
                    "Insert Into Faeces (SampleID) Values " & _
                    "(@SampleID0) " & _
                    "End"

640           sql = Replace(sql, "@SampleID0", SampleIDWithOffset)

650           Cnxn(0).Execute sql
              '    sql = "SELECT * from faeces WHERE sampleid = '" & SampleIDWithOffset & "'"
              '    Set tb = New Recordset
              '    RecOpenServer 0, tb, sql
              '    If tb.EOF Then
              '        tb.AddNew
              '        tb!SampleID = SampleIDWithOffset
              '        tb.Update
              '    End If
660       End If

670       cmdSave.Enabled = False

680       Exit Sub

SaveDetails_Error:

          Dim strES As String
          Dim intEL As Integer

690       intEL = Erl
700       strES = Err.Description
710       LogError "frmMicroOrders", "SaveDetails", intEL, strES, sql

End Sub

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


Private Sub chkUrine_Click(Index As Integer)
          Dim n As Integer

          'If Index = 6 Then
          '  chkUrine(7) = 0
          'ElseIf Index = 7 Then
          '  chkUrine(6) = 0
          'End If

10        On Error GoTo chkUrine_Click_Error

20        cmdSave.Enabled = True


30        If Index = 6 Then
40            For n = 0 To 5
50                chkUrine(n).Value = 0
60            Next
70        End If

80        Exit Sub

chkUrine_Click_Error:

          Dim strES As String
          Dim intEL As Integer



90        intEL = Erl
100       strES = Err.Description
110       LogError "frmMicroOrders", "chkUrine_Click", intEL, strES


End Sub

Private Sub cmdCancel_Click()

10        If cmdSave.Enabled Then
20            If iMsg("Cancel without Saving?", vbQuestion = vbYesNo) = vbNo Then
30                Exit Sub
40            End If
50        End If

60        Unload Me

End Sub


Private Sub cmdSave_Click()

10    SaveDetails
20    SaveMicroOrders
End Sub

Private Sub Form_Activate()

10    LoadDetails
20    LoadMicroPanel
End Sub



Private Sub optU_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

10        cmdSave.Enabled = True

End Sub

Private Sub txtSampleID_LostFocus()

10        LoadDetails

End Sub

Private Sub UpDown1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        LoadDetails

End Sub

Private Sub InitializeMicroPanelGrid()


10    With GrdMicroPanel
20        .Rows = 2
30        .Cols = 3
40        .Rows = 1

50        .ColWidth(0) = 0
60        .ColWidth(1) = 1800
70        .ColWidth(2) = 240
80    End With


End Sub
Private Sub LoadMicroPanel()

      Dim tb As Recordset
      Dim sql As String
      Dim SampleIDWithOffset As Double
      Dim n As Long



10    On Error GoTo LoadMicroPanel_Error

20    InitializeMicroPanelGrid
30    SampleIDWithOffset = Val(txtSampleID) + SysOptMicroOffset(0)

      'Load Urine Panels

40    sql = "SELECT distinct Hospital, PanelName as Name from MicroPanels Order by Name"

50    Set tb = New Recordset
60    RecOpenServer 0, tb, sql
70    Do While Not tb.EOF
80        GrdMicroPanel.AddItem tb!Hospital & vbTab & tb!Name
90        GrdMicroPanel.Row = GrdMicroPanel.Rows - 1
100       GrdMicroPanel.Col = 2
110       Set GrdMicroPanel.CellPicture = imgRedCross
120       tb.MoveNext
130   Loop

      'fill known orders
140   sql = "SELECT MR.*, MP.PanelName as Name " & _
            "from MicRequests as MR, MicroPanels as MP WHERE " & _
            "SampleID = '" & Val(SampleIDWithOffset) & "' " & _
            "and MR.Code = MP.PanelName " & _
            "and MR.SampleType = 'U'"
150   Set tb = New Recordset
160   RecOpenServer 0, tb, sql
170   Do While Not tb.EOF

180       For n = 1 To GrdMicroPanel.Rows - 1
190           If tb!Name = GrdMicroPanel.TextMatrix(n, 1) Then
200               GrdMicroPanel.Row = n
210               GrdMicroPanel.Col = 2
220               Set GrdMicroPanel.CellPicture = imgGreenTick
230               Exit For
240           End If
250       Next
260       tb.MoveNext
270   Loop

280   cmdSave.Enabled = False

290   Exit Sub
LoadMicroPanel_Error:
         
300   LogError "frmMicroOrders", "LoadMicroPanel", Erl, Err.Description, sql


End Sub
Private Sub SaveMicroOrders()

      Dim sql As String
      Dim SampleIDWithOffset As Double
      Dim tb As Recordset
      Dim i As Integer

10    On Error GoTo SaveMicroOrders_Error

20    SampleIDWithOffset = Val(txtSampleID) + SysOptMicroOffset(0)
30    Cnxn(0).Execute ("DELETE from MicRequests WHERE " & _
                       "SampleID = '" & SampleIDWithOffset & "' " & _
                       "and Programmed = 0")

40    For i = 1 To GrdMicroPanel.Rows - 1
50        GrdMicroPanel.Row = i
60        GrdMicroPanel.Col = 2
70        If GrdMicroPanel.CellPicture = imgGreenTick Then


80            sql = "INSERT into micRequests " & _
                    "(SampleID, Code, DateTime, SampleType, Programmed, AnalyserID,  Hospital,OrderControl) VALUES " & _
                    "('" & SampleIDWithOffset & "', " & _
                    "'" & GrdMicroPanel.TextMatrix(i, 1) & "', " & _
                    "'" & Format$(Now, "dd/mmm/yyyy hh:mm") & "', " & _
                    "'U', " & _
                    "'0', " & _
                    "'abc', " & _
                    "'" & GrdMicroPanel.TextMatrix(i, 0) & "', " & _
                    "'NW')"
90            Cnxn(0).Execute sql

              
100       End If
110   Next

120   Exit Sub
SaveMicroOrders_Error:
         
130   LogError "frmMicroOrders", "SaveMicroOrders", Erl, Err.Description, sql

End Sub

Private Sub GrdMicroPanel_Click()

10    GrdMicroPanel.Row = GrdMicroPanel.MouseRow
20    GrdMicroPanel.Col = 2
30    If GrdMicroPanel.CellPicture = imgRedCross Then
40        Set GrdMicroPanel.CellPicture = imgGreenTick
50    Else
60        Set GrdMicroPanel.CellPicture = imgRedCross
70    End If
80    cmdSave.Enabled = True
End Sub
