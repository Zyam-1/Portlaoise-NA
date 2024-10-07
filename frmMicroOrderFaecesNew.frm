VERSION 5.00
Begin VB.Form frmMicroOrderFaecesNew 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire"
   ClientHeight    =   5580
   ClientLeft      =   8640
   ClientTop       =   4065
   ClientWidth     =   4410
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   4410
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   705
      Left            =   240
      Picture         =   "frmMicroOrderFaecesNew.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Save Changes"
      Top             =   4740
      Width           =   1275
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Caption         =   "Exit"
      Height          =   705
      Left            =   1920
      Picture         =   "frmMicroOrderFaecesNew.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   19
      Tag             =   "101"
      ToolTipText     =   "Exit Screen"
      Top             =   4740
      Width           =   1275
   End
   Begin VB.Frame frFaeces 
      Caption         =   "Faecal Requests"
      Height          =   4125
      Left            =   240
      TabIndex        =   2
      Top             =   510
      Width           =   3915
      Begin VB.CheckBox chkFaecal 
         Caption         =   "Gardia lambila"
         Height          =   195
         Index           =   17
         Left            =   2160
         TabIndex        =   23
         Top             =   3270
         Width           =   1365
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "PCR"
         Height          =   195
         Index           =   16
         Left            =   2160
         TabIndex        =   22
         Top             =   3015
         Width           =   1005
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "GDH"
         Height          =   195
         Index           =   15
         Left            =   2160
         TabIndex        =   21
         Top             =   2760
         Width           =   1005
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "C.Diff Culture"
         Height          =   195
         Index           =   14
         Left            =   630
         TabIndex        =   18
         Top             =   3270
         Width           =   2055
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "Red Sub"
         Height          =   195
         Index           =   13
         Left            =   630
         TabIndex        =   17
         Top             =   3780
         Width           =   1005
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "O / P"
         Height          =   195
         Index           =   10
         Left            =   630
         TabIndex        =   16
         Top             =   2760
         Width           =   705
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "Campylobacter"
         Height          =   195
         Index           =   2
         Left            =   630
         TabIndex        =   15
         Top             =   1020
         Width           =   1365
      End
      Begin VB.CheckBox chkKCandS 
         Caption         =   "K - C && S"
         Height          =   195
         Left            =   1560
         TabIndex        =   14
         Top             =   390
         Width           =   975
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "Adeno"
         Height          =   195
         Index           =   6
         Left            =   630
         TabIndex        =   13
         Top             =   1980
         Width           =   885
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "C && S"
         Height          =   195
         Index           =   0
         Left            =   630
         TabIndex        =   12
         Top             =   390
         Width           =   705
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "Cryptosporidium"
         Height          =   195
         Index           =   4
         Left            =   630
         TabIndex        =   11
         Top             =   1500
         Width           =   1425
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "Rota"
         Height          =   195
         Index           =   5
         Left            =   630
         TabIndex        =   10
         Top             =   1740
         Width           =   735
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "Coli 0157"
         Height          =   195
         Index           =   3
         Left            =   630
         TabIndex        =   9
         Top             =   1260
         Width           =   975
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "Toxin A/B"
         Height          =   195
         Index           =   11
         Left            =   630
         TabIndex        =   8
         Top             =   3015
         Width           =   1035
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "H.Pylori"
         Height          =   195
         Index           =   12
         Left            =   630
         TabIndex        =   7
         Top             =   3525
         Width           =   855
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "Check9"
         Height          =   195
         Index           =   7
         Left            =   630
         TabIndex        =   6
         Top             =   2370
         Width           =   255
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "Check10"
         Height          =   195
         Index           =   8
         Left            =   900
         TabIndex        =   5
         Top             =   2370
         Width           =   225
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "Occult Blood"
         Height          =   195
         Index           =   9
         Left            =   1170
         TabIndex        =   4
         Top             =   2370
         Width           =   1245
      End
      Begin VB.CheckBox chkFaecal 
         Caption         =   "Salmonella / Shigella"
         Height          =   195
         Index           =   1
         Left            =   630
         TabIndex        =   3
         Top             =   780
         Width           =   1815
      End
      Begin VB.Line Line1 
         X1              =   1860
         X2              =   1860
         Y1              =   2760
         Y2              =   3180
      End
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
      Left            =   1710
      MaxLength       =   12
      TabIndex        =   0
      Top             =   180
      Width           =   1545
   End
   Begin VB.Label Label42 
      AutoSize        =   -1  'True
      Caption         =   "Sample ID"
      Height          =   195
      Left            =   840
      TabIndex        =   1
      Top             =   210
      Width           =   735
   End
End
Attribute VB_Name = "frmMicroOrderFaecesNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


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
110           If UCase$(HospName(0)) = "MULLINGAR" Or UCase$(HospName(0)) = "TULLAMORE" Then
120               chkFaecal(4).Value = 1
130           End If
140       End If

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
20            If iMsg("Cancel without Saving?", vbQuestion = vbYesNo) = vbNo Then
30                Exit Sub
40            End If
50        End If

60        Unload Me

End Sub


Private Sub cmdSave_Click()

10        SaveDetails

20        Me.Hide

End Sub


Private Sub Form_Activate()

10        LoadFaecalOrders

End Sub

Private Sub LoadFaecalOrders()

      Dim tb As Recordset
      Dim sql As String
      Dim SampleIDWithOffset As Double
      Dim n As Integer

10    On Error GoTo LoadFaecalOrders_Error

20    SampleIDWithOffset = Val(txtSampleID) + SysOptMicroOffset(0)

30    For n = 0 To 17
40        chkFaecal(n) = 0
50    Next

60    sql = "SELECT " & _
            "COALESCE(CS, 0) CS, " & _
            "COALESCE(ssScreen, 0) ssScreen, " & _
            "COALESCE(Campylobacter, 0) Campylobacter, " & _
            "COALESCE(Coli0157, 0) Coli0157, " & _
            "COALESCE(Cryptosporidium, 0) Cryptosporidium, " & _
            "COALESCE(Rota, 0) Rota, " & _
            "COALESCE(Adeno, 0) Adeno, " & _
            "COALESCE(OB0, 0) OB0, " & _
            "COALESCE(OB1, 0) OB1, " & _
            "COALESCE(OB2, 0) OB2, " & _
            "COALESCE(OP, 0) OP, " & _
            "COALESCE(ToxinAB, 0) ToxinAB, " & _
            "COALESCE(CDiff, 0) CDiffCulture, " & _
            "COALESCE(HPylori, 0) HPylori, " & _
            "COALESCE(GDH, 0) GDH, " & _
            "COALESCE(PCR, 0) PCR, " & _
            "COALESCE(GL, 0) GL, " & _
            "COALESCE(RedSub, 0) RedSub " & _
            "FROM FaecalRequests where " & _
            "SampleID = " & SampleIDWithOffset & " "
70    Set tb = New Recordset
80    RecOpenServer 0, tb, sql
90    If Not tb.EOF Then
100       chkFaecal(0) = IIf(tb!cS, 1, 0)
110       chkFaecal(1) = IIf(tb!ssScreen, 1, 0)
120       chkFaecal(2) = IIf(tb!Campylobacter, 1, 0)
130       chkFaecal(3) = IIf(tb!Coli0157, 1, 0)
140       chkFaecal(4) = IIf(tb!Cryptosporidium, 1, 0)
150       chkFaecal(5) = IIf(tb!Rota, 1, 0)
160       chkFaecal(6) = IIf(tb!Adeno, 1, 0)
170       chkFaecal(7) = IIf(tb!OB0, 1, 0)
180       chkFaecal(8) = IIf(tb!OB1, 1, 0)
190       chkFaecal(9) = IIf(tb!OB2, 1, 0)
200       chkFaecal(10) = IIf(tb!OP, 1, 0)
210       chkFaecal(11) = IIf(tb!ToxinAB, 1, 0)
220       chkFaecal(12) = IIf(tb!HPylori, 1, 0)
230       chkFaecal(13) = IIf(tb!RedSub, 1, 0)
240       chkFaecal(14) = IIf(tb!CDiffCulture, 1, 0)
250       chkFaecal(15) = IIf(tb!GDH, 1, 0)
260       chkFaecal(16) = IIf(tb!PCR, 1, 0)
270       chkFaecal(17) = IIf(tb!GL, 1, 0)
280   End If

290   cmdSave.Enabled = False

300   Exit Sub

LoadFaecalOrders_Error:

      Dim strES As String
      Dim intEL As Integer

310   intEL = Erl
320   strES = Err.Description
330   LogError "frmMicroOrderFaecesNew", "LoadFaecalOrders", intEL, strES, sql

End Sub


Private Sub SaveDetails()

          Dim lngF As Long
          Dim n As Integer
          Dim sql As String
          Dim SampleIDWithOffset As Double

10        On Error GoTo SaveDetails_Error

20        lngF = 0
30        For n = 0 To 16
40            If chkFaecal(n) Then
50                lngF = lngF + 2 ^ n
60            End If
70        Next

80        If lngF = 0 Then
90            iMsg "Nothing to Save!", vbExclamation
100           cmdSave.Enabled = False
110           Exit Sub
120       End If

130       SampleIDWithOffset = Val(txtSampleID) + SysOptMicroOffset(0)

          'Created on 07/10/2010 15:53:11
          'Autogenerated by SQL Scripting

140       sql = "If Exists(Select 1 From FaecalRequests " & _
                "Where SampleID = @SampleID ) " & _
                "Begin " & _
                "Update FaecalRequests Set " & _
                "SampleID = @SampleID, OP = @OP, Rota = @Rota, Adeno = @Adeno, Coli0157 = @Coli0157, OB0 = @OB0, OB1 = @OB1, OB2 = @OB2, ssScreen = '@ssScreen', cS = '@cS', Campylobacter = '@Campylobacter', Cryptosporidium = '@Cryptosporidium', ToxinAB = '@ToxinAB', HPylori = '@HPylori', RedSub = @RedSub, CDiff = @CDiff, GDH = @GDH, PCR = @PCR ,GL=@GL" & _
                "Where SampleID = @SampleID  " & _
                "End  " & _
                "Else " & _
                "Begin  " & _
                "Insert Into FaecalRequests (SampleID, OP, Rota, Adeno, Coli0157, OB0, OB1, OB2, ssScreen, cS, Campylobacter, Cryptosporidium, ToxinAB, HPylori, RedSub, CDiff, GDH, PCR,GL) Values (@SampleID, @OP, @Rota, @Adeno, @Coli0157, @OB0, @OB1, @OB2, '@ssScreen', '@cS', '@Campylobacter', '@Cryptosporidium', '@ToxinAB', '@HPylori', @RedSub, @CDiff, @GDH, @PCR,@GL) " & _
                "End"

150       sql = Replace(sql, "@SampleID", SampleIDWithOffset)
160       sql = Replace(sql, "@OP", IIf(chkFaecal(10), 1, 0))
170       sql = Replace(sql, "@Rota", IIf(chkFaecal(5), 1, 0))
180       sql = Replace(sql, "@Adeno", IIf(chkFaecal(6), 1, 0))
190       sql = Replace(sql, "@Coli0157", IIf(chkFaecal(3), 1, 0))
200       sql = Replace(sql, "@OB0", IIf(chkFaecal(7), 1, 0))
210       sql = Replace(sql, "@OB1", IIf(chkFaecal(8), 1, 0))
220       sql = Replace(sql, "@OB2", IIf(chkFaecal(9), 1, 0))
230       sql = Replace(sql, "@ssScreen", IIf(chkFaecal(1), 1, 0))
240       sql = Replace(sql, "@cS", IIf(chkFaecal(0), 1, 0))
250       sql = Replace(sql, "@Campylobacter", IIf(chkFaecal(2), 1, 0))
260       sql = Replace(sql, "@Cryptosporidium", IIf(chkFaecal(4), 1, 0))
270       sql = Replace(sql, "@ToxinAB", IIf(chkFaecal(11), 1, 0))
280       sql = Replace(sql, "@HPylori", IIf(chkFaecal(12), 1, 0))
290       sql = Replace(sql, "@RedSub", IIf(chkFaecal(13), 1, 0))
300       sql = Replace(sql, "@CDiff", IIf(chkFaecal(14), 1, 0))
310       sql = Replace(sql, "@GDH", IIf(chkFaecal(15), 1, 0))
320       sql = Replace(sql, "@PCR", IIf(chkFaecal(16), 1, 0))
330       sql = Replace(sql, "@GL", IIf(chkFaecal(17), 1, 0))
          
340       Cnxn(0).Execute sql




350       SaveInitialMicroSiteDetails "Faeces", SampleIDWithOffset, ""

360       cmdSave.Enabled = False

370       Exit Sub

SaveDetails_Error:

          Dim strES As String
          Dim intEL As Integer

380       intEL = Erl
390       strES = Err.Description
400       LogError "frmMicroOrderFaecesNew", "SaveDetails", intEL, strES, sql

End Sub



Private Sub txtSampleID_LostFocus()

10        LoadFaecalOrders

End Sub



