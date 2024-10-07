VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frmOrderPhoresis 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Order Phoresis"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ComCtl2.UpDown udPhoresisNumber 
      Height          =   585
      Left            =   3015
      TabIndex        =   21
      Top             =   2460
      Width           =   240
      _ExtentX        =   450
      _ExtentY        =   1032
      _Version        =   327681
      BuddyControl    =   "txtPhoresisNumber"
      BuddyDispid     =   196609
      OrigLeft        =   3510
      OrigTop         =   2400
      OrigRight       =   3750
      OrigBottom      =   2775
      Max             =   9999
      SyncBuddy       =   -1  'True
      Wrap            =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtPhoresisNumber 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   480
      Left            =   2220
      MaxLength       =   4
      TabIndex        =   20
      Text            =   "0000"
      Top             =   2490
      Width           =   765
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "E&xit"
      Height          =   1065
      Left            =   3600
      Picture         =   "frmOrderPhoresis.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5280
      Width           =   885
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   1065
      Left            =   3600
      Picture         =   "frmOrderPhoresis.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3090
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.ListBox lstOrder 
      Height          =   3270
      IntegralHeight  =   0   'False
      Left            =   990
      TabIndex        =   0
      Top             =   3060
      Width           =   2265
   End
   Begin VB.Label lblSampleID 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   60
      TabIndex        =   22
      Top             =   3180
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Phoresis Number"
      Height          =   195
      Left            =   1020
      TabIndex        =   19
      Top             =   2610
      Width           =   1200
   End
   Begin VB.Label lblSampleDate 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   990
      TabIndex        =   18
      Top             =   1770
      Width           =   1410
   End
   Begin VB.Label lblAandE 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3090
      TabIndex        =   17
      Top             =   570
      Width           =   1410
   End
   Begin VB.Label lblSex 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4110
      TabIndex        =   16
      Top             =   990
      Width           =   390
   End
   Begin VB.Label lblAge 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3090
      TabIndex        =   15
      Top             =   990
      Width           =   420
   End
   Begin VB.Label lblClinician 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   990
      TabIndex        =   14
      Top             =   1380
      Width           =   3510
   End
   Begin VB.Label lblDoB 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   990
      TabIndex        =   13
      Top             =   990
      Width           =   1410
   End
   Begin VB.Label lblChart 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   990
      TabIndex        =   12
      Top             =   570
      Width           =   1410
   End
   Begin VB.Label lblName 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   990
      TabIndex        =   11
      Top             =   180
      Width           =   3510
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Sample Date"
      Height          =   195
      Left            =   30
      TabIndex        =   10
      Top             =   1800
      Width           =   915
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Clinician"
      Height          =   195
      Left            =   360
      TabIndex        =   9
      Top             =   1410
      Width           =   585
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Sex"
      Height          =   195
      Left            =   3810
      TabIndex        =   8
      Top             =   1020
      Width           =   270
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Age"
      Height          =   195
      Left            =   2760
      TabIndex        =   7
      Top             =   1020
      Width           =   285
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "D o B"
      Height          =   195
      Left            =   540
      TabIndex        =   6
      Top             =   1020
      Width           =   405
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "A and E"
      Height          =   195
      Left            =   2490
      TabIndex        =   5
      Top             =   600
      Width           =   570
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Chart"
      Height          =   195
      Left            =   570
      TabIndex        =   4
      Top             =   600
      Width           =   375
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      Height          =   195
      Left            =   525
      TabIndex        =   3
      Top             =   210
      Width           =   420
   End
End
Attribute VB_Name = "frmOrderPhoresis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

10        Unload Me

End Sub

Private Sub cmdSave_Click()

          Dim tb As Recordset
          Dim sql As String
          Dim Code As String
          Dim PatientID As String
          Dim Albumin As String
          Dim IgA As String
          Dim IgG As String
          Dim IgM As String
          Dim TotalProtein As String

10        On Error GoTo cmdSave_Click_Error

20        TotalProtein = Space(5)
30        Albumin = Space(30)
40        IgA = Space(30)
50        IgG = Space(30)
60        IgM = Space(30)

          'Get Total Protien from BioResults
          'sql = "Select Result From ImmResults Where SampleID = '" & lblSampleID & "' And Code = '" & GetOptionSetting("BioTProtCode", "61") & "'"
          'Set tb = New Recordset
          'RecOpenClient 0, tb, sql
          'If Not tb.EOF Then
          '    TotalProtein = Left(tb!Result & Space(5), 5)
          'End If

          'Get Albumin, IgA, IgG, IgM from ImmResults
70        sql = "Select Code, Result From ImmResults Where SampleID = '" & lblSampleID & "' And Code In ('" & GetOptionSetting("ImmAlbuminCode", "46") & "'," & _
                "'" & GetOptionSetting("ImmIgACode", "2") & "'," & _
                "'" & GetOptionSetting("ImmIgGCode", "1") & "'," & _
                "'" & GetOptionSetting("ImmIgMCode", "3") & "'," & _
                "'" & GetOptionSetting("BioTProtCode", "61") & "')"
80        Set tb = New Recordset
90        RecOpenClient 0, tb, sql
100       If Not tb.EOF Then
110           While Not tb.EOF
120               If tb!Code = "46" Then
130                   Albumin = Left(tb!Result & Space(30), 30)
140               ElseIf tb!Code = "1" Then
150                   IgG = Left(tb!Result & Space(30), 30)
160               ElseIf tb!Code = "2" Then
170                   IgA = Left(tb!Result & Space(30), 30)
180               ElseIf tb!Code = "3" Then
190                   IgM = Left(tb!Result & Space(30), 30)
200               ElseIf tb!Code = "61" Then
210                   TotalProtein = Left(tb!Result & Space(5), 5)
220               End If
230               tb.MoveNext
240           Wend
250       End If

260       If (IgA = "" Or IgG = "" Or IgM = "" Or TotalProtein = "") Or _
             (Not IsNumeric(IgG) And Not CBool(InStr(1, IgG, "<")) And Not CBool(InStr(1, IgG, ">"))) Or _
             (Not IsNumeric(IgA) And Not CBool(InStr(1, IgA, "<")) And Not CBool(InStr(1, IgA, ">"))) Or _
             (Not IsNumeric(IgM) And Not CBool(InStr(1, IgM, "<")) And Not CBool(InStr(1, IgM, ">"))) Or _
             (Not IsNumeric(TotalProtein) And Not CBool(InStr(1, TotalProtein, "<")) And Not CBool(InStr(1, TotalProtein, ">"))) Then
270           iMsg "one or more parameters values are wrong" & vbCrLf & _
                   "Current parameters values are: " & vbCrLf & _
                   "IgG           : " & IgG & vbCrLf & _
                   "IgA           : " & IgA & vbCrLf & _
                   "IgM           : " & IgM & vbCrLf & _
                   "Total Protein : " & TotalProtein

280           Exit Sub
290       End If

300       Code = ListCodeFor("IC", lstOrder.Text)

310       If Trim$(lblAandE) <> "" Then
320           PatientID = lblAandE
330       Else
340           PatientID = lblChart
350       End If

360       sql = "SELECT * FROM PhoresisRequests WHERE " & _
                "PatientID = '" & PatientID & "' " & _
                "AND PatientName = '" & AddTicks(lblName) & "' " & _
                "AND SampleDate = '" & Format(lblSampledate, "dd/MMM/yyyy") & "' " & _
                "AND Programmed = 0 " & _
                "AND AnalysisProgramCode = '" & Code & "'"
370       Set tb = New Recordset
380       RecOpenServer 0, tb, sql
390       If tb.EOF Then
400           tb.AddNew
410       End If

420       tb!AnalysisProgramCode = Code
430       tb!PhoresisSampleNumber = txtPhoresisNumber
440       tb!PatientID = PatientID
450       tb!PatientName = Left$(lblName, 30)
460       tb!Dob = Format$(IIf(lblDoB = "", "01/01/1900", lblDoB), "dd/MMM/yyyy")
470       tb!sex = Left$(lblSex, 1)
480       tb!Age = Left$(Format$(Val(lblAge)), 3)
490       tb!Department = lblClinician
500       tb!SampleDate = Format$(lblSampledate, "dd/MMM/yyyy")
510       tb!Username = Username
520       tb!Programmed = 0
530       tb!DateTimeOfRecord = Format$(Now, "dd/MMM/yyyy HH:nn:ss")
540       tb!SampleID = lblSampleID
550       tb!Concentration = CDbl(Trim(TotalProtein))
560       tb!Albumin = Albumin
570       tb!IgA = IgA
580       tb!IgG = IgG
590       tb!IgM = IgM
600       tb.Update

610       Unload Me

620       Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

630       intEL = Erl
640       strES = Err.Description
650       LogError "frmOrderPhoresis", "cmdSave_Click", intEL, strES, sql


End Sub

Private Sub Form_Load()

10        PopulateTestCodes

End Sub


Private Sub PopulateTestCodes()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo PopulateTestCodes_Error

20        sql = "Select Text From Lists Where ListType = 'IC'"
30        Set tb = New Recordset
40        RecOpenClient 0, tb, sql

50        If Not tb.EOF Then
60            lstOrder.Clear
70            While Not tb.EOF
80                lstOrder.AddItem tb!Text
90                tb.MoveNext
100           Wend
110       End If

120       Exit Sub

PopulateTestCodes_Error:

          Dim strES As String
          Dim intEL As Integer

130       intEL = Erl
140       strES = Err.Description
150       LogError "frmOrderPhoresis", "PopulateTestCodes", intEL, strES, sql

End Sub




Private Sub lstOrder_Click()

10        cmdSave.Visible = True

End Sub


Private Sub udPhoresisNumber_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        txtPhoresisNumber = Format$(txtPhoresisNumber, "0000")

End Sub

