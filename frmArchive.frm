VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmArchive 
   Caption         =   "NetAcquire - Archive"
   ClientHeight    =   8235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14430
   Icon            =   "frmArchive.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8235
   ScaleWidth      =   14430
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   615
      Left            =   3195
      Picture         =   "frmArchive.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1035
      Width           =   1605
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   690
      Left            =   3195
      Picture         =   "frmArchive.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   225
      Width           =   1590
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   1500
      Left            =   6480
      TabIndex        =   3
      Top             =   90
      Width           =   6000
      Begin VB.CommandButton cmdImmunology 
         Caption         =   "&Immunology"
         Height          =   315
         Left            =   3090
         TabIndex        =   24
         Top             =   150
         Width           =   1305
      End
      Begin VB.CommandButton cmdEndocrinology 
         Caption         =   "&Endocrinology"
         Height          =   315
         Left            =   3090
         TabIndex        =   23
         Top             =   480
         Width           =   1305
      End
      Begin VB.CommandButton cmdBiochemistry 
         Caption         =   "&Biochemistry"
         Height          =   315
         Left            =   180
         TabIndex        =   22
         Top             =   510
         Width           =   1065
      End
      Begin VB.OptionButton optCh 
         Caption         =   "Biochemistry"
         Height          =   240
         Index           =   14
         Left            =   1710
         TabIndex        =   19
         Tag             =   "BioRepeats"
         Top             =   1170
         Width           =   1275
      End
      Begin VB.OptionButton optCh 
         Caption         =   "Haematology"
         Height          =   240
         Index           =   13
         Left            =   1710
         TabIndex        =   18
         Tag             =   "HaemRepeats"
         Top             =   900
         Width           =   1275
      End
      Begin VB.OptionButton optCh 
         Caption         =   "Coagulation"
         Height          =   240
         Index           =   12
         Left            =   3105
         TabIndex        =   17
         Tag             =   "CoagRepeats"
         Top             =   900
         Width           =   1275
      End
      Begin VB.OptionButton optCh 
         Caption         =   "Immunology"
         Height          =   240
         Index           =   11
         Left            =   4455
         TabIndex        =   16
         Tag             =   "ImmRepeats"
         Top             =   900
         Width           =   1275
      End
      Begin VB.OptionButton optCh 
         Caption         =   "Endocinology"
         Height          =   240
         Index           =   10
         Left            =   3105
         TabIndex        =   15
         Tag             =   "EndRepeats"
         Top             =   1170
         Width           =   1275
      End
      Begin VB.OptionButton optCh 
         Caption         =   "Blood Gas"
         Height          =   240
         Index           =   8
         Left            =   4455
         TabIndex        =   14
         Tag             =   "BGAResults"
         Top             =   1170
         Width           =   1275
      End
      Begin VB.OptionButton optCh 
         Caption         =   "Blood Gas"
         Height          =   240
         Index           =   7
         Left            =   4455
         TabIndex        =   11
         Tag             =   "BGAResults"
         Top             =   495
         Width           =   1275
      End
      Begin VB.OptionButton optCh 
         Caption         =   "Externals"
         Height          =   240
         Index           =   6
         Left            =   4455
         TabIndex        =   10
         Tag             =   "ExtResults"
         Top             =   225
         Width           =   1275
      End
      Begin VB.OptionButton optCh 
         Caption         =   "Endocinology"
         Height          =   240
         Index           =   5
         Left            =   1350
         TabIndex        =   9
         Tag             =   "EndResults"
         Top             =   1290
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.OptionButton optCh 
         Caption         =   "Immunology"
         Height          =   240
         Index           =   4
         Left            =   2310
         TabIndex        =   8
         Tag             =   "ImmResults"
         Top             =   1320
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.OptionButton optCh 
         Caption         =   "Coagulation"
         Height          =   240
         Index           =   3
         Left            =   1710
         TabIndex        =   7
         Tag             =   "CoagResults"
         Top             =   510
         Width           =   1275
      End
      Begin VB.OptionButton optCh 
         Caption         =   "Haematology"
         Height          =   240
         Index           =   2
         Left            =   1710
         TabIndex        =   6
         Tag             =   "HaemResults"
         Top             =   225
         Width           =   1275
      End
      Begin VB.OptionButton optCh 
         Caption         =   "Biochemistry"
         Height          =   240
         Index           =   1
         Left            =   60
         TabIndex        =   5
         Tag             =   "BioResults"
         Top             =   1290
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.OptionButton optCh 
         Caption         =   "Demographics"
         Height          =   240
         Index           =   0
         Left            =   135
         TabIndex        =   4
         Tag             =   "Demographics"
         Top             =   225
         Value           =   -1  'True
         Width           =   1590
      End
      Begin VB.Label Label2 
         Caption         =   "Repeats"
         Height          =   285
         Left            =   810
         TabIndex        =   20
         Top             =   990
         Width           =   780
      End
   End
   Begin VB.TextBox txtSampleId 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1305
      TabIndex        =   1
      Top             =   360
      Width           =   1770
   End
   Begin MSFlexGridLib.MSFlexGrid grdArchive 
      Height          =   3030
      Left            =   135
      TabIndex        =   0
      Top             =   1755
      Width           =   14010
      _ExtentX        =   24712
      _ExtentY        =   5345
      _Version        =   393216
      Cols            =   200
      FixedCols       =   0
      AllowUserResizing=   1
   End
   Begin MSFlexGridLib.MSFlexGrid grdArc 
      Height          =   3030
      Left            =   180
      TabIndex        =   13
      Top             =   4995
      Width           =   14010
      _ExtentX        =   24712
      _ExtentY        =   5345
      _Version        =   393216
      Cols            =   200
      FixedCols       =   0
      AllowUserResizing=   1
   End
   Begin VB.Label Label1 
      Caption         =   "Sample Id"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   135
      TabIndex        =   2
      Top             =   405
      Width           =   1095
   End
End
Attribute VB_Name = "frmArchive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Table As String

Private Sub cmdBiochemistry_Click()

10        With frmAudit
20            .TableName = "BioResults"
30            .SampleID = txtSampleID
40            .Show 1
50        End With

End Sub

Private Sub cmdCancel_Click()

10        Unload Me

End Sub

Private Sub cmdEndocrinology_Click()

10        With frmAudit
20            .TableName = "EndResults"
30            .SampleID = txtSampleID
40            .Show 1
50        End With

End Sub

Private Sub cmdImmunology_Click()

10        With frmAudit
20            .TableName = "ImmResults"
30            .SampleID = txtSampleID
40            .Show 1
50        End With

End Sub


Private Sub cmdStart_Click()

          Dim tb As New Recordset
          Dim sql As String
          Dim f As Field
          Dim Nums As Long
          Dim Str As String
          Dim Num As Long
          Dim intS As Long
          Dim n As Long

10        On Error GoTo cmdStart_Click_Error

20        If Not IsNumeric(txtSampleID) Then
30            txtSampleID = ""
40            Exit Sub
50        End If

60        If Val(txtSampleID) = 0 Then
70            txtSampleID = ""
80            Exit Sub
90        End If

100       ClearFGrid grdArchive
110       ClearFGrid grdArc

120       grdArchive.LeftCol = 0
130       grdArc.LeftCol = 0
140       intS = 0
150       Nums = 0

160       sql = "SELECT * from " & Table & " WHERE sampleid = '" & txtSampleID & "'"
170       Set tb = New Recordset
180       RecOpenServer 0, tb, sql

190       For Each f In tb.Fields
200           Nums = Nums + 1
210           grdArchive.TextMatrix(0, Nums) = f.Name
220       Next

230       sql = "SELECT * from Arc" & Table & " WHERE sampleid = '" & txtSampleID & "'"
240       Set tb = New Recordset
250       RecOpenServer 0, tb, sql

260       For Each f In tb.Fields
270           intS = intS + 1
280           grdArc.TextMatrix(0, intS) = f.Name
290       Next

300       For Num = 1 To Nums
310           grdArchive.ColWidth(Num) = 1000
320       Next


330       For Num = 1 To intS
340           grdArc.ColWidth(Num) = 1000
350       Next

360       For Num = Nums + 1 To grdArchive.Cols - 1
370           grdArchive.ColWidth(Num) = 0
380       Next

390       For Num = intS + 1 To grdArc.Cols - 1
400           grdArc.ColWidth(Num) = 0
410       Next

420       sql = "SELECT * from " & Table & " WHERE sampleid = '" & txtSampleID & "'"
430       Set tb = New Recordset
440       RecOpenServer 0, tb, sql

450       Do While Not tb.EOF
460           Str = "Current" & vbTab
470           For Num = 1 To Nums
480               Str = Str & tb(grdArchive.TextMatrix(0, Num)) & vbTab
490           Next
500           grdArchive.AddItem Str
510           tb.MoveNext
520       Loop

530       sql = "SELECT * from Arc" & Table & " WHERE sampleid = '" & txtSampleID & "' "
540       Set tb = New Recordset
550       RecOpenServer 0, tb, sql

560       Do While Not tb.EOF
570           Str = vbTab
580           For Num = 1 To intS
590               Str = Str & tb(grdArc.TextMatrix(0, Num)) & vbTab
600           Next
610           grdArc.AddItem Str
620           tb.MoveNext
630       Loop

640       For n = 0 To grdArc.Cols - 1
650           grdArc.TextMatrix(0, n) = grdArc.TextMatrix(0, n)
660       Next

670       For n = 0 To grdArchive.Cols - 1
680           grdArchive.TextMatrix(0, n) = grdArchive.TextMatrix(0, n)
690       Next

700       FixG grdArchive

710       FixG grdArc

720       Exit Sub

cmdStart_Click_Error:

          Dim strES As String
          Dim intEL As Integer

730       intEL = Erl
740       strES = Err.Description
750       LogError "frmSystemArchive", "cmdStart_Click", intEL, strES, sql

End Sub

Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        Table = optCh(0).Tag

30        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmArchive", "Form_Load", intEL, strES


End Sub

Private Sub grdArc_Scroll()


10        On Error GoTo grdArc_Scroll_Error

20        grdArchive.LeftCol = grdArc.LeftCol

30        Exit Sub

grdArc_Scroll_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmArchive", "grdArc_Scroll", intEL, strES


End Sub

Private Sub grdArchive_Scroll()

10        On Error GoTo grdArchive_Scroll_Error

20        grdArc.LeftCol = grdArchive.LeftCol

30        Exit Sub

grdArchive_Scroll_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmArchive", "grdArchive_Scroll", intEL, strES


End Sub

Private Sub optCh_Click(Index As Integer)

10        On Error GoTo optCh_Click_Error

20        Table = optCh(Index).Tag

30        Exit Sub

optCh_Click_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmArchive", "optCh_Click", intEL, strES


End Sub
