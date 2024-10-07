VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmArchiveMicro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Microbiology Archive"
   ClientHeight    =   9240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14505
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   14505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Department"
      Height          =   765
      Left            =   13050
      TabIndex        =   18
      Top             =   90
      Width           =   1185
      Begin VB.OptionButton optMS 
         Caption         =   "Semen"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   20
         Top             =   480
         Width           =   795
      End
      Begin VB.OptionButton optMS 
         Caption         =   "Micro"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   19
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.CommandButton bPrint 
      Caption         =   "&Print to Default Printer"
      Height          =   1140
      Left            =   13140
      Picture         =   "frmArchiveMicro.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      Tag             =   "bprint"
      Top             =   6420
      Width           =   1170
   End
   Begin RichTextLib.RichTextBox rtb 
      Height          =   8115
      Left            =   300
      TabIndex        =   11
      Top             =   900
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   14314
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmArchiveMicro.frx":1982
   End
   Begin VB.Frame Frame1 
      Caption         =   "Department"
      Height          =   795
      Left            =   4620
      TabIndex        =   10
      Top             =   60
      Width           =   8265
      Begin VB.OptionButton optDept 
         Alignment       =   1  'Right Justify
         Caption         =   "Comments"
         Enabled         =   0   'False
         Height          =   195
         Index           =   10
         Left            =   5400
         TabIndex        =   21
         Top             =   510
         Width           =   1095
      End
      Begin VB.OptionButton optDept 
         Caption         =   "Demographics"
         Enabled         =   0   'False
         Height          =   195
         Index           =   9
         Left            =   6540
         TabIndex        =   17
         Top             =   270
         Width           =   1335
      End
      Begin VB.OptionButton optDept 
         Caption         =   "Semen"
         Enabled         =   0   'False
         Height          =   195
         Index           =   11
         Left            =   6540
         TabIndex        =   15
         Top             =   510
         Width           =   795
      End
      Begin VB.OptionButton optDept 
         Alignment       =   1  'Right Justify
         Caption         =   "Faecal Requests"
         Enabled         =   0   'False
         Height          =   195
         Index           =   7
         Left            =   2280
         TabIndex        =   14
         Top             =   510
         Width           =   1545
      End
      Begin VB.OptionButton optDept 
         Alignment       =   1  'Right Justify
         Caption         =   "Urine Requests"
         Enabled         =   0   'False
         Height          =   195
         Index           =   6
         Left            =   2400
         TabIndex        =   13
         Top             =   270
         Width           =   1425
      End
      Begin VB.OptionButton optDept 
         Caption         =   "Sensitivities"
         Enabled         =   0   'False
         Height          =   195
         Index           =   5
         Left            =   3870
         TabIndex        =   12
         Top             =   510
         Width           =   1125
      End
      Begin VB.OptionButton optDept 
         Alignment       =   1  'Right Justify
         Caption         =   "Identification"
         Enabled         =   0   'False
         Height          =   195
         Index           =   8
         Left            =   5250
         TabIndex        =   8
         Top             =   270
         Width           =   1245
      End
      Begin VB.OptionButton optDept 
         Alignment       =   1  'Right Justify
         Caption         =   "Isolates"
         Enabled         =   0   'False
         Height          =   195
         Index           =   4
         Left            =   360
         TabIndex        =   3
         Top             =   270
         Width           =   885
      End
      Begin VB.OptionButton optDept 
         Caption         =   "Generic"
         Enabled         =   0   'False
         Height          =   195
         Index           =   3
         Left            =   3870
         TabIndex        =   7
         Top             =   270
         Width           =   915
      End
      Begin VB.OptionButton optDept 
         Caption         =   "Urine"
         Enabled         =   0   'False
         Height          =   195
         Index           =   2
         Left            =   1260
         TabIndex        =   5
         Top             =   270
         Width           =   675
      End
      Begin VB.OptionButton optDept 
         Alignment       =   1  'Right Justify
         Caption         =   "Site Details"
         Enabled         =   0   'False
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   4
         Top             =   510
         Width           =   1095
      End
      Begin VB.OptionButton optDept 
         Caption         =   "Faeces"
         Enabled         =   0   'False
         Height          =   195
         Index           =   0
         Left            =   1260
         TabIndex        =   6
         Top             =   510
         Width           =   825
      End
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Default         =   -1  'True
      Height          =   780
      Left            =   2340
      Picture         =   "frmArchiveMicro.frx":1A11
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   60
      Width           =   960
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   1140
      Left            =   13140
      Picture         =   "frmArchiveMicro.frx":3393
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7890
      Width           =   1170
   End
   Begin VB.TextBox txtSampleID 
      Height          =   315
      Left            =   300
      TabIndex        =   0
      Top             =   300
      Width           =   1935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Sample ID"
      Height          =   195
      Left            =   330
      TabIndex        =   9
      Top             =   90
      Width           =   735
   End
End
Attribute VB_Name = "frmArchiveMicro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pSampleIDNoOffset As Double

Private mOffset As Double
Private Sub DoArcCommentsUserName()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo DoArcCommentsUserName_Error

20        sql = "SELECT TOP 1 UserName FROM CommentsArc WHERE " & _
                "SampleID = '" & Format$(Val(txtSampleID) + mOffset) & "' " & _
                "ORDER BY DateTimeOfArchive ASC"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If Not tb.EOF Then
60            rtb.SelText = "Original Entry by " & tb!Username & vbCrLf & vbCrLf
70        End If

80        Exit Sub

DoArcCommentsUserName_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmArchiveMicro", "DoArcCommentsUserName", intEL, strES, sql

End Sub

Private Sub DoCommentsUserName()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo DoCommentsUserName_Error

20        sql = "SELECT TOP 1 UserName FROM Comments WHERE SampleID = '" & Format$(Val(txtSampleID) + mOffset) & "'"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If Not tb.EOF Then
60            rtb.SelText = "Original Entry by " & tb!Username & vbCrLf & vbCrLf
70        End If

80        Exit Sub

DoCommentsUserName_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmArchiveMicro", "DoCommentsUserName", intEL, strES, sql

End Sub

Private Sub DoComments()

          Dim tb As Recordset
          Dim sql As String
          Dim SID As Double
          Dim LatestValue(1 To 7) As String
          Dim dbName(1 To 7) As String
          Dim ShowName(1 To 7) As String
          Dim n As Integer
          Dim X As Integer

10        On Error GoTo DoComments_Error

          'SampleID, Demographic, Biochemistry, Haematology, Coagulation, Immunology,
          'BloodGas, Semen, Microcs, MicroIdent, MicroGeneral, MicroConsultant,
          'Film , Endocrinology, Histology, Cytology, rowguid, CSFFluid,
          'ImmunologyA , ImmunologyB, ImmunologyC, UserName, DateTimeOfRecord

20        For n = 1 To 7
30            dbName(n) = Choose(n, "Demographic", "Semen", "MicroCS", "MicroIdent", _
                                 "MicroGeneral", "MicroConsultant", "CSFFluid")
40            ShowName(n) = Choose(n, "Demographic Comment", "Semen Comment", "Culture Comment", "Identification Comment", _
                                   "General Comment", "Consultant Comment", "CSF Fluid Comment")
50        Next

60        SID = Val(txtSampleID) + mOffset

70        sql = "SELECT " & _
                "CASE UserName WHEN '' THEN '<BLANK>' ELSE ISNULL(UserName, '<BLANK>') END UserName, " & _
                "CASE CAST(Demographic AS nvarchar(4000)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Demographic, '<BLANK>') END Demographic, " & _
                "CASE CAST(Semen AS nvarchar(4000)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Semen, '<BLANK>') END Semen, " & _
                "CASE CAST(MicroCS AS nvarchar(4000)) WHEN '' THEN '<BLANK>' ELSE ISNULL(MicroCS, '<BLANK>') END MicroCS, " & _
                "CASE CAST(MicroIdent AS nvarchar(4000)) WHEN '' THEN '<BLANK>' ELSE ISNULL(MicroIdent, '<BLANK>') END MicroIdent, " & _
                "CASE CAST(MicroGeneral AS nvarchar(4000)) WHEN '' THEN '<BLANK>' ELSE ISNULL(MicroGeneral, '<BLANK>') END MicroGeneral, " & _
                "CASE CAST(MicroConsultant AS nvarchar(4000)) WHEN '' THEN '<BLANK>' ELSE ISNULL(MicroConsultant, '<BLANK>') END MicroConsultant, " & _
                "CASE CAST(CSFFluid AS nvarchar(4000)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CSFFluid, '<BLANK>') END CSFFluid " & _
                "FROM Comments WHERE " & _
                "SampleID = " & SID & ""
80        Set tb = New Recordset
90        RecOpenServer 0, tb, sql
100       If Not tb.EOF Then
110           For n = 1 To 7
120               LatestValue(n) = tb(dbName(n)) & ""
130           Next
140       Else
150           For n = 1 To 7
160               LatestValue(n) = "<BLANK>"
170           Next
180       End If

190       sql = "SELECT DateTimeOfArchive, " & _
                "CASE ArchivedBy WHEN '' THEN '<BLANK>' ELSE ISNULL(ArchivedBy, '<BLANK>') END ArchivedBy, " & _
                "CASE UserName WHEN '' THEN '<BLANK>' ELSE ISNULL(UserName, '<BLANK>') END UserName, " & _
                "CASE LTRIM(RTRIM(CAST(Demographic AS nvarchar(4000)))) WHEN '' THEN '<BLANK>' ELSE ISNULL(Demographic, '<BLANK>') END Demographic, " & _
                "CASE LTRIM(RTRIM(CAST(Semen AS nvarchar(4000)))) WHEN '' THEN '<BLANK>' ELSE ISNULL(Semen, '<BLANK>') END Semen, " & _
                "CASE LTRIM(RTRIM(CAST(MicroCS AS nvarchar(4000)))) WHEN '' THEN '<BLANK>' ELSE ISNULL(MicroCS, '<BLANK>') END MicroCS, " & _
                "CASE LTRIM(RTRIM(CAST(MicroIdent AS nvarchar(4000)))) WHEN '' THEN '<BLANK>' ELSE ISNULL(MicroIdent, '<BLANK>') END MicroIdent, " & _
                "CASE LTRIM(RTRIM(CAST(MicroGeneral AS nvarchar(4000)))) WHEN '' THEN '<BLANK>' ELSE ISNULL(MicroGeneral, '<BLANK>') END MicroGeneral, " & _
                "CASE LTRIM(RTRIM(CAST(MicroConsultant AS nvarchar(4000)))) WHEN '' THEN '<BLANK>' ELSE ISNULL(MicroConsultant, '<BLANK>') END MicroConsultant, " & _
                "CASE LTRIM(RTRIM(CAST(CSFFluid AS nvarchar(4000)))) WHEN '' THEN '<BLANK>' ELSE ISNULL(CSFFluid, '<BLANK>') END CSFFluid " & _
                "FROM CommentsArc WHERE " & _
                "SampleID = " & SID & " " & _
                "ORDER BY DateTimeOfArchive DESC"
200       Set tb = New Recordset
210       RecOpenServer 0, tb, sql
220       Do While Not tb.EOF
230           rtb.SelColor = vbBlack
240           If Not IsNull(tb!DateTimeOfArchive) Then
250               rtb.SelText = Format$(tb!DateTimeOfArchive, "dd/MM/yy HH:mm:ss")
260           End If
270           For X = 1 To 7
280               If LatestValue(X) <> tb(dbName(X)) & "" Then
290                   rtb.SelText = " " & ShowName(X)
300                   rtb.SelText = " changed from "
310                   rtb.SelColor = vbRed
320                   rtb.SelText = tb.Fields(dbName(X))
330                   rtb.SelColor = vbBlack
340                   rtb.SelText = " to "
350                   rtb.SelColor = vbBlue
360                   rtb.SelText = LatestValue(X)
370                   rtb.SelColor = vbBlack
380                   rtb.SelText = " by "
390                   rtb.SelColor = vbGreen
400                   rtb.SelText = tb!ArchivedBy
410                   rtb.SelColor = vbBlack
420                   rtb.SelText = vbCrLf
430               End If
440           Next
450           For X = 1 To 7
460               LatestValue(X) = tb(dbName(X))
470           Next
480           rtb.SelText = vbCrLf
490           tb.MoveNext
500       Loop

510       Exit Sub

DoComments_Error:

          Dim strES As String
          Dim intEL As Integer

520       intEL = Erl
530       strES = Err.Description
540       LogError "frmArchiveMicro", "DoComments", intEL, strES, sql

End Sub


Private Sub DoFaeces()

          Dim tb As Recordset
          Dim sql As String
          Dim SID As Double
          Dim LatestValue(1 To 12) As String
          Dim dbName(1 To 12) As String
          Dim ShowName(1 To 12) As String
          Dim n As Integer
          Dim X As Integer

10        On Error GoTo DoFaeces_Error

20        For n = 1 To 12
30            dbName(n) = Choose(n, "Rota", "Adeno", "OP0", "OP1", "OP2", "Valid", _
                                 "OB0", "OB1", "OB2", "ToxinAB", "Cryptosporidium", "HPylori")
40            ShowName(n) = Choose(n, "Rota Virus", "Adeno Virus", "Ova/Parasites(1)", "Ova/Parasites(2)", "Ova/Parasites(3)", "Valid", _
                                   "Occult Blood(1)", "Occult Blood(2)", "Occult Blood(3)", "Toxin AB", "Cryptosporidium", "HPylori")
50        Next

60        SID = Val(txtSampleID) + mOffset

70        sql = "SELECT " & _
                "CASE UserName WHEN '' THEN '<BLANK>' ELSE ISNULL(UserName, '<BLANK>') END UserName, " & _
                "CASE LTRIM(RTRIM(Rota)) WHEN '' THEN '<BLANK>' WHEN 'N' THEN 'Negative' " & _
                "     WHEN 'P' THEN 'Positive' WHEN 'I' THEN 'Inconclusive' WHEN 'U' THEN 'Insufficient Sample' " & _
                "     ELSE ISNULL(CAST(Rota AS nvarchar(50)), '<BLANK>') END Rota, " & _
                "CASE LTRIM(RTRIM(Adeno)) WHEN '' THEN '<BLANK>' WHEN 'N' THEN 'Negative' " & _
                "     WHEN 'P' THEN 'Positive' WHEN 'I' THEN 'Inconclusive' WHEN 'U' THEN 'Insufficient Sample' " & _
                "     ELSE ISNULL(CAST(Adeno AS nvarchar(50)), '<BLANK>') END Adeno, " & _
                "CASE LTRIM(RTRIM(CAST(OP0 AS nvarchar(50)))) WHEN '' THEN '<BLANK>' ELSE ISNULL(OP0, '<BLANK>') END OP0, " & _
                "CASE LTRIM(RTRIM(CAST(OP1 AS nvarchar(50)))) WHEN '' THEN '<BLANK>' ELSE ISNULL(OP1, '<BLANK>') END OP1, " & _
                "CASE LTRIM(RTRIM(CAST(OP2 AS nvarchar(50)))) WHEN '' THEN '<BLANK>' ELSE ISNULL(OP2, '<BLANK>') END OP2, " & _
                "CASE LTRIM(RTRIM(OB0)) WHEN '' THEN '<BLANK>' WHEN 'N' THEN 'Negative' WHEN 'P' THEN 'Positive' WHEN 'U' THEN 'Insufficient Sample' ELSE ISNULL(CAST(OB0 AS nvarchar(50)), '<BLANK>') END OB0, " & _
                "CASE LTRIM(RTRIM(OB1)) WHEN '' THEN '<BLANK>' WHEN 'N' THEN 'Negative' WHEN 'P' THEN 'Positive' WHEN 'U' THEN 'Insufficient Sample' ELSE ISNULL(CAST(OB1 AS nvarchar(50)), '<BLANK>') END OB1, " & _
                "CASE LTRIM(RTRIM(OB2)) WHEN '' THEN '<BLANK>' WHEN 'N' THEN 'Negative' WHEN 'P' THEN 'Positive' WHEN 'U' THEN 'Insufficient Sample' ELSE ISNULL(CAST(OB2 AS nvarchar(50)), '<BLANK>') END OB2, " & _
                "CASE LTRIM(RTRIM(ToxinAB)) WHEN '' THEN '<BLANK>' WHEN 'N' THEN 'Not Detected' " & _
                "     WHEN 'P' THEN 'Positive' WHEN 'I' THEN 'Inconclusive' WHEN 'R' THEN 'Sample Rejected' " & _
                "     ELSE ISNULL(CAST(ToxinAB AS nvarchar(50)), '<BLANK>') END ToxinAB, " & _
                "CASE LTRIM(RTRIM(Cryptosporidium)) WHEN '' THEN '<BLANK>' WHEN 'N' THEN 'Not detected' " & _
                "     WHEN 'O' THEN 'Oocysts detected' WHEN 'I' THEN 'Inconclusive' WHEN 'U' then 'Insufficient Sample' " & _
                "     ELSE ISNULL(Cryptosporidium, '<BLANK>') END Cryptosporidium, " & _
                "CASE LTRIM(RTRIM(HPylori)) WHEN '' THEN '<BLANK>' ELSE ISNULL(HPylori, '<BLANK>') END HPylori, " & _
                "CASE Valid WHEN 1 THEN 'Valid' ELSE 'Not Valid' END Valid " & _
                "FROM Faeces WHERE " & _
                "SampleID = " & SID
80        Set tb = New Recordset
90        RecOpenServer 0, tb, sql
100       If Not tb.EOF Then
110           For n = 1 To 12
120               LatestValue(n) = tb(dbName(n)) & ""
130           Next

140           sql = "SELECT UserName, ArchiveDateTime, ArchivedBy, " & _
                    "CASE LTRIM(RTRIM(Rota)) WHEN '' THEN '<BLANK>' WHEN 'N' THEN 'Negative' " & _
                    "     WHEN 'P' THEN 'Positive' WHEN 'I' THEN 'Inconclusive' WHEN 'U' THEN 'Insufficient Sample' " & _
                    "     ELSE ISNULL(CAST(Rota AS nvarchar(50)), '<BLANK>') END Rota, " & _
                    "CASE LTRIM(RTRIM(Adeno)) WHEN '' THEN '<BLANK>' WHEN 'N' THEN 'Negative' " & _
                    "     WHEN 'P' THEN 'Positive' WHEN 'I' THEN 'Inconclusive' WHEN 'U' THEN 'Insufficient Sample' " & _
                    "     ELSE ISNULL(CAST(Adeno AS nvarchar(50)), '<BLANK>') END Adeno, " & _
                    "CASE LTRIM(RTRIM(CAST(OP0 AS nvarchar(50)))) WHEN '' THEN '<BLANK>' ELSE ISNULL(OP0, '<BLANK>') END OP0, " & _
                    "CASE LTRIM(RTRIM(CAST(OP1 AS nvarchar(50)))) WHEN '' THEN '<BLANK>' ELSE ISNULL(OP1, '<BLANK>') END OP1, " & _
                    "CASE LTRIM(RTRIM(CAST(OP2 AS nvarchar(50)))) WHEN '' THEN '<BLANK>' ELSE ISNULL(OP2, '<BLANK>') END OP2, " & _
                    "CASE LTRIM(RTRIM(OB0)) WHEN '' THEN '<BLANK>' WHEN 'N' THEN 'Negative' WHEN 'P' THEN 'Positive' WHEN 'U' THEN 'Insufficient Sample' ELSE ISNULL(CAST(OB0 AS nvarchar(50)), '<BLANK>') END OB0, " & _
                    "CASE LTRIM(RTRIM(OB1)) WHEN '' THEN '<BLANK>' WHEN 'N' THEN 'Negative' WHEN 'P' THEN 'Positive' WHEN 'U' THEN 'Insufficient Sample' ELSE ISNULL(CAST(OB1 AS nvarchar(50)), '<BLANK>') END OB1, " & _
                    "CASE LTRIM(RTRIM(OB2)) WHEN '' THEN '<BLANK>' WHEN 'N' THEN 'Negative' WHEN 'P' THEN 'Positive' WHEN 'U' THEN 'Insufficient Sample' ELSE ISNULL(CAST(OB2 AS nvarchar(50)), '<BLANK>') END OB2, " & _
                    "CASE LTRIM(RTRIM(ToxinAB)) WHEN '' THEN '<BLANK>' WHEN 'N' THEN 'Not Detected' " & _
                    "     WHEN 'P' THEN 'Positive' WHEN 'I' THEN 'Inconclusive' WHEN 'R' THEN 'Sample Rejected' " & _
                    "     ELSE ISNULL(CAST(ToxinAB AS nvarchar(50)), '<BLANK>') END ToxinAB, " & _
                    "CASE LTRIM(RTRIM(Cryptosporidium)) WHEN '' THEN '<BLANK>' WHEN 'N' THEN 'Not detected' " & _
                    "     WHEN 'O' THEN 'Oocysts detected' WHEN 'I' THEN 'Inconclusive' WHEN 'U' then 'Insufficient Sample' " & _
                    "     ELSE ISNULL(Cryptosporidium, '<BLANK>') END Cryptosporidium, " & _
                    "CASE LTRIM(RTRIM(HPylori)) WHEN '' THEN '<BLANK>' ELSE ISNULL(HPylori, '<BLANK>') END HPylori, " & _
                    "CASE Valid WHEN 1 THEN 'Valid' ELSE 'Not Valid' END Valid " & _
                    "FROM FaecesArc WHERE " & _
                    "SampleID = " & SID & " " & _
                    "ORDER BY ArchiveDateTime DESC"
150           Set tb = New Recordset
160           RecOpenServer 0, tb, sql
170           Do While Not tb.EOF
180               rtb.SelColor = vbBlack
190               If Not IsNull(tb!ArchiveDateTime) Then
200                   rtb.SelText = Format$(tb!ArchiveDateTime, "dd/MM/yy HH:mm:ss")
210               End If
220               rtb.SelText = " : " & vbCrLf
230               For X = 1 To 12
240                   If LatestValue(X) <> tb(dbName(X)) & "" Then
250                       rtb.SelText = ShowName(X)
260                       rtb.SelText = " : "
270                       rtb.SelColor = vbRed
280                       rtb.SelText = tb!Username & ""
290                       rtb.SelColor = vbBlack
300                       rtb.SelText = " entered "
310                       rtb.SelColor = vbRed
320                       rtb.SelText = tb.Fields(dbName(X))
330                       rtb.SelColor = vbBlack
340                       rtb.SelText = ". Changed to "
350                       rtb.SelColor = vbBlue
360                       rtb.SelText = LatestValue(X)
370                       rtb.SelColor = vbBlack
380                       rtb.SelText = " by "
390                       rtb.SelColor = vbGreen
400                       rtb.SelText = tb!ArchivedBy
410                       rtb.SelColor = vbBlack
420                       rtb.SelText = vbCrLf
430                   End If
440               Next

                  '240       For x = 1 To 12
                  '250         If LatestValue(x) <> tb(dbName(x)) & "" Then
                  '260           rtb.SelText = ShowName(x)
                  '270           rtb.SelText = " changed from "
                  '280           rtb.SelColor = vbRed
                  '290           rtb.SelText = tb.Fields(dbName(x))
                  '300           rtb.SelColor = vbBlack
                  '310           rtb.SelText = " to "
                  '320           rtb.SelColor = vbBlue
                  '330           rtb.SelText = LatestValue(x)
                  '340           rtb.SelColor = vbBlack
                  '350           rtb.SelText = " by "
                  '360           rtb.SelColor = vbGreen
                  '370           rtb.SelText = tb!ArchivedBy
                  '380           rtb.SelColor = vbBlack
                  '390           rtb.SelText = vbCrLf
                  '400         End If
                  '410       Next


450               For X = 1 To 12
460                   LatestValue(X) = tb(dbName(X))
470               Next
480               rtb.SelText = vbCrLf
490               tb.MoveNext
500           Loop
510       End If

520       Exit Sub

DoFaeces_Error:

          Dim strES As String
          Dim intEL As Integer

530       intEL = Erl
540       strES = Err.Description
550       LogError "frmArchiveMicro", "DoFaeces", intEL, strES, sql


End Sub
Private Sub DoArcFaecesUserName()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo DoArcFaecesUserName_Error

20        sql = "SELECT TOP 1 UserName FROM FaecesArc WHERE " & _
                "SampleID = '" & Format$(Val(txtSampleID) + mOffset) & "' " & _
                "ORDER BY ArchiveDateTime ASC"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If Not tb.EOF Then
60            rtb.SelText = "Original Entry by " & tb!Username & vbCrLf & vbCrLf
70        End If

80        Exit Sub

DoArcFaecesUserName_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmArchiveMicro", "DoArcFaecesUserName", intEL, strES, sql

End Sub

Private Sub DoFaecesUserName()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo DoFaecesUserName_Error

20        sql = "SELECT TOP 1 UserName FROM Faeces WHERE " & _
                "SampleID = '" & Format$(Val(txtSampleID) + mOffset) & "' "
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If Not tb.EOF Then
60            rtb.SelText = "Original Entry by " & tb!Username & vbCrLf & vbCrLf
70        End If

80        Exit Sub

DoFaecesUserName_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmArchiveMicro", "DoFaecesUserName", intEL, strES, sql

End Sub

Private Sub DoUrineUserName()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo DoUrineUserName_Error

20        sql = "SELECT UserName FROM Urine WHERE SampleID = '" & Format$(Val(txtSampleID) + mOffset) & "'"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If Not tb.EOF Then
60            rtb.SelText = "Original Entry by " & tb!Username & vbCrLf & vbCrLf
70        End If

80        Exit Sub

DoUrineUserName_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmArchiveMicro", "DoUrineUserName", intEL, strES, sql

End Sub


Private Sub DoArcUrineUserName()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo DoArcUrineUserName_Error

20        sql = "SELECT TOP 1 UserName FROM UrineArc WHERE " & _
                "SampleID = '" & Format$(Val(txtSampleID) + mOffset) & "' " & _
                "ORDER BY ArchiveDateTime ASC"

30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If Not tb.EOF Then
60            rtb.SelText = "Original Entry by " & tb!Username & vbCrLf & vbCrLf
70        End If

80        Exit Sub

DoArcUrineUserName_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmArchiveMicro", "DoArcUrineUserName", intEL, strES, sql

End Sub

Private Sub DoGenericUserName()

          Dim tb As Recordset
          Dim sql As String
          Dim s As String

10        On Error GoTo DoGenericUserName_Error

20        sql = "SELECT DISTINCT UserName FROM GenericResults WHERE " & _
                "SampleID = '" & Format$(Val(txtSampleID) + mOffset) & "'"
30        Set tb = New Recordset
40        RecOpenClient 0, tb, sql
50        If Not tb.EOF Then
60            tb.MoveLast
70            If tb.RecordCount > 1 Then
80                s = "Multiple original entries :- "
90                Do While Not tb.BOF
100                   s = s & tb!Username & ", "
110                   tb.MovePrevious
120               Loop
130               s = Left(s, Len(s) - 2)
140           Else
150               s = "Original Entry by " & tb!Username & ""
160           End If
170           rtb.SelText = s & vbCrLf & vbCrLf
180       End If

190       Exit Sub

DoGenericUserName_Error:

          Dim strES As String
          Dim intEL As Integer

200       intEL = Erl
210       strES = Err.Description
220       LogError "frmArchiveMicro", "DoGenericUserName", intEL, strES, sql

End Sub

Private Sub DoArcGenericUserName()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo DoArcGenericUserName_Error

20        sql = "SELECT TOP 1 UserName FROM GenericResultsArc WHERE " & _
                "SampleID = '" & Format$(Val(txtSampleID) + mOffset) & "' " & _
                "ORDER BY ArchiveDateTime ASC"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If Not tb.EOF Then
60            rtb.SelText = "Original Entry by " & tb!Username & vbCrLf & vbCrLf
70        End If

80        Exit Sub

DoArcGenericUserName_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmArchiveMicro", "DoArcGenericUserName", intEL, strES, sql

End Sub

Private Sub DoIsolatesUserName()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo DoIsolatesUserName_Error

20        sql = "SELECT TOP 1 UserName FROM Isolates WHERE SampleID = '" & Format$(Val(txtSampleID) + mOffset) & "'"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If Not tb.EOF Then
60            rtb.SelText = "Original Entry by " & tb!Username & vbCrLf & vbCrLf
70        End If

80        Exit Sub

DoIsolatesUserName_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmArchiveMicro", "DoIsolatesUserName", intEL, strES, sql

End Sub


Private Sub DoArcIsolatesUserName()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo DoArcIsolatesUserName_Error

20        sql = "SELECT TOP 1 UserName FROM IsolatesArc WHERE " & _
                "SampleID = '" & Format$(Val(txtSampleID) + mOffset) & "' " & _
                "ORDER BY ArchiveDateTime ASC"

30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If Not tb.EOF Then
60            rtb.SelText = "Original Entry by " & tb!Username & vbCrLf & vbCrLf
70        End If

80        Exit Sub

DoArcIsolatesUserName_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmArchiveMicro", "DoArcIsolatesUserName", intEL, strES, sql

End Sub



Private Sub DoUrineRequestsUserName()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo DoUrineRequestsUserName_Error

20        sql = "SELECT TOP 1 UserName FROM UrineRequests WHERE SampleID = '" & Format$(Val(txtSampleID) + mOffset) & "'"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If Not tb.EOF Then
60            rtb.SelText = "Original Entry by " & tb!Username & vbCrLf & vbCrLf
70        End If

80        Exit Sub

DoUrineRequestsUserName_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmArchiveMicro", "DoUrineRequestsUserName", intEL, strES, sql

End Sub

Private Sub DoArcUrineRequestsUserName()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo DoArcUrineRequestsUserName_Error

20        sql = "SELECT TOP 1 UserName FROM UrineRequestsArc WHERE " & _
                "SampleID = '" & Format$(Val(txtSampleID) + mOffset) & "' " & _
                "ORDER BY ArchiveDateTime ASC"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If Not tb.EOF Then
60            rtb.SelText = "Original Entry by " & tb!Username & vbCrLf & vbCrLf
70        End If

80        Exit Sub

DoArcUrineRequestsUserName_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmArchiveMicro", "DoArcUrineRequestsUserName", intEL, strES, sql

End Sub


Private Sub DoDemographicsOperator()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo DoDemographicsOperator_Error

20        sql = "SELECT TOP 1 Operator FROM Demographics WHERE " & _
                "SampleID = '" & Format$(Val(txtSampleID) + mOffset) & "'"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If Not tb.EOF Then
60            rtb.SelText = "Original Entry by " & tb!Operator & vbCrLf & vbCrLf
70        End If

80        Exit Sub

DoDemographicsOperator_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmArchiveMicro", "DoDemographicsOperator", intEL, strES, sql

End Sub

Private Sub DoArcDemographicsOperator()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo DoArcDemographicsOperator_Error

20        sql = "SELECT TOP 1 Operator FROM ArcDemographics WHERE " & _
                "SampleID = '" & Format$(Val(txtSampleID) + mOffset) & "' " & _
                "ORDER BY DateTimeOfArchive ASC"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If Not tb.EOF Then
60            rtb.SelText = "Original Entry by " & tb!Operator & vbCrLf & vbCrLf
70        End If

80        Exit Sub

DoArcDemographicsOperator_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmArchiveMicro", "DoArcDemographicsOperator", intEL, strES, sql

End Sub

Private Sub DoIdentUserName()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo DoIdentUserName_Error

20        sql = "SELECT TOP 1 UserName FROM UrineIdent WHERE SampleID = '" & Format$(Val(txtSampleID) + mOffset) & "'"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If Not tb.EOF Then
60            rtb.SelText = "Original Entry by " & tb!Username & vbCrLf & vbCrLf
70        End If

80        Exit Sub

DoIdentUserName_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmArchiveMicro", "DoIdentUserName", intEL, strES, sql

End Sub


Private Sub DoArcIdentUserName()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo DoArcIdentUserName_Error

20        sql = "SELECT TOP 1 UserName FROM UrineIdentArc WHERE " & _
                "SampleID = '" & Format$(Val(txtSampleID) + mOffset) & "' " & _
                "ORDER BY ArchiveDateTime ASC"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If Not tb.EOF Then
60            rtb.SelText = "Original Entry by " & tb!Username & vbCrLf & vbCrLf
70        End If

80        Exit Sub

DoArcIdentUserName_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmArchiveMicro", "DoArcIdentUserName", intEL, strES, sql

End Sub



Private Sub DoFaecalRequestsUserName()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo DoFaecalRequestsUserName_Error

20        sql = "SELECT TOP 1 UserName FROM FaecalRequests WHERE SampleID = '" & Format$(Val(txtSampleID) + mOffset) & "'"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If Not tb.EOF Then
60            rtb.SelText = "Original Entry by " & tb!Username & vbCrLf & vbCrLf
70        End If

80        Exit Sub

DoFaecalRequestsUserName_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmArchiveMicro", "DoFaecalRequestsUserName", intEL, strES, sql

End Sub


Private Sub DoArcFaecalRequestsUserName()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo DoArcFaecalRequestsUserName_Error

20        sql = "SELECT TOP 1 UserName FROM FaecalRequestsArc WHERE " & _
                "SampleID = '" & Format$(Val(txtSampleID) + mOffset) & "' " & _
                "ORDER BY ArchiveDateTime ASC"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If Not tb.EOF Then
60            rtb.SelText = "Original Entry by " & tb!Username & vbCrLf & vbCrLf
70        End If

80        Exit Sub

DoArcFaecalRequestsUserName_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmArchiveMicro", "DoArcFaecalRequestsUserName", intEL, strES, sql

End Sub

Private Sub DoSemenUserName()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo DoSemenUserName_Error

20        sql = "SELECT TOP 1 UserName FROM SemenResults WHERE SampleID = '" & Format$(Val(txtSampleID) + mOffset) & "'"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If Not tb.EOF Then
60            rtb.SelText = "Original Entry by " & tb!Username & vbCrLf & vbCrLf
70        End If

80        Exit Sub

DoSemenUserName_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmArchiveMicro", "DoSemenUserName", intEL, strES, sql

End Sub



Private Sub DoArcSemenUserName()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo DoArcSemenUserName_Error

20        sql = "SELECT TOP 1 UserName FROM SemenResultsArc WHERE " & _
                "SampleID = '" & Format$(Val(txtSampleID) + mOffset) & "' " & _
                "ORDER BY ArchiveDateTime ASC"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If Not tb.EOF Then
60            rtb.SelText = "Original Entry by " & tb!Username & vbCrLf & vbCrLf
70        End If

80        Exit Sub

DoArcSemenUserName_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmArchiveMicro", "DoArcSemenUserName", intEL, strES, sql

End Sub


Private Sub DoSensitivitiesUserCode()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo DoSensitivitiesUserCode_Error

20        sql = "SELECT TOP 1 U.Name UserName FROM Sensitivities S, Users U WHERE S.SampleID = '" & Format$(Val(txtSampleID) + mOffset) & "' AND U.Code = S.UserCode"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If Not tb.EOF Then
60            rtb.SelText = "Original Entry by " & tb!Username & vbCrLf & vbCrLf
70        End If

80        Exit Sub

DoSensitivitiesUserCode_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmArchiveMicro", "DoSensitivitiesUserCode", intEL, strES, sql

End Sub


Private Sub DoArcSensitivitiesUserCode()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo DoArcSensitivitiesUserCode_Error

20        sql = "SELECT TOP 1 U.Name FROM Users U, SensitivitiesArc S WHERE " & _
                "S.SampleID = '" & Format$(Val(txtSampleID) + mOffset) & "' " & _
                "AND S.UserCode = U.Code " & _
                "ORDER BY S.ArchiveDateTime ASC"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If Not tb.EOF Then
60            rtb.SelText = "Original Entry by " & tb!Name & vbCrLf & vbCrLf
70        End If

80        Exit Sub

DoArcSensitivitiesUserCode_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmArchiveMicro", "DoArcSensitivitiesUserCode", intEL, strES, sql

End Sub

Private Sub DoSiteDetailsUserName()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo DoSiteDetailsUserName_Error

20        sql = "SELECT UserName FROM MicroSiteDetails WHERE SampleID = '" & Format$(Val(txtSampleID) + mOffset) & "'"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If Not tb.EOF Then
60            rtb.SelText = "Original Entry by " & tb!Username & vbCrLf & vbCrLf
70        End If

80        Exit Sub

DoSiteDetailsUserName_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmArchiveMicro", "DoSiteDetailsUserName", intEL, strES, sql

End Sub



Private Sub DoArcSiteDetailsUserName()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo DoArcSiteDetailsUserName_Error

20        sql = "SELECT TOP 1 UserName FROM MicroSiteDetailsArc WHERE " & _
                "SampleID = '" & Format$(Val(txtSampleID) + mOffset) & "' " & _
                "ORDER BY ArchiveDateTime ASC"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If Not tb.EOF Then
60            rtb.SelText = "Original Entry by " & tb!Username & vbCrLf & vbCrLf
70        End If

80        Exit Sub

DoArcSiteDetailsUserName_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmArchiveMicro", "DoArcSiteDetailsUserName", intEL, strES, sql

End Sub


Private Sub DoGeneric()

          Dim tb As Recordset
          Dim sql As String
          Dim SID As Double
          Dim LatestValue As String
          Dim TestNames As New Collection
          Dim Y As Integer

10        On Error GoTo DoGeneric_Error

20        SID = Val(txtSampleID) + mOffset

30        sql = "SELECT DISTINCT(TestName) FROM GenericResultsArc WHERE " & _
                "SampleID = '" & SID & "' "
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        Do While Not tb.EOF
70            TestNames.Add tb!TestName & ""
80            tb.MoveNext
90        Loop

100       For Y = 1 To TestNames.Count
110           sql = "SELECT " & _
                    "CASE UserName WHEN '' THEN '<BLANK>' ELSE ISNULL(UserName, '<BLANK>') END UserName, " & _
                    "CASE LTRIM(RTRIM(CAST(TestName AS nvarchar(50)))) WHEN '' THEN '<BLANK>' ELSE ISNULL(TestName, '<BLANK>') END TestName, " & _
                    "CASE LTRIM(RTRIM(CAST(Result AS nvarchar(50)))) WHEN '' THEN '<BLANK>' ELSE ISNULL(Result, '<BLANK>') END Result " & _
                    "FROM GenericResults WHERE " & _
                    "SampleID = '" & SID & "' " & _
                    "AND TestName = '" & TestNames(Y) & "'"
120           Set tb = New Recordset
130           RecOpenServer 0, tb, sql
140           If Not tb.EOF Then
150               LatestValue = tb!Result & ""
160           Else
170               LatestValue = "<BLANK>"
180           End If
190           sql = "SELECT ArchiveDateTime, " & _
                    "CASE ArchivedBy WHEN '' THEN '<BLANK>' ELSE ISNULL(ArchivedBy, '<BLANK>') END ArchivedBy, " & _
                    "CASE LTRIM(RTRIM(CAST(TestName AS nvarchar(50)))) WHEN '' THEN '<BLANK>' ELSE ISNULL(TestName, '<BLANK>') END TestName, " & _
                    "CASE LTRIM(RTRIM(CAST(Result AS nvarchar(50)))) WHEN '' THEN '<BLANK>' ELSE ISNULL(Result, '<BLANK>') END Result " & _
                    "FROM GenericResultsArc WHERE " & _
                    "SampleID = '" & SID & "' " & _
                    "AND TestName = '" & TestNames(Y) & "'" & _
                    "ORDER BY ArchiveDateTime DESC"
200           Set tb = New Recordset
210           RecOpenServer 0, tb, sql
220           Do While Not tb.EOF
230               rtb.SelColor = vbBlack
240               If Not IsNull(tb!ArchiveDateTime) Then
250                   rtb.SelText = Format$(tb!ArchiveDateTime, "dd/MM/yy HH:mm:ss")
260               End If
270               rtb.SelText = " : " & vbCrLf
280               rtb.SelText = TestNames(Y)
290               rtb.SelText = " changed from "
300               rtb.SelColor = vbRed
310               rtb.SelText = tb!Result & ""
320               rtb.SelColor = vbBlack
330               rtb.SelText = " to "
340               rtb.SelColor = vbBlue
350               rtb.SelText = LatestValue
360               rtb.SelColor = vbBlack
370               rtb.SelText = " by "
380               rtb.SelColor = vbGreen
390               rtb.SelText = tb!ArchivedBy
400               rtb.SelColor = vbBlack
410               rtb.SelText = vbCrLf

420               LatestValue = tb!Result & ""
430               rtb.SelText = vbCrLf
440               tb.MoveNext
450           Loop

460       Next

470       Exit Sub

DoGeneric_Error:

          Dim strES As String
          Dim intEL As Integer

480       intEL = Erl
490       strES = Err.Description
500       LogError "frmArchiveMicro", "DoGeneric", intEL, strES, sql


End Sub
Private Sub DoGenericSemen()

          Dim tb As Recordset
          Dim sql As String
          Dim SID As Double
          Dim LatestValue As String
          Dim TestNames As New Collection
          Dim Y As Integer

10        On Error GoTo DoGenericSemen_Error

20        SID = Val(txtSampleID) + mOffset

30        sql = "SELECT DISTINCT(TestName) FROM GenericResultsArc WHERE " & _
                "SampleID = '" & SID & "' "
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        Do While Not tb.EOF
70            TestNames.Add tb!TestName & ""
80            tb.MoveNext
90        Loop

100       For Y = 1 To TestNames.Count
110           sql = "SELECT " & _
                    "CASE UserName WHEN '' THEN '<BLANK>' ELSE ISNULL(UserName, '<BLANK>') END UserName, " & _
                    "CASE LTRIM(RTRIM(CAST(TestName AS nvarchar(50)))) WHEN '' THEN '<BLANK>' ELSE ISNULL(TestName, '<BLANK>') END TestName, " & _
                    "CASE LTRIM(RTRIM(CAST(Result AS nvarchar(50)))) WHEN '' THEN '<BLANK>' ELSE ISNULL(Result, '<BLANK>') END Result " & _
                    "FROM GenericResults WHERE " & _
                    "SampleID = '" & SID & "' " & _
                    "AND TestName = '" & TestNames(Y) & "'"
120           Set tb = New Recordset
130           RecOpenServer 0, tb, sql
140           If Not tb.EOF Then
150               LatestValue = tb!Result & ""

160               sql = "SELECT ArchiveDateTime, " & _
                        "CASE ArchivedBy WHEN '' THEN '<BLANK>' ELSE ISNULL(ArchivedBy, '<BLANK>') END ArchivedBy, " & _
                        "CASE LTRIM(RTRIM(CAST(TestName AS nvarchar(50)))) WHEN '' THEN '<BLANK>' ELSE ISNULL(TestName, '<BLANK>') END TestName, " & _
                        "CASE LTRIM(RTRIM(CAST(Result AS nvarchar(50)))) WHEN '' THEN '<BLANK>' ELSE ISNULL(Result, '<BLANK>') END Result " & _
                        "FROM GenericResultsArc WHERE " & _
                        "SampleID = '" & SID & "' " & _
                        "AND TestName = '" & TestNames(Y) & "'" & _
                        "ORDER BY ArchiveDateTime DESC"
170               Set tb = New Recordset
180               RecOpenServer 0, tb, sql
190               Do While Not tb.EOF
200                   rtb.SelColor = vbBlack
210                   If Not IsNull(tb!ArchiveDateTime) Then
220                       rtb.SelText = Format$(tb!ArchiveDateTime, "dd/MM/yy HH:mm:ss")
230                   End If
240                   rtb.SelText = " : " & vbCrLf
250                   rtb.SelText = TestNames(Y)
260                   rtb.SelText = " changed from "
270                   rtb.SelColor = vbRed
280                   rtb.SelText = tb!Result & ""
290                   rtb.SelColor = vbBlack
300                   rtb.SelText = " to "
310                   rtb.SelColor = vbBlue
320                   rtb.SelText = LatestValue
330                   rtb.SelColor = vbBlack
340                   rtb.SelText = " by "
350                   rtb.SelColor = vbGreen
360                   rtb.SelText = tb!ArchivedBy
370                   rtb.SelColor = vbBlack
380                   rtb.SelText = vbCrLf

390                   LatestValue = tb!Result & ""
400                   rtb.SelText = vbCrLf
410                   tb.MoveNext
420               Loop
430           End If
440       Next

450       Exit Sub

DoGenericSemen_Error:

          Dim strES As String
          Dim intEL As Integer

460       intEL = Erl
470       strES = Err.Description
480       LogError "frmArchiveMicro", "DoGenericSemen", intEL, strES, sql


End Sub
Private Sub DoIdent()

          Dim tb As Recordset
          Dim sql As String
          Dim SID As Double
          Dim LatestValue(1 To 14) As String
          Dim dbName(1 To 14) As String
          Dim ShowName(1 To 14) As String
          Dim n As Integer
          Dim X As Integer

          Dim IsolateNumber As Integer

10        On Error GoTo DoIdent_Error

20        For n = 1 To 14
30            dbName(n) = Choose(n, "Gram", "ZN", "WetPrep", "Indole", "Coagulase", "Catalase", "Oxidase", "Rapidec", "Chromogenic", _
                                 "Reincubation", "UrineSensitivity", "ExtraSensitivity", "Notes", "Valid")
40            ShowName(n) = Choose(n, "Gram Stain", "ZN Stain", "Wet Prep", "Indole", "Coagulase", "Catalase", "Oxidase", "Rapidec", "Chromogenic", _
                                   "Reincubation", "Urine Sensitivity", "Extra Sensitivity", "Notes", "Valid")

50        Next

60        SID = Val(txtSampleID) + mOffset

70        For IsolateNumber = 1 To 4
80            sql = "SELECT " & _
                    "CASE UserName WHEN '' THEN '<BLANK>' ELSE ISNULL(UserName, '<BLANK>') END UserName, " & _
                    "CASE LTRIM(RTRIM(Gram)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Gram, '<BLANK>') END Gram, " & _
                    "CASE LTRIM(RTRIM(ZN)) WHEN '' THEN '<BLANK>' ELSE ISNULL(ZN, '<BLANK>') END ZN, " & _
                    "CASE LTRIM(RTRIM(WetPrep)) WHEN '' THEN '<BLANK>' ELSE ISNULL(WetPrep, '<BLANK>') END WetPrep, " & _
                    "CASE LTRIM(RTRIM(Indole)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Indole, '<BLANK>') END Indole, " & _
                    "CASE LTRIM(RTRIM(Coagulase)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Coagulase, '<BLANK>') END Coagulase, " & _
                    "CASE LTRIM(RTRIM(Catalase)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Catalase, '<BLANK>') END Catalase, " & _
                    "CASE LTRIM(RTRIM(Oxidase)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Oxidase, '<BLANK>') END Oxidase, " & _
                    "CASE LTRIM(RTRIM(Rapidec)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Rapidec, '<BLANK>') END Rapidec, " & _
                    "CASE LTRIM(RTRIM(Chromogenic)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Chromogenic, '<BLANK>') END Chromogenic, " & _
                    "CASE LTRIM(RTRIM(Reincubation)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Reincubation, '<BLANK>') END Reincubation, " & _
                    "CASE LTRIM(RTRIM(UrineSensitivity)) WHEN '' THEN '<BLANK>' ELSE ISNULL(UrineSensitivity, '<BLANK>') END UrineSensitivity, " & _
                    "CASE LTRIM(RTRIM(ExtraSensitivity)) WHEN '' THEN '<BLANK>' ELSE ISNULL(ExtraSensitivity, '<BLANK>') END ExtraSensitivity, " & _
                    "CASE LTRIM(RTRIM(Notes)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Notes, '<BLANK>') END Notes, " & _
                    "CASE Valid WHEN 1 THEN 'Valid' ELSE 'Not Valid' END Valid " & _
                    "FROM UrineIdent WHERE " & _
                    "SampleID = '" & SID & "' " & _
                    "AND Isolate = " & IsolateNumber
90            Set tb = New Recordset
100           RecOpenServer 0, tb, sql
110           If Not tb.EOF Then
120               For n = 1 To 14
130                   LatestValue(n) = tb(dbName(n)) & ""
140               Next
150           Else
160               For n = 1 To 14
170                   LatestValue(n) = "<BLANK>"
180               Next
190           End If

200           sql = "SELECT ArchiveDateTime, " & _
                    "CASE ArchivedBy WHEN '' THEN '<BLANK>' ELSE ISNULL(ArchivedBy, '<BLANK>') END ArchivedBy, " & _
                    "CASE LTRIM(RTRIM(Gram)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Gram, '<BLANK>') END Gram, " & _
                    "CASE LTRIM(RTRIM(ZN)) WHEN '' THEN '<BLANK>' ELSE ISNULL(ZN, '<BLANK>') END ZN, " & _
                    "CASE LTRIM(RTRIM(WetPrep)) WHEN '' THEN '<BLANK>' ELSE ISNULL(WetPrep, '<BLANK>') END WetPrep, " & _
                    "CASE LTRIM(RTRIM(Indole)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Indole, '<BLANK>') END Indole, " & _
                    "CASE LTRIM(RTRIM(Coagulase)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Coagulase, '<BLANK>') END Coagulase, " & _
                    "CASE LTRIM(RTRIM(Catalase)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Catalase, '<BLANK>') END Catalase, " & _
                    "CASE LTRIM(RTRIM(Oxidase)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Oxidase, '<BLANK>') END Oxidase, " & _
                    "CASE LTRIM(RTRIM(Rapidec)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Rapidec, '<BLANK>') END Rapidec, " & _
                    "CASE LTRIM(RTRIM(Chromogenic)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Chromogenic, '<BLANK>') END Chromogenic, " & _
                    "CASE LTRIM(RTRIM(Reincubation)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Reincubation, '<BLANK>') END Reincubation, " & _
                    "CASE LTRIM(RTRIM(UrineSensitivity)) WHEN '' THEN '<BLANK>' ELSE ISNULL(UrineSensitivity, '<BLANK>') END UrineSensitivity, " & _
                    "CASE LTRIM(RTRIM(ExtraSensitivity)) WHEN '' THEN '<BLANK>' ELSE ISNULL(ExtraSensitivity, '<BLANK>') END ExtraSensitivity, " & _
                    "CASE LTRIM(RTRIM(Notes)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Notes, '<BLANK>') END Notes, " & _
                    "CASE Valid WHEN 1 THEN 'Valid' ELSE 'Not Valid' END Valid " & _
                    "FROM UrineIdentArc WHERE " & _
                    "SampleID = '" & SID & "' " & _
                    "AND Isolate = " & IsolateNumber & " " & _
                    "ORDER BY ArchiveDateTime DESC"
210           Set tb = New Recordset
220           RecOpenServer 0, tb, sql
230           Do While Not tb.EOF
240               rtb.SelColor = vbBlack
250               If Not IsNull(tb!ArchiveDateTime) Then
260                   rtb.SelText = Format$(tb!ArchiveDateTime, "dd/MM/yy HH:mm:ss")
270               End If
280               rtb.SelText = " : Isolate " & Format$(IsolateNumber) & vbCrLf
290               For X = 1 To 14
300                   If LatestValue(X) <> tb(dbName(X)) & "" Then
310                       rtb.SelText = ShowName(X)
320                       rtb.SelText = " changed from "
330                       rtb.SelColor = vbRed
340                       rtb.SelText = tb.Fields(dbName(X))
350                       rtb.SelColor = vbBlack
360                       rtb.SelText = " to "
370                       rtb.SelColor = vbBlue
380                       rtb.SelText = LatestValue(X)
390                       rtb.SelColor = vbBlack
400                       rtb.SelText = " by "
410                       rtb.SelColor = vbGreen
420                       rtb.SelText = tb!ArchivedBy
430                       rtb.SelColor = vbBlack
440                       rtb.SelText = vbCrLf
450                   End If
460               Next
470               For X = 1 To 14
480                   LatestValue(X) = tb(dbName(X))
490               Next
500               rtb.SelText = vbCrLf
510               tb.MoveNext
520           Loop

530       Next

540       Exit Sub

DoIdent_Error:

          Dim strES As String
          Dim intEL As Integer

550       intEL = Erl
560       strES = Err.Description
570       LogError "frmArchiveMicro", "DoIdent", intEL, strES, sql

End Sub

Private Sub DoDemographics()

          Dim tb As Recordset
          Dim sql As String
          Dim SID As Double
          Dim LatestValue(1 To 21) As String
          Dim dbName(1 To 21) As String
          Dim ShowName(1 To 21) As String
          Dim n As Integer
          Dim X As Integer

10        On Error GoTo DoDemographics_Error

20        For n = 1 To 21
30            dbName(n) = Choose(n, "PatName", "Age", "Sex", "RunDate", "DoB", "Addr0", "Addr1", _
                                 "Ward", "Clinician", "GP", "SampleDate", "ClDetails", "Hospital", "RooH", _
                                 "AandE", "Chart", "MRN", "RecDate", "Valid", "Pregnant", "PenicillinAllergy")
40            ShowName(n) = Choose(n, "Patient Name", "Age", "Sex", "Run Date", "Date of Birth", "Address (1)", "Address (2)", _
                                   "Ward", "Clinician", "GP", "Sample Date", "Clinical Details", "Hospital", "Routine", _
                                   "A and E", "Chart", "MRN", "Received Date", "Valid", "Pregnant", "Penicillin Allergy")
50        Next

60        SID = Val(txtSampleID) + mOffset


70        sql = "SELECT " & _
                "CASE LTRIM(RTRIM(PatName)) WHEN '' THEN '<BLANK>' ELSE ISNULL(PatName, '<BLANK>') END PatName, " & _
                "CASE LTRIM(RTRIM(Age)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Age, '<BLANK>') END Age, " & _
                "CASE LTRIM(RTRIM(Sex)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Sex, '<BLANK>') END Sex, " & _
                "CASE LTRIM(RTRIM(RunDate)) WHEN '' THEN '<BLANK>' ELSE ISNULL(RunDate, '<BLANK>') END RunDate, " & _
                "CASE LTRIM(RTRIM(DoB)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CONVERT(nvarchar(50), DoB, 103), '<BLANK>') END DoB, " & _
                "CASE LTRIM(RTRIM(Addr0)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Addr0, '<BLANK>') END Addr0, " & _
                "CASE LTRIM(RTRIM(Addr1)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Addr1, '<BLANK>') END Addr1, "
80        sql = sql & _
                "CASE LTRIM(RTRIM(Ward)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Ward, '<BLANK>') END Ward, " & _
                "CASE LTRIM(RTRIM(Clinician)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Clinician, '<BLANK>') END Clinician, " & _
                "CASE LTRIM(RTRIM(GP)) WHEN '' THEN '<BLANK>' ELSE ISNULL(GP, '<BLANK>') END GP, " & _
                "CASE LTRIM(RTRIM(SampleDate)) WHEN '' THEN '<BLANK>' ELSE ISNULL(SampleDate, '<BLANK>') END SampleDate, " & _
                "CASE LTRIM(RTRIM(ClDetails)) WHEN '' THEN '<BLANK>' ELSE ISNULL(ClDetails, '<BLANK>') END ClDetails, " & _
                "CASE LTRIM(RTRIM(Hospital)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Hospital, '<BLANK>') END Hospital, " & _
                "CASE RooH WHEN 1 THEN 'Routine' ELSE 'Out of Hours' END RooH, "
90        sql = sql & _
                "CASE LTRIM(RTRIM(AandE)) WHEN '' THEN '<BLANK>' ELSE ISNULL(LTRIM(RTRIM(AandE)), '<BLANK>') END AandE, " & _
                "CASE LTRIM(RTRIM(Chart)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Chart, '<BLANK>') END Chart, " & _
                "CASE LTRIM(RTRIM(MRN)) WHEN '' THEN '<BLANK>' ELSE ISNULL(MRN, '<BLANK>') END MRN, " & _
                "CASE LTRIM(RTRIM(RecDate)) WHEN '' THEN '<BLANK>' ELSE ISNULL(RecDate, '<BLANK>') END RecDate, " & _
                "CASE Valid WHEN 1 THEN 'Valid' ELSE 'Not Valid' END Valid, " & _
                "CASE Pregnant WHEN 1 THEN 'Pregnant' ELSE 'Not Pregnant' END Pregnant, " & _
                "CASE PenicillinAllergy WHEN 1 THEN 'Penicillin Allergy' ELSE 'No Penicillin Allergy' END PenicillinAllergy "
100       sql = sql & _
                "FROM Demographics WHERE " & _
                "SampleID = '" & SID & "'"
110       Set tb = New Recordset
120       RecOpenServer 0, tb, sql
130       If Not tb.EOF Then
140           For n = 1 To 21
150               LatestValue(n) = tb(dbName(n)) & ""
160           Next
170       Else
180           For n = 1 To 21
190               LatestValue(n) = "<BLANK>"
200           Next
210       End If

220       sql = "SELECT DateTimeOfArchive, " & _
                "CASE ArchiveOperator WHEN '' THEN '<BLANK>' ELSE ISNULL(ArchiveOperator, '<BLANK>') END ArchiveOperator, " & _
                "CASE LTRIM(RTRIM(PatName)) WHEN '' THEN '<BLANK>' ELSE ISNULL(PatName, '<BLANK>') END PatName, " & _
                "CASE LTRIM(RTRIM(Age)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Age, '<BLANK>') END Age, " & _
                "CASE LTRIM(RTRIM(Sex)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Sex, '<BLANK>') END Sex, " & _
                "CASE LTRIM(RTRIM(RunDate)) WHEN '' THEN '<BLANK>' ELSE ISNULL(RunDate, '<BLANK>') END RunDate, " & _
                "CASE LTRIM(RTRIM(DoB)) WHEN '' THEN '<BLANK>' ELSE ISNULL(CONVERT(nvarchar(50), DoB, 103), '<BLANK>') END DoB, " & _
                "CASE LTRIM(RTRIM(Addr0)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Addr0, '<BLANK>') END Addr0, " & _
                "CASE LTRIM(RTRIM(Addr1)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Addr1, '<BLANK>') END Addr1, "
230       sql = sql & _
                "CASE LTRIM(RTRIM(Ward)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Ward, '<BLANK>') END Ward, " & _
                "CASE LTRIM(RTRIM(Clinician)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Clinician, '<BLANK>') END Clinician, " & _
                "CASE LTRIM(RTRIM(GP)) WHEN '' THEN '<BLANK>' ELSE ISNULL(GP, '<BLANK>') END GP, " & _
                "CASE LTRIM(RTRIM(SampleDate)) WHEN '' THEN '<BLANK>' ELSE ISNULL(SampleDate, '<BLANK>') END SampleDate, " & _
                "CASE LTRIM(RTRIM(ClDetails)) WHEN '' THEN '<BLANK>' ELSE ISNULL(ClDetails, '<BLANK>') END ClDetails, " & _
                "CASE LTRIM(RTRIM(Hospital)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Hospital, '<BLANK>') END Hospital, " & _
                "CASE RooH WHEN 1 THEN 'Routine' ELSE 'Out of Hours' END RooH, "
240       sql = sql & _
                "CASE LTRIM(RTRIM(AandE)) WHEN '' THEN '<BLANK>' ELSE ISNULL(LTRIM(RTRIM(AandE)), '<BLANK>') END AandE, " & _
                "CASE LTRIM(RTRIM(Chart)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Chart, '<BLANK>') END Chart, " & _
                "CASE LTRIM(RTRIM(MRN)) WHEN '' THEN '<BLANK>' ELSE ISNULL(MRN, '<BLANK>') END MRN, " & _
                "CASE LTRIM(RTRIM(RecDate)) WHEN '' THEN '<BLANK>' ELSE ISNULL(RecDate, '<BLANK>') END RecDate, " & _
                "CASE Valid WHEN 1 THEN 'Valid' ELSE 'Not Valid' END Valid, " & _
                "CASE Pregnant WHEN 1 THEN 'Pregnant' ELSE 'Not Pregnant' END Pregnant, " & _
                "CASE PenicillinAllergy WHEN 1 THEN 'Penicillin Allergy' ELSE 'No Penicillin Allergy' END PenicillinAllergy "
250       sql = sql & _
                "FROM ArcDemographics WHERE " & _
                "SampleID = '" & SID & "' " & _
                "ORDER BY DateTimeOfArchive DESC"
260       Set tb = New Recordset
270       RecOpenServer 0, tb, sql
280       Do While Not tb.EOF
290           rtb.SelColor = vbBlack
300           If Not IsNull(tb!DateTimeOfArchive) Then
310               rtb.SelText = Format$(tb!DateTimeOfArchive, "dd/MM/yy HH:mm:ss")
320           Else
330               rtb.SelText = "Archive Time not known."
340           End If
350           rtb.SelText = vbCrLf
360           For X = 1 To 21
370               If LatestValue(X) <> tb(dbName(X)) & "" Then
380                   rtb.SelText = ShowName(X)
390                   rtb.SelText = " changed from "
400                   rtb.SelColor = vbRed
410                   rtb.SelText = tb.Fields(dbName(X))
420                   rtb.SelColor = vbBlack
430                   rtb.SelText = " to "
440                   rtb.SelColor = vbBlue
450                   rtb.SelText = LatestValue(X)
460                   rtb.SelColor = vbBlack
470                   rtb.SelText = " by "
480                   rtb.SelColor = vbGreen
490                   rtb.SelText = tb!ArchiveOperator
500                   rtb.SelColor = vbBlack
510                   rtb.SelText = vbCrLf
520               End If
530           Next
540           For X = 1 To 21
550               LatestValue(X) = tb(dbName(X))
560           Next
570           rtb.SelText = vbCrLf
580           tb.MoveNext
590       Loop

600       Exit Sub

DoDemographics_Error:

          Dim strES As String
          Dim intEL As Integer

610       intEL = Erl
620       strES = Err.Description
630       LogError "frmArchiveMicro", "DoDemographics", intEL, strES, sql

End Sub

Private Sub DoIsolates()

          Dim tb As Recordset
          Dim sql As String
          Dim SID As Double
          Dim LatestValue(1 To 4) As String
          Dim dbName(1 To 4) As String
          Dim ShowName(1 To 4) As String
          Dim n As Integer
          Dim X As Integer

          Dim IsolateNumber As Integer

10        On Error GoTo DoIsolates_Error

20        For n = 1 To 4
30            dbName(n) = Choose(n, "OrganismGroup", "OrganismName", "Qualifier", "Valid")
40            ShowName(n) = Choose(n, "Organism Group", "Organism Name", "Qualifier", "Valid")
50        Next

60        SID = Val(txtSampleID) + mOffset

70        For IsolateNumber = 1 To 4
80            sql = "SELECT " & _
                    "CASE UserName WHEN '' THEN '<BLANK>' ELSE ISNULL(UserName, '<BLANK>') END UserName, " & _
                    "CASE LTRIM(RTRIM(OrganismGroup)) WHEN '' THEN '<BLANK>' ELSE ISNULL(OrganismGroup, '<BLANK>') END OrganismGroup, " & _
                    "CASE LTRIM(RTRIM(OrganismName)) WHEN '' THEN '<BLANK>' ELSE ISNULL(OrganismName, '<BLANK>') END OrganismName, " & _
                    "CASE LTRIM(RTRIM(Qualifier)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Qualifier, '<BLANK>') END Qualifier, " & _
                    "CASE Valid WHEN 1 THEN 'Valid' ELSE 'Not Valid' END Valid " & _
                    "FROM Isolates WHERE " & _
                    "SampleID = '" & SID & "' " & _
                    "AND IsolateNumber = " & IsolateNumber
90            Set tb = New Recordset
100           RecOpenServer 0, tb, sql
110           If Not tb.EOF Then
120               For n = 1 To 4
130                   LatestValue(n) = tb(dbName(n)) & ""
140               Next

150               sql = "SELECT ArchiveDateTime, " & _
                        "CASE ArchivedBy WHEN '' THEN '<BLANK>' ELSE ISNULL(ArchivedBy, '<BLANK>') END ArchivedBy, " & _
                        "CASE LTRIM(RTRIM(OrganismGroup)) WHEN '' THEN '<BLANK>' ELSE ISNULL(OrganismGroup, '<BLANK>') END OrganismGroup, " & _
                        "CASE LTRIM(RTRIM(OrganismName)) WHEN '' THEN '<BLANK>' ELSE ISNULL(OrganismName, '<BLANK>') END OrganismName, " & _
                        "CASE LTRIM(RTRIM(Qualifier)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Qualifier, '<BLANK>') END Qualifier, " & _
                        "CASE Valid WHEN 1 THEN 'Valid' ELSE 'Not Valid' END Valid " & _
                        "FROM IsolatesArc WHERE " & _
                        "SampleID = '" & SID & "' " & _
                        "AND IsolateNumber = " & IsolateNumber & " " & _
                        "ORDER BY ArchiveDateTime DESC"
160               Set tb = New Recordset
170               RecOpenServer 0, tb, sql
180               Do While Not tb.EOF
190                   rtb.SelColor = vbBlack
200                   If Not IsNull(tb!ArchiveDateTime) Then
210                       rtb.SelText = Format$(tb!ArchiveDateTime, "dd/MM/yy HH:mm:ss")
220                   End If
230                   rtb.SelText = " : Isolate " & Format$(IsolateNumber) & vbCrLf
240                   For X = 1 To 4
250                       If LatestValue(X) <> tb(dbName(X)) & "" Then
260                           rtb.SelText = ShowName(X)
270                           rtb.SelText = " changed from "
280                           rtb.SelColor = vbRed
290                           rtb.SelText = tb.Fields(dbName(X))
300                           rtb.SelColor = vbBlack
310                           rtb.SelText = " to "
320                           rtb.SelColor = vbBlue
330                           rtb.SelText = LatestValue(X)
340                           rtb.SelColor = vbBlack
350                           rtb.SelText = " by "
360                           rtb.SelColor = vbGreen
370                           rtb.SelText = tb!ArchivedBy
380                           rtb.SelColor = vbBlack
390                           rtb.SelText = vbCrLf
400                       End If
410                   Next
420                   For X = 1 To 4
430                       LatestValue(X) = tb(dbName(X))
440                   Next
450                   rtb.SelText = vbCrLf
460                   tb.MoveNext
470               Loop
480           End If
490       Next

500       Exit Sub

DoIsolates_Error:

          Dim strES As String
          Dim intEL As Integer

510       intEL = Erl
520       strES = Err.Description
530       LogError "frmArchiveMicro", "DoIsolates", intEL, strES, sql

End Sub

Private Sub DoSensitivities()

          Dim tb As Recordset
          Dim sql As String
          Dim SID As Double
          Dim LatestValue As String
          Dim AntiBsName As Collection
          Dim AntiBsCode As Collection
          Dim Y As Integer

          Dim IsolateNumber As Integer

10        On Error GoTo DoSensitivities_Error

20        SID = Val(txtSampleID) + mOffset

30        For IsolateNumber = 1 To 4
40            Set AntiBsName = New Collection
50            Set AntiBsCode = New Collection
60            sql = "SELECT DISTINCT(AntibioticCode) FROM SensitivitiesArc WHERE " & _
                    "SampleID = '" & SID & "' "
70            Set tb = New Recordset
80            RecOpenServer 0, tb, sql
90            Do While Not tb.EOF
100               AntiBsCode.Add Trim$(tb!AntibioticCode & "")
110               AntiBsName.Add Trim$(AntibioticNameFor(Trim$(tb!AntibioticCode & "")))
120               tb.MoveNext
130           Loop

140           For Y = 1 To AntiBsName.Count
150               sql = "SELECT " & _
                        "CASE UserCode WHEN '' THEN '<BLANK>' ELSE ISNULL(UserCode, '<BLANK>') END UserCode, " & _
                        "CASE LTRIM(RTRIM(RSI)) WHEN '' THEN '<BLANK>' ELSE ISNULL(RSI, '<BLANK>') END RSI " & _
                        "FROM Sensitivities WHERE " & _
                        "SampleID = '" & SID & "' " & _
                        "AND IsolateNumber = " & IsolateNumber & " " & _
                        "AND AntibioticCode = '" & AntiBsCode(Y) & "'"
160               Set tb = New Recordset
170               RecOpenServer 0, tb, sql
180               If Not tb.EOF Then
190                   LatestValue = Trim$(tb!RSI & "")
200               End If

210               sql = "SELECT ArchiveDateTime, " & _
                        "CASE ArchivedBy WHEN '' THEN '<BLANK>' ELSE ISNULL(ArchivedBy, '<BLANK>') END ArchivedBy, " & _
                        "CASE LTRIM(RTRIM(RSI)) WHEN '' THEN '<BLANK>' ELSE ISNULL(RSI, '<BLANK>') END RSI " & _
                        "FROM SensitivitiesArc WHERE " & _
                        "SampleID = '" & SID & "' " & _
                        "AND IsolateNumber = " & IsolateNumber & " " & _
                        "AND AntibioticCode = '" & AntiBsCode(Y) & "' " & _
                        "ORDER BY ArchiveDateTime DESC"
220               Set tb = New Recordset
230               RecOpenServer 0, tb, sql
240               Do While Not tb.EOF
250                   rtb.SelColor = vbBlack
260                   If Not IsNull(tb!ArchiveDateTime) Then
270                       rtb.SelText = Format$(tb!ArchiveDateTime, "dd/MM/yy HH:mm:ss")
280                   End If
290                   rtb.SelText = " : " & vbCrLf
300                   If LatestValue <> tb!RSI & "" Then
310                       rtb.SelText = AntiBsName(Y)
320                       rtb.SelText = " changed from "
330                       rtb.SelColor = vbRed
340                       If Trim$(tb!RSI & "") = "R" Then
350                           rtb.SelText = "Resistant"
360                       ElseIf Trim$(tb!RSI & "") = "S" Then
370                           rtb.SelText = "Sensitive"
380                       ElseIf Trim$(tb!RSI & "") = "I" Then
390                           rtb.SelText = "Intermediate"
400                       ElseIf Trim$(tb!RSI & "") = "" Then
410                           rtb.SelText = "<BLANK>"
420                       Else
430                           rtb.SelText = "<Unknown>"
440                       End If
450                       rtb.SelColor = vbBlack
460                       rtb.SelText = " to "
470                       rtb.SelColor = vbBlue
480                       If Trim$(LatestValue) = "R" Then
490                           rtb.SelText = "Resistant"
500                       ElseIf Trim$(LatestValue) = "S" Then
510                           rtb.SelText = "Sensitive"
520                       ElseIf Trim$(LatestValue) = "I" Then
530                           rtb.SelText = "Intermediate"
540                       ElseIf Trim$(LatestValue) = "" Then
550                           rtb.SelText = "<BLANK>"
560                       Else
570                           rtb.SelText = "<Unknown>"
580                       End If
590                       rtb.SelColor = vbBlack
600                       rtb.SelText = " by "
610                       rtb.SelColor = vbGreen
620                       rtb.SelText = tb!ArchivedBy
630                       rtb.SelColor = vbBlack
640                       rtb.SelText = vbCrLf
650                   End If
660                   LatestValue = Trim$(tb!RSI & "")
670                   rtb.SelText = vbCrLf
680                   tb.MoveNext
690               Loop
700           Next
710       Next

720       Exit Sub

DoSensitivities_Error:

          Dim strES As String
          Dim intEL As Integer

730       intEL = Erl
740       strES = Err.Description
750       LogError "frmArchiveMicro", "DoSensitivities", intEL, strES, sql

End Sub
Private Sub DoSemen()

          Dim tb As Recordset
          Dim sql As String
          Dim SID As Double
          Dim LatestValue(1 To 8) As String
          Dim dbName(1 To 8) As String
          Dim ShowName(1 To 8) As String
          Dim n As Integer
          Dim X As Integer

10        On Error GoTo DoSemen_Error

20        For n = 1 To 8
30            dbName(n) = Choose(n, "Volume", "SemenCount", "MotilityPro", "MotilityNonPro", "MotilityNonMotile", _
                                 "Consistency", "Motility", "Valid")
40            ShowName(n) = Choose(n, "Volume", "Semen Count", "Motility(Progressive)", "Motility(Non Progressive)", "Motility(Non Motile)", _
                                   "Consistency", "Motility", "Valid")
50        Next

60        SID = Val(txtSampleID) + mOffset

70        sql = "SELECT " & _
                "CASE UserName WHEN '' THEN '<BLANK>' ELSE ISNULL(UserName, '<BLANK>') END UserName, " & _
                "CASE LTRIM(RTRIM(Volume)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Volume, '<BLANK>') END Volume, " & _
                "CASE LTRIM(RTRIM(SemenCount)) WHEN '' THEN '<BLANK>' ELSE ISNULL(SemenCount, '<BLANK>') END SemenCount, " & _
                "CASE LTRIM(RTRIM(MotilityPro)) WHEN '' THEN '<BLANK>' ELSE ISNULL(MotilityPro, '<BLANK>') END MotilityPro, " & _
                "CASE LTRIM(RTRIM(MotilityNonPro)) WHEN '' THEN '<BLANK>' ELSE ISNULL(MotilityNonPro, '<BLANK>') END MotilityNonPro, " & _
                "CASE LTRIM(RTRIM(CAST(MotilityNonMotile AS nvarchar(50)))) WHEN '' THEN '<BLANK>' ELSE ISNULL(MotilityNonMotile, '<BLANK>') END MotilityNonMotile, " & _
                "CASE LTRIM(RTRIM(CAST(Consistency AS nvarchar(50)))) WHEN '' THEN '<BLANK>' ELSE ISNULL(Consistency, '<BLANK>') END Consistency, " & _
                "CASE LTRIM(RTRIM(CAST(Motility AS nvarchar(50)))) WHEN '' THEN '<BLANK>' ELSE ISNULL(Motility, '<BLANK>') END Motility, " & _
                "CASE Valid WHEN 1 THEN 'Valid' ELSE 'Not Valid' END Valid " & _
                "FROM SemenResults WHERE " & _
                "SampleID = '" & SID & "'"
80        Set tb = New Recordset
90        RecOpenServer 0, tb, sql
100       If Not tb.EOF Then
110           For n = 1 To 8
120               LatestValue(n) = tb(dbName(n)) & ""
130           Next

140           sql = "SELECT ArchiveDateTime, ArchivedBy, " & _
                    "CASE LTRIM(RTRIM(Volume)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Volume, '<BLANK>') END Volume, " & _
                    "CASE LTRIM(RTRIM(SemenCount)) WHEN '' THEN '<BLANK>' ELSE ISNULL(SemenCount, '<BLANK>') END SemenCount, " & _
                    "CASE LTRIM(RTRIM(MotilityPro)) WHEN '' THEN '<BLANK>' ELSE ISNULL(MotilityPro, '<BLANK>') END MotilityPro, " & _
                    "CASE LTRIM(RTRIM(MotilityNonPro)) WHEN '' THEN '<BLANK>' ELSE ISNULL(MotilityNonPro, '<BLANK>') END MotilityNonPro, " & _
                    "CASE LTRIM(RTRIM(CAST(MotilityNonMotile AS nvarchar(50)))) WHEN '' THEN '<BLANK>' ELSE ISNULL(MotilityNonMotile, '<BLANK>') END MotilityNonMotile, " & _
                    "CASE LTRIM(RTRIM(CAST(Consistency AS nvarchar(50)))) WHEN '' THEN '<BLANK>' ELSE ISNULL(Consistency, '<BLANK>') END Consistency, " & _
                    "CASE LTRIM(RTRIM(CAST(Motility AS nvarchar(50)))) WHEN '' THEN '<BLANK>' ELSE ISNULL(Motility, '<BLANK>') END Motility, " & _
                    "CASE Valid WHEN 1 THEN 'Valid' ELSE 'Not Valid' END Valid " & _
                    "FROM SemenResultsArc WHERE " & _
                    "SampleID = '" & SID & "' " & _
                    "ORDER BY ArchiveDateTime DESC"
150           Set tb = New Recordset
160           RecOpenServer 0, tb, sql
170           Do While Not tb.EOF
180               rtb.SelColor = vbBlack
190               If Not IsNull(tb!ArchiveDateTime) Then
200                   rtb.SelText = Format$(tb!ArchiveDateTime, "dd/MM/yy HH:mm:ss")
210               End If
220               rtb.SelText = " : " & vbCrLf
230               For X = 1 To 8
240                   If Trim$(LatestValue(X)) <> Trim$(tb(dbName(X)) & "") Then
250                       rtb.SelText = ShowName(X)
260                       rtb.SelText = " changed from "
270                       rtb.SelColor = vbRed
280                       rtb.SelText = tb.Fields(dbName(X))
290                       rtb.SelColor = vbBlack
300                       rtb.SelText = " to "
310                       rtb.SelColor = vbBlue
320                       rtb.SelText = LatestValue(X)
330                       rtb.SelColor = vbBlack
340                       rtb.SelText = " by "
350                       rtb.SelColor = vbGreen
360                       rtb.SelText = tb!ArchivedBy
370                       rtb.SelColor = vbBlack
380                       rtb.SelText = vbCrLf
390                   End If
400               Next
410               For X = 1 To 8
420                   LatestValue(X) = tb(dbName(X))
430               Next
440               rtb.SelText = vbCrLf
450               tb.MoveNext
460           Loop
470       End If

480       DoGenericSemen

490       Exit Sub

DoSemen_Error:

          Dim strES As String
          Dim intEL As Integer

500       intEL = Erl
510       strES = Err.Description
520       LogError "frmArchiveMicro", "DoSemen", intEL, strES, sql


End Sub


Private Sub DoSiteDetails()

          Dim tb As Recordset
          Dim sql As String
          Dim SID As Double
          Dim LatestValue(1 To 6) As String
          Dim dbName(1 To 6) As String
          Dim ShowName(1 To 6) As String
          Dim n As Integer
          Dim X As Integer

10        On Error GoTo DoSiteDetails_Error

20        For n = 1 To 6
30            dbName(n) = Choose(n, "Site", "SiteDetails", "PCA0", "PCA1", "PCA2", "PCA3")
40            ShowName(n) = Choose(n, "Site", "Site Details", "Current Antibiotics(1)", "Current Antibiotics(2)", "Current Antibiotics(3)", "Current Antibiotics(4)")
50        Next

60        SID = Val(txtSampleID) + mOffset

70        sql = "SELECT " & _
                "CASE UserName WHEN '' THEN '<BLANK>' ELSE ISNULL(UserName, '<BLANK>') END UserName, " & _
                "CASE LTRIM(RTRIM(CAST(Site AS nvarchar(50)))) WHEN '' THEN '<BLANK>' ELSE ISNULL(Site, '<BLANK>') END Site, " & _
                "CASE LTRIM(RTRIM(CAST(SiteDetails AS nvarchar(50)))) WHEN '' THEN '<BLANK>' ELSE ISNULL(SiteDetails, '<BLANK>') END SiteDetails, " & _
                "CASE LTRIM(RTRIM(CAST(PCA0 AS nvarchar(50)))) WHEN '' THEN '<BLANK>' ELSE ISNULL(PCA0, '<BLANK>') END PCA0, " & _
                "CASE LTRIM(RTRIM(CAST(PCA1 AS nvarchar(50)))) WHEN '' THEN '<BLANK>' ELSE ISNULL(PCA1, '<BLANK>') END PCA1, " & _
                "CASE LTRIM(RTRIM(CAST(PCA2 AS nvarchar(50)))) WHEN '' THEN '<BLANK>' ELSE ISNULL(PCA2, '<BLANK>') END PCA2, " & _
                "CASE LTRIM(RTRIM(CAST(PCA3 AS nvarchar(50)))) WHEN '' THEN '<BLANK>' ELSE ISNULL(PCA3, '<BLANK>') END PCA3 " & _
                "FROM MicroSiteDetails WHERE " & _
                "SampleID = '" & SID & "' "
80        Set tb = New Recordset
90        RecOpenServer 0, tb, sql
100       If Not tb.EOF Then
110           For n = 1 To 6
120               LatestValue(n) = tb(dbName(n)) & ""
130           Next

140           sql = "SELECT ArchiveDateTime, " & _
                    "ArchivedBy, " & _
                    "CASE UserName WHEN '' THEN '<BLANK>' ELSE ISNULL(UserName, '<BLANK>') END UserName, " & _
                    "CASE LTRIM(RTRIM(CAST(Site AS nvarchar(50)))) WHEN '' THEN '<BLANK>' ELSE ISNULL(Site, '<BLANK>') END Site, " & _
                    "CASE LTRIM(RTRIM(CAST(SiteDetails AS nvarchar(50)))) WHEN '' THEN '<BLANK>' ELSE ISNULL(SiteDetails, '<BLANK>') END SiteDetails, " & _
                    "CASE LTRIM(RTRIM(CAST(PCA0 AS nvarchar(50)))) WHEN '' THEN '<BLANK>' ELSE ISNULL(PCA0, '<BLANK>') END PCA0, " & _
                    "CASE LTRIM(RTRIM(CAST(PCA1 AS nvarchar(50)))) WHEN '' THEN '<BLANK>' ELSE ISNULL(PCA1, '<BLANK>') END PCA1, " & _
                    "CASE LTRIM(RTRIM(CAST(PCA2 AS nvarchar(50)))) WHEN '' THEN '<BLANK>' ELSE ISNULL(PCA2, '<BLANK>') END PCA2, " & _
                    "CASE LTRIM(RTRIM(CAST(PCA3 AS nvarchar(50)))) WHEN '' THEN '<BLANK>' ELSE ISNULL(PCA3, '<BLANK>') END PCA3 " & _
                    "FROM MicroSiteDetailsArc WHERE " & _
                    "SampleID = '" & SID & "' " & _
                    "ORDER BY ArchiveDateTime DESC"
150           Set tb = New Recordset
160           RecOpenServer 0, tb, sql
170           Do While Not tb.EOF
180               rtb.SelColor = vbBlack
190               If Not IsNull(tb!ArchiveDateTime) Then
200                   rtb.SelText = Format$(tb!ArchiveDateTime, "dd/MM/yy HH:mm:ss")
210               End If
220               rtb.SelText = " : " & vbCrLf
230               For X = 1 To 6
240                   If LatestValue(X) <> tb(dbName(X)) & "" Then
250                       rtb.SelText = ShowName(X)
260                       rtb.SelText = " changed from "
270                       rtb.SelColor = vbRed
280                       rtb.SelText = tb.Fields(dbName(X))
290                       rtb.SelColor = vbBlack
300                       rtb.SelText = " to "
310                       rtb.SelColor = vbBlue
320                       rtb.SelText = LatestValue(X)
330                       rtb.SelColor = vbBlack
340                       rtb.SelText = " by "
350                       rtb.SelColor = vbGreen
360                       rtb.SelText = tb!ArchivedBy & ""
370                       rtb.SelColor = vbBlack
380                       rtb.SelText = vbCrLf
390                   End If
400               Next
410               For X = 1 To 6
420                   LatestValue(X) = tb(dbName(X))
430               Next
440               rtb.SelText = vbCrLf
450               tb.MoveNext
460           Loop
470       End If

480       Exit Sub

DoSiteDetails_Error:

          Dim strES As String
          Dim intEL As Integer

490       intEL = Erl
500       strES = Err.Description
510       LogError "frmArchiveMicro", "DoSiteDetails", intEL, strES, sql


End Sub

Private Sub DoUrineRequests()

          Dim tb As Recordset
          Dim sql As String
          Dim SID As Double
          Dim LatestValue(1 To 3) As String
          Dim dbName(1 To 3) As String
          Dim ShowName(1 To 3) As String
          Dim n As Integer
          Dim X As Integer

10        On Error GoTo DoUrineRequests_Error

20        For n = 1 To 3
30            dbName(n) = Choose(n, "CS", "Pregnancy", "RedSub")
40            ShowName(n) = Choose(n, "C & S", "Pregnancy", "Red Sub")
50        Next

60        SID = Val(txtSampleID) + mOffset

70        sql = "SELECT " & _
                "CASE UserName WHEN '' THEN '<BLANK>' ELSE ISNULL(UserName, '<BLANK>') END UserName, " & _
                "CASE CS WHEN 0 THEN 'Not Selected' WHEN 1 THEN 'Selected' ELSE ISNULL(CAST(CS AS nvarchar(50)), '<BLANK>') END CS, " & _
                "CASE Pregnancy WHEN 0 THEN 'Not Selected' WHEN 1 THEN 'Selected' ELSE ISNULL(CAST(Pregnancy AS nvarchar(50)), '<BLANK>') END Pregnancy, " & _
                "CASE RedSub WHEN 0 THEN 'Not Selected' WHEN 1 THEN 'Selected' ELSE ISNULL(CAST(RedSub AS nvarchar(50)), '<BLANK>') END RedSub " & _
                "FROM UrineRequests WHERE " & _
                "SampleID = '" & SID & "' "
80        Set tb = New Recordset
90        RecOpenServer 0, tb, sql
100       If Not tb.EOF Then
110           For n = 1 To 3
120               LatestValue(n) = tb(dbName(n)) & ""
130           Next

140           sql = "SELECT ArchiveDateTime, " & _
                    "CASE ArchivedBy WHEN '' THEN '<BLANK>' ELSE ISNULL(ArchivedBy, '<BLANK>') END ArchivedBy, " & _
                    "CASE CS WHEN 0 THEN 'Not Selected' WHEN 1 THEN 'Selected' ELSE ISNULL(CAST(CS AS nvarchar(50)), '<BLANK>') END CS, " & _
                    "CASE Pregnancy WHEN 0 THEN 'Not Selected' WHEN 1 THEN 'Selected' ELSE ISNULL(CAST(Pregnancy AS nvarchar(50)), '<BLANK>') END Pregnancy, " & _
                    "CASE RedSub WHEN 0 THEN 'Not Selected' WHEN 1 THEN 'Selected' ELSE ISNULL(CAST(RedSub AS nvarchar(50)), '<BLANK>') END RedSub " & _
                    "FROM UrineRequestsArc WHERE " & _
                    "SampleID = '" & SID & "' " & _
                    "ORDER BY ArchiveDateTime DESC"
150           Set tb = New Recordset
160           RecOpenServer 0, tb, sql
170           Do While Not tb.EOF
180               rtb.SelColor = vbBlack
190               If Not IsNull(tb!ArchiveDateTime) Then
200                   rtb.SelText = Format$(tb!ArchiveDateTime, "dd/MM/yy HH:mm:ss")
210               End If
220               rtb.SelText = " : " & vbCrLf
230               For X = 1 To 3
240                   If LatestValue(X) <> tb(dbName(X)) & "" Then
250                       rtb.SelText = ShowName(X)
260                       rtb.SelText = " changed from "
270                       rtb.SelColor = vbRed
280                       rtb.SelText = tb.Fields(dbName(X))
290                       rtb.SelColor = vbBlack
300                       rtb.SelText = " to "
310                       rtb.SelColor = vbBlue
320                       rtb.SelText = LatestValue(X)
330                       rtb.SelColor = vbBlack
340                       rtb.SelText = " by "
350                       rtb.SelColor = vbGreen
360                       rtb.SelText = tb!ArchivedBy
370                       rtb.SelColor = vbBlack
380                       rtb.SelText = vbCrLf
390                   End If
400               Next
410               For X = 1 To 3
420                   LatestValue(X) = tb(dbName(X))
430               Next
440               rtb.SelText = vbCrLf
450               tb.MoveNext
460           Loop
470       End If

480       Exit Sub

DoUrineRequests_Error:

          Dim strES As String
          Dim intEL As Integer

490       intEL = Erl
500       strES = Err.Description
510       LogError "frmArchiveMicro", "DoUrineRequests", intEL, strES, sql

End Sub

Private Sub DoFaecalRequests()

          Dim tb As Recordset
          Dim sql As String
          Dim SID As Double
          Dim LatestValue(1 To 18) As String
          Dim dbName(1 To 18) As String
          Dim ShowName(1 To 18) As String
          Dim n As Integer
          Dim X As Integer

10        On Error GoTo DoFaecalRequests_Error

20        For n = 1 To 18
30            dbName(n) = Choose(n, "OP", "Rota", "Adeno", "EPC", "Culture", "CDiff", "ToxinA", "Coli0157", "OB0", "OB1", "OB2", _
                                 "ssScreen", "cS", "Campylobacter", "Cryptosporidium", "ToxinAB", "HPylori", "RedSub")
40            ShowName(n) = Choose(n, "Ova/Parasites", "Rota Virus", "Adeno Virus", "EPC", "Culture", "C.Diff", "ToxinA", "Coli0157", "O/B(0)", "O/B(1)", "O/B(2)", _
                                   "s/s Screen", "Culture/Sensitivity", "Campylobacter", "Cryptosporidium", "Toxin AB", "H.Pylori", "Red Sub")
50        Next

60        SID = Val(txtSampleID) + mOffset

70        sql = "SELECT " & _
                "CASE UserName WHEN '' THEN '<BLANK>' ELSE ISNULL(UserName, '<BLANK>') END UserName, " & _
                "CASE OP WHEN 0 THEN 'Not Selected' WHEN 1 THEN 'Selected' ELSE ISNULL(CAST(OP AS nvarchar(50)), '<BLANK>') END OP, " & _
                "CASE Rota WHEN 0 THEN 'Not Selected' WHEN 1 THEN 'Selected' ELSE ISNULL(CAST(Rota AS nvarchar(50)), '<BLANK>') END Rota, " & _
                "CASE Adeno WHEN 0 THEN 'Not Selected' WHEN 1 THEN 'Selected' ELSE ISNULL(CAST(Adeno AS nvarchar(50)), '<BLANK>') END Adeno, " & _
                "CASE EPC WHEN 0 THEN 'Not Selected' WHEN 1 THEN 'Selected' ELSE ISNULL(CAST(EPC AS nvarchar(50)), '<BLANK>') END EPC, " & _
                "CASE Culture WHEN 0 THEN 'Not Selected' WHEN 1 THEN 'Selected' ELSE ISNULL(CAST(Culture AS nvarchar(50)), '<BLANK>') END Culture, " & _
                "CASE CDiff WHEN 0 THEN 'Not Selected' WHEN 1 THEN 'Selected' ELSE ISNULL(CAST(CDiff AS nvarchar(50)), '<BLANK>') END CDiff, " & _
                "CASE ToxinA WHEN 0 THEN 'Not Selected' WHEN 1 THEN 'Selected' ELSE ISNULL(CAST(ToxinA AS nvarchar(50)), '<BLANK>') END ToxinA, " & _
                "CASE Coli0157 WHEN 0 THEN 'Not Selected' WHEN 1 THEN 'Selected' ELSE ISNULL(CAST(Coli0157 AS nvarchar(50)), '<BLANK>') END Coli0157, " & _
                "CASE OB0 WHEN 0 THEN 'Not Selected' WHEN 1 THEN 'Selected' ELSE ISNULL(CAST(OB0 AS nvarchar(50)), '<BLANK>') END OB0, " & _
                "CASE OB1 WHEN 0 THEN 'Not Selected' WHEN 1 THEN 'Selected' ELSE ISNULL(CAST(OB1 AS nvarchar(50)), '<BLANK>') END OB1, " & _
                "CASE OB2 WHEN 0 THEN 'Not Selected' WHEN 1 THEN 'Selected' ELSE ISNULL(CAST(OB2 AS nvarchar(50)), '<BLANK>') END OB2, " & _
                "CASE ssScreen WHEN 0 THEN 'Not Selected' WHEN -1 THEN 'Selected' ELSE ISNULL(ssScreen, '<BLANK>') END ssScreen, " & _
                "CASE cS WHEN 0 THEN 'Not Selected' WHEN -1 THEN 'Selected' ELSE ISNULL(cS, '<BLANK>') END cS , " & _
                "CASE Campylobacter WHEN 0 THEN 'Not Selected' WHEN -1 THEN 'Selected' ELSE ISNULL(Campylobacter, '<BLANK>') END Campylobacter, " & _
                "CASE ssScreen WHEN 0 THEN 'Not Selected' WHEN -1 THEN 'Selected' ELSE ISNULL(ssScreen, '<BLANK>') END ssScreen, " & _
                "CASE Cryptosporidium  WHEN 0 THEN 'Not Selected' WHEN 1 THEN 'Selected' ELSE ISNULL(Cryptosporidium, '<BLANK>') END Cryptosporidium , " & _
                "CASE ToxinAB WHEN 0 THEN 'Not Selected' WHEN 1 THEN 'Selected' ELSE ISNULL(ToxinAB, '<BLANK>') END ToxinAB, " & _
                "CASE HPylori WHEN 0 THEN 'Not Selected' WHEN 1 THEN 'Selected' ELSE ISNULL(HPylori, '<BLANK>') END HPylori, " & _
                "CASE RedSub WHEN 0 THEN 'Not Selected' WHEN 1 THEN 'Selected' ELSE ISNULL(CAST(RedSub AS nvarchar(50)), '<BLANK>') END RedSub " & _
                "FROM FaecalRequests WHERE " & _
                "SampleID = '" & SID & "' "
80        Set tb = New Recordset
90        RecOpenServer 0, tb, sql
100       If Not tb.EOF Then
110           For n = 1 To 18
120               LatestValue(n) = tb(dbName(n)) & ""
130           Next

140           sql = "SELECT ArchiveDateTime, " & _
                    "CASE ArchivedBy WHEN '' THEN '<BLANK>' ELSE ISNULL(ArchivedBy, '<BLANK>') END ArchivedBy, " & _
                    "CASE OP WHEN 0 THEN 'Not Selected' WHEN 1 THEN 'Selected' ELSE ISNULL(CAST(OP AS nvarchar(50)), '<BLANK>') END OP, " & _
                    "CASE Rota WHEN 0 THEN 'Not Selected' WHEN 1 THEN 'Selected' ELSE ISNULL(CAST(Rota AS nvarchar(50)), '<BLANK>') END Rota, " & _
                    "CASE Adeno WHEN 0 THEN 'Not Selected' WHEN 1 THEN 'Selected' ELSE ISNULL(CAST(Adeno AS nvarchar(50)), '<BLANK>') END Adeno, " & _
                    "CASE EPC WHEN 0 THEN 'Not Selected' WHEN 1 THEN 'Selected' ELSE ISNULL(CAST(EPC AS nvarchar(50)), '<BLANK>') END EPC, " & _
                    "CASE Culture WHEN 0 THEN 'Not Selected' WHEN 1 THEN 'Selected' ELSE ISNULL(CAST(Culture AS nvarchar(50)), '<BLANK>') END Culture, " & _
                    "CASE CDiff WHEN 0 THEN 'Not Selected' WHEN 1 THEN 'Selected' ELSE ISNULL(CAST(CDiff AS nvarchar(50)), '<BLANK>') END CDiff, " & _
                    "CASE ToxinA WHEN 0 THEN 'Not Selected' WHEN 1 THEN 'Selected' ELSE ISNULL(CAST(ToxinA AS nvarchar(50)), '<BLANK>') END ToxinA, " & _
                    "CASE Coli0157 WHEN 0 THEN 'Not Selected' WHEN 1 THEN 'Selected' ELSE ISNULL(CAST(Coli0157 AS nvarchar(50)), '<BLANK>') END Coli0157, " & _
                    "CASE OB0 WHEN 0 THEN 'Not Selected' WHEN 1 THEN 'Selected' ELSE ISNULL(CAST(OB0 AS nvarchar(50)), '<BLANK>') END OB0, " & _
                    "CASE OB1 WHEN 0 THEN 'Not Selected' WHEN 1 THEN 'Selected' ELSE ISNULL(CAST(OB1 AS nvarchar(50)), '<BLANK>') END OB1, " & _
                    "CASE OB2 WHEN 0 THEN 'Not Selected' WHEN 1 THEN 'Selected' ELSE ISNULL(CAST(OB2 AS nvarchar(50)), '<BLANK>') END OB2, " & _
                    "CASE ssScreen WHEN 0 THEN 'Not Selected' WHEN -1 THEN 'Selected' ELSE ISNULL(ssScreen, '<BLANK>') END ssScreen, " & _
                    "CASE cS  WHEN 0 THEN 'Not Selected' WHEN -1 THEN 'Selected' ELSE ISNULL(cS, '<BLANK>') END cS , " & _
                    "CASE Campylobacter WHEN 0 THEN 'Not Selected' WHEN -1 THEN 'Selected' ELSE ISNULL(Campylobacter, '<BLANK>') END Campylobacter, " & _
                    "CASE ssScreen WHEN 0 THEN 'Not Selected' WHEN -1 THEN 'Selected' ELSE ISNULL(ssScreen, '<BLANK>') END ssScreen, " & _
                    "CASE Cryptosporidium  WHEN 0 THEN 'Not Selected' WHEN 1 THEN 'Selected' ELSE ISNULL(Cryptosporidium, '<BLANK>') END Cryptosporidium , " & _
                    "CASE ToxinAB WHEN 0 THEN 'Not Selected' WHEN 1 THEN 'Selected' ELSE ISNULL(ToxinAB, '<BLANK>') END ToxinAB, " & _
                    "CASE HPylori WHEN 0 THEN 'Not Selected' WHEN 1 THEN 'Selected' ELSE ISNULL(HPylori, '<BLANK>') END HPylori, " & _
                    "CASE RedSub WHEN 0 THEN 'Not Selected' WHEN 1 THEN 'Selected' ELSE ISNULL(CAST(RedSub AS nvarchar(50)), '<BLANK>') END RedSub " & _
                    "FROM FaecalRequestsArc WHERE " & _
                    "SampleID = '" & SID & "' " & _
                    "ORDER BY ArchiveDateTime DESC"
150           Set tb = New Recordset
160           RecOpenServer 0, tb, sql
170           Do While Not tb.EOF
180               rtb.SelColor = vbBlack
190               If Not IsNull(tb!ArchiveDateTime) Then
200                   rtb.SelText = Format$(tb!ArchiveDateTime, "dd/MM/yy HH:mm:ss")
210               End If
220               rtb.SelText = " : " & vbCrLf
230               For X = 1 To 18
240                   If LatestValue(X) <> tb(dbName(X)) & "" Then
250                       rtb.SelText = ShowName(X)
260                       rtb.SelText = " changed from "
270                       rtb.SelColor = vbRed
280                       rtb.SelText = tb.Fields(dbName(X))
290                       rtb.SelColor = vbBlack
300                       rtb.SelText = " to "
310                       rtb.SelColor = vbBlue
320                       rtb.SelText = LatestValue(X)
330                       rtb.SelColor = vbBlack
340                       rtb.SelText = " by "
350                       rtb.SelColor = vbGreen
360                       rtb.SelText = tb!ArchivedBy
370                       rtb.SelColor = vbBlack
380                       rtb.SelText = vbCrLf
390                   End If
400               Next
410               For X = 1 To 18
420                   LatestValue(X) = tb(dbName(X))
430               Next
440               rtb.SelText = vbCrLf
450               tb.MoveNext
460           Loop
470       End If

480       Exit Sub

DoFaecalRequests_Error:

          Dim strES As String
          Dim intEL As Integer

490       intEL = Erl
500       strES = Err.Description
510       LogError "frmArchiveMicro", "DoFaecalRequests", intEL, strES, sql

End Sub

Private Sub DoUrine()

          Dim tb As Recordset
          Dim sql As String
          Dim SID As Double
          Dim LatestValue(1 To 21) As String
          Dim dbName(1 To 21) As String
          Dim ShowName(1 To 21) As String
          Dim n As Integer
          Dim X As Integer

10        On Error GoTo DoUrine_Error

20        For n = 1 To 21
30            dbName(n) = Choose(n, "Bacteria", "Valid", "Misc0", "Misc1", "Misc2", "Casts", _
                                 "Crystals", "RCC", "WCC", "BloodHb", "Bilirubin", "Urobilinogen", _
                                 "Ketones", "Glucose", "Protein", "pH", "FatGlobules", _
                                 "SG", "BenceJones", "HCGLevel", "Pregnancy")
40            ShowName(n) = Choose(n, "Bacteria", "Valid", "Misc(1)", "Misc(2)", "Misc(3)", "Casts", _
                                   "Crystals", "RCC", "WCC", "Blood Hb", "Bilirubin", "Urobilinogen", _
                                   "Ketones", "Glucose", "Protein", "pH", "Fat Globules", _
                                   "SG", "BenceJones", "HCGLevel", "Pregnancy")
50        Next

60        SID = Val(txtSampleID) + mOffset

70        sql = "SELECT " & _
                "CASE UserName WHEN '' THEN '<BLANK>' ELSE ISNULL(UserName, '<BLANK>') END UserName, " & _
                "CASE LTRIM(RTRIM(Bacteria)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Bacteria, '<BLANK>') END Bacteria, " & _
                "CASE LTRIM(RTRIM(Misc0)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Misc0, '<BLANK>') END Misc0, " & _
                "CASE LTRIM(RTRIM(Misc1)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Misc1, '<BLANK>') END Misc1, " & _
                "CASE LTRIM(RTRIM(Misc2)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Misc2, '<BLANK>') END Misc2, " & _
                "CASE LTRIM(RTRIM(CAST(Casts AS nvarchar(50)))) WHEN '' THEN '<BLANK>' ELSE ISNULL(Casts, '<BLANK>') END Casts, " & _
                "CASE LTRIM(RTRIM(CAST(Crystals AS nvarchar(50)))) WHEN '' THEN '<BLANK>' ELSE ISNULL(Crystals, '<BLANK>') END Crystals, " & _
                "CASE LTRIM(RTRIM(CAST(RCC AS nvarchar(50)))) WHEN '' THEN '<BLANK>' ELSE ISNULL(RCC, '<BLANK>') END RCC, " & _
                "CASE LTRIM(RTRIM(CAST(WCC AS nvarchar(50)))) WHEN '' THEN '<BLANK>' ELSE ISNULL(WCC, '<BLANK>') END WCC, " & _
                "CASE LTRIM(RTRIM(BloodHb)) WHEN '' THEN '<BLANK>' ELSE ISNULL(BloodHb, '<BLANK>') END BloodHb, " & _
                "CASE LTRIM(RTRIM(Bilirubin)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Bilirubin, '<BLANK>') END Bilirubin, " & _
                "CASE LTRIM(RTRIM(Urobilinogen)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Urobilinogen, '<BLANK>') END Urobilinogen, " & _
                "CASE LTRIM(RTRIM(Ketones)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Ketones, '<BLANK>') END Ketones, " & _
                "CASE LTRIM(RTRIM(Glucose)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Glucose, '<BLANK>') END Glucose, " & _
                "CASE LTRIM(RTRIM(Protein)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Protein, '<BLANK>') END Protein, " & _
                "CASE LTRIM(RTRIM(pH)) WHEN '' THEN '<BLANK>' ELSE ISNULL(pH, '<BLANK>') END pH, " & _
                "CASE LTRIM(RTRIM(FatGlobules)) WHEN '' THEN '<BLANK>' ELSE ISNULL(FatGlobules, '<BLANK>') END FatGlobules, " & _
                "CASE LTRIM(RTRIM(SG)) WHEN '' THEN '<BLANK>' ELSE ISNULL(SG, '<BLANK>') END SG, " & _
                "CASE LTRIM(RTRIM(BenceJones)) WHEN '' THEN '<BLANK>' ELSE ISNULL(BenceJones, '<BLANK>') END BenceJones, " & _
            "CASE LTRIM(RTRIM(HCGLevel)) WHEN '' THEN '<BLANK>' ELSE ISNULL(HCGLevel, '<BLANK>') END HCGLevel, "
80        sql = sql & "CASE LTRIM(RTRIM(Pregnancy)) WHEN '' THEN '<BLANK>' " & _
                "      WHEN 'N' THEN 'Negative' WHEN 'P' THEN 'Positive' " & _
                "      WHEN 'E' THEN 'Equivocal'WHEN 'S' THEN 'Specimen Unsuitable' " & _
                "      ELSE ISNULL(Pregnancy, '<BLANK>') END Pregnancy, " & _
                "CASE Valid WHEN 1 THEN 'Valid' ELSE 'Not Valid' END Valid " & _
                "FROM Urine WHERE " & _
                "SampleID = '" & SID & "'"
90        Set tb = New Recordset
100       RecOpenServer 0, tb, sql
110       If Not tb.EOF Then
120           For n = 1 To 21
130               LatestValue(n) = tb(dbName(n)) & ""
140           Next

150           sql = "SELECT ArchiveDateTime, ArchivedBy, " & _
                    "CASE LTRIM(RTRIM(Bacteria)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Bacteria, '<BLANK>') END Bacteria, " & _
                    "CASE LTRIM(RTRIM(Misc0)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Misc0, '<BLANK>') END Misc0, " & _
                    "CASE LTRIM(RTRIM(Misc1)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Misc1, '<BLANK>') END Misc1, " & _
                    "CASE LTRIM(RTRIM(Misc2)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Misc2, '<BLANK>') END Misc2, " & _
                    "CASE LTRIM(RTRIM(CAST(Casts AS nvarchar(50)))) WHEN '' THEN '<BLANK>' ELSE ISNULL(Casts, '<BLANK>') END Casts, " & _
                    "CASE LTRIM(RTRIM(CAST(Crystals AS nvarchar(50)))) WHEN '' THEN '<BLANK>' ELSE ISNULL(Crystals, '<BLANK>') END Crystals, " & _
                    "CASE LTRIM(RTRIM(CAST(RCC AS nvarchar(50)))) WHEN '' THEN '<BLANK>' ELSE ISNULL(RCC, '<BLANK>') END RCC, " & _
                    "CASE LTRIM(RTRIM(CAST(WCC AS nvarchar(50)))) WHEN '' THEN '<BLANK>' ELSE ISNULL(WCC, '<BLANK>') END WCC, " & _
                    "CASE LTRIM(RTRIM(BloodHb)) WHEN '' THEN '<BLANK>' ELSE ISNULL(BloodHb, '<BLANK>') END BloodHb, " & _
                    "CASE LTRIM(RTRIM(Bilirubin)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Bilirubin, '<BLANK>') END Bilirubin, " & _
                    "CASE LTRIM(RTRIM(Urobilinogen)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Urobilinogen, '<BLANK>') END Urobilinogen, " & _
                    "CASE LTRIM(RTRIM(Ketones)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Ketones, '<BLANK>') END Ketones, " & _
                    "CASE LTRIM(RTRIM(Glucose)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Glucose, '<BLANK>') END Glucose, " & _
                    "CASE LTRIM(RTRIM(Protein)) WHEN '' THEN '<BLANK>' ELSE ISNULL(Protein, '<BLANK>') END Protein, " & _
                    "CASE LTRIM(RTRIM(pH)) WHEN '' THEN '<BLANK>' ELSE ISNULL(pH, '<BLANK>') END pH, " & _
                    "CASE LTRIM(RTRIM(FatGlobules)) WHEN '' THEN '<BLANK>' ELSE ISNULL(FatGlobules, '<BLANK>') END FatGlobules, " & _
                    "CASE LTRIM(RTRIM(SG)) WHEN '' THEN '<BLANK>' ELSE ISNULL(SG, '<BLANK>') END SG, " & _
                    "CASE LTRIM(RTRIM(BenceJones)) WHEN '' THEN '<BLANK>' ELSE ISNULL(BenceJones, '<BLANK>') END BenceJones, " & _
              "CASE LTRIM(RTRIM(HCGLevel)) WHEN '' THEN '<BLANK>' ELSE ISNULL(HCGLevel, '<BLANK>') END HCGLevel, "
160           sql = sql & "CASE LTRIM(RTRIM(Pregnancy)) WHEN '' THEN '<BLANK>' " & _
                    "      WHEN 'N' THEN 'Negative' WHEN 'P' THEN 'Positive' " & _
                    "      WHEN 'E' THEN 'Equivocal'WHEN 'S' THEN 'Specimen Unsuitable' " & _
                    "      ELSE ISNULL(Pregnancy, '<BLANK>') END Pregnancy, " & _
                    "CASE Valid WHEN 1 THEN 'Valid' ELSE 'Not Valid' END Valid " & _
                    "FROM UrineArc WHERE " & _
                    "SampleID = '" & SID & "' " & _
                    "ORDER BY ArchiveDateTime DESC"
170           Set tb = New Recordset
180           RecOpenServer 0, tb, sql
190           Do While Not tb.EOF
200               rtb.SelColor = vbBlack
210               If Not IsNull(tb!ArchiveDateTime) Then
220                   rtb.SelText = Format$(tb!ArchiveDateTime, "dd/MM/yy HH:mm:ss")
230               End If
240               rtb.SelText = " : " & vbCrLf
250               For X = 1 To 21
260                   If LatestValue(X) <> tb(dbName(X)) & "" Then
270                       rtb.SelText = ShowName(X)
280                       rtb.SelText = " changed from "
290                       rtb.SelColor = vbRed
300                       rtb.SelText = tb.Fields(dbName(X))
310                       rtb.SelColor = vbBlack
320                       rtb.SelText = " to "
330                       rtb.SelColor = vbBlue
340                       rtb.SelText = LatestValue(X)
350                       rtb.SelColor = vbBlack
360                       rtb.SelText = " by "
370                       rtb.SelColor = vbGreen
380                       rtb.SelText = tb!ArchivedBy
390                       rtb.SelColor = vbBlack
400                       rtb.SelText = vbCrLf
410                   End If
420               Next
430               For X = 1 To 21
440                   LatestValue(X) = tb(dbName(X))
450               Next
460               rtb.SelText = vbCrLf
470               tb.MoveNext
480           Loop
490       End If

500       Exit Sub

DoUrine_Error:

          Dim strES As String
          Dim intEL As Integer

510       intEL = Erl
520       strES = Err.Description
530       LogError "frmArchiveMicro", "DoUrine", intEL, strES, sql

End Sub

Private Sub FillOptions()

          Dim tb As Recordset
          Dim sql As String
          Dim SID As Double
          Dim Index As Integer
          Dim Dept As String

10        On Error GoTo FillOptions_Error

20        For Index = 0 To 11
30            optDept(Index).Enabled = False
40            optDept(Index).ForeColor = vbBlack
50        Next

60        rtb.Text = ""

70        SID = Val(txtSampleID) + mOffset

80        sql = "SELECT COUNT(*) AS Tot FROM ArcDemographics WHERE " & _
                "SampleID = '" & SID & "'"
90        Set tb = New Recordset
100       RecOpenServer 0, tb, sql
110       If tb!Tot > 0 Then
120           optDept(9).Enabled = True
130           optDept(9).ForeColor = vbRed
140       End If

150       If optMS(0) Then
160           For Index = 1 To 11
170               Dept = Choose(Index, "FaecesArc", "MicroSiteDetailsArc", _
                                "UrineArc", "GenericResultsArc", "IsolatesArc", "SensitivitiesArc", _
                                "UrineRequestsArc", "FaecalRequestsArc", "UrineIdentArc", "ArcDemographics", "CommentsArc")
180               sql = "SELECT COUNT(*) AS Tot FROM " & Dept & " WHERE " & _
                        "SampleID = '" & SID & "'"
190               Set tb = New Recordset
200               RecOpenServer 0, tb, sql
210               If tb!Tot > 0 Then
220                   optDept(Index - 1).Enabled = True
230                   optDept(Index - 1).ForeColor = vbRed
240               Else
250                   Dept = Choose(Index, "Faeces", "MicroSiteDetails", _
                                    "Urine", "GenericResults", "Isolates", "Sensitivities", _
                                    "UrineRequests", "FaecalRequests", "UrineIdent", "Demographics", "Comments")
260                   sql = "SELECT COUNT(*) AS Tot FROM " & Dept & " WHERE " & _
                            "SampleID = '" & SID & "'"
270                   Set tb = New Recordset
280                   RecOpenServer 0, tb, sql
290                   If tb!Tot > 0 Then
300                       optDept(Index - 1).Enabled = True
310                       optDept(Index - 1).ForeColor = vbBlack
320                   End If
330               End If
340           Next
350       Else
360           sql = "SELECT COUNT(*) AS Tot FROM SemenResultsArc WHERE " & _
                    "SampleID = '" & Val(txtSampleID) + mOffset & "'"
370           Set tb = New Recordset
380           RecOpenServer 0, tb, sql
390           If tb!Tot > 0 Then
400               optDept(11).Enabled = True
410               optDept(11).ForeColor = vbRed
420           Else
430               sql = "SELECT COUNT(*) AS Tot FROM SemenResults WHERE " & _
                        "SampleID = '" & Val(txtSampleID) + mOffset & "'"
440               Set tb = New Recordset
450               RecOpenServer 0, tb, sql
460               If tb!Tot > 0 Then
470                   optDept(11).Enabled = True
480                   optDept(11).ForeColor = vbBlack
490               End If
500           End If
510       End If

520       Exit Sub

FillOptions_Error:

          Dim strES As String
          Dim intEL As Integer

530       intEL = Erl
540       strES = Err.Description
550       LogError "frmArchiveMicro", "FillOptions", intEL, strES, sql

End Sub

Private Sub bprint_Click()

10        On Error GoTo bprint_Click_Error

20        rtb.SelStart = 0
30        rtb.SelLength = 10000000#
40        rtb.SelPrint Printer.hDC

50        Exit Sub

bprint_Click_Error:

          Dim strES As String
          Dim intEL As Integer

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmArchiveMicro", "bPrint_Click", intEL, strES
End Sub

Private Sub cmdCancel_Click()

10        Unload Me

End Sub


Private Sub cmdStart_Click()

10        FillOptions

End Sub

Private Sub Form_Activate()

10        mOffset = SysOptMicroOffset(0)

20        If pSampleIDNoOffset > 0 Then
30            txtSampleID = pSampleIDNoOffset
40            FillOptions
50        End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

10        pSampleIDNoOffset = 0

End Sub

Private Sub optDept_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

10        rtb.Text = ""

20        If optDept(Index).ForeColor = vbRed Then
30            Select Case Index
              Case 0: DoArcFaecesUserName: DoFaeces
40            Case 1: DoArcSiteDetailsUserName: DoSiteDetails
50            Case 2:
60                With frmAuditMicro
70                    .TableName = "Urine"
80                    .SampleID = txtSampleID + SysOptMicroOffset(0)
90                    .Show 1
100               End With

                  'DoArcUrineUserName: DoUrine
110           Case 3: DoArcGenericUserName: DoGeneric
120           Case 4: DoArcIsolatesUserName: DoIsolates
130           Case 5: DoArcSensitivitiesUserCode: DoSensitivities
140           Case 6: DoArcUrineRequestsUserName: DoUrineRequests
150           Case 7: DoArcFaecalRequestsUserName: DoFaecalRequests
160           Case 8: DoArcIdentUserName: DoIdent
170           Case 9: DoArcDemographicsOperator: DoDemographics
180           Case 10: DoArcCommentsUserName: DoComments
190           Case 11: DoArcSemenUserName: DoSemen
200           End Select
210       Else
220           Select Case Index
              Case 0: DoFaecesUserName
230           Case 1: DoSiteDetailsUserName
240           Case 2:
250               With frmAuditMicro
260                   .TableName = "Urine"
270                   .SampleID = txtSampleID + SysOptMicroOffset(0)
280                   .Show 1
290               End With

                  'DoUrineUserName
300           Case 3: DoGenericUserName
310           Case 4: DoIsolatesUserName
320           Case 5: DoSensitivitiesUserCode
330           Case 6: DoUrineRequestsUserName
340           Case 7: DoFaecalRequestsUserName
350           Case 8: DoIdentUserName
360           Case 9: DoDemographicsOperator
370           Case 10: DoCommentsUserName
380           Case 11: DoSemenUserName
390           End Select
400       End If

End Sub


Public Property Let SID(ByVal lNewValue As Double)

10        pSampleIDNoOffset = lNewValue

End Property

Private Sub optMS_Click(Index As Integer)

10        If Index = 0 Then
20            mOffset = SysOptMicroOffset(0)
30        Else
40            mOffset = SysOptSemenOffset(0)
50        End If

60        FillOptions

End Sub


