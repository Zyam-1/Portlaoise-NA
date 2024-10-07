VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmArchiveNew 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire"
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   13590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Exit"
      Height          =   1100
      Left            =   11970
      Picture         =   "frmArchiveNew.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7560
      Width           =   1200
   End
   Begin VB.Frame Frame2 
      Caption         =   "Search"
      Height          =   855
      Left            =   11760
      TabIndex        =   9
      Top             =   1140
      Visible         =   0   'False
      Width           =   1545
      Begin VB.OptionButton optSearch 
         Caption         =   "Products"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   11
         Top             =   510
         Width           =   975
      End
      Begin VB.OptionButton optSearch 
         Caption         =   "Patients"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   10
         Top             =   270
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Show"
      Height          =   855
      Left            =   11760
      TabIndex        =   6
      Top             =   180
      Width           =   1575
      Begin VB.OptionButton optShow 
         Caption         =   "All"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   270
         Width           =   495
      End
      Begin VB.OptionButton optShow 
         Caption         =   "Only Changes"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   510
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   1100
      Left            =   11970
      Picture         =   "frmArchiveNew.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6180
      Width           =   1200
   End
   Begin VB.TextBox txtSampleId 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   11760
      TabIndex        =   3
      Top             =   3120
      Width           =   1545
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start"
      Height          =   1100
      Left            =   11970
      Picture         =   "frmArchiveNew.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3600
      Width           =   1200
   End
   Begin RichTextLib.RichTextBox rtb 
      Height          =   8505
      Left            =   60
      TabIndex        =   1
      Top             =   270
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   15002
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmArchiveNew.frx":149E
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sample ID"
      Height          =   285
      Left            =   11760
      TabIndex        =   4
      Top             =   2820
      Width           =   1545
   End
End
Attribute VB_Name = "frmArchiveNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pTableName As String
Private pTableNameAudit As String




Private Sub FillCurrent()

          Dim sql As String
          Dim tb As Recordset
          Dim dt As String
          Dim OP As String
          Dim FieldChanged As Boolean
          Dim X As Integer
          Dim Y As Integer
          Dim adt As Integer
          Dim AB As Integer
          Dim NameDisplayed As Boolean
          Dim yy As Integer
          Dim Previous As String
          Dim SearchBy As String


10        On Error GoTo FillCurrent_Error

20        rtb.Text = ""
30        rtb.SelFontSize = 12

40        If Trim$(txtSampleID) = "" Then Exit Sub


50        SearchBy = "SampleID"

60        sql = "SELECT 1 Tag, *, 'u' ArchivedBy, '1/1/2030' ArchiveDateTime FROM " & pTableName & " WHERE " & _
                SearchBy & " = '" & txtSampleID & "' " & _
                "UNION " & _
                "SELECT 2 Tag, * FROM " & pTableNameAudit & " WHERE " & _
                SearchBy & " = '" & txtSampleID & "' " & _
                "ORDER BY ArchiveDateTime DESC"
70        Set tb = New Recordset
80        RecOpenClient 0, tb, sql
90        If tb.EOF Then
100           rtb.SelText = "No Current Record found." & vbCrLf
110           Exit Sub
120       End If

130       ReDim Records(0 To tb.RecordCount - 1, 0 To tb.Fields.Count - 1)
140       ReDim titles(0 To tb.Fields.Count - 1)
150       Y = -1
160       Do While Not tb.EOF
170           Y = Y + 1
180           For X = 0 To tb.Fields.Count - 1
190               titles(X) = tb.Fields(X).Name
200               Records(Y, X) = tb.Fields(X).Value
210           Next
220           tb.MoveNext
230       Loop

240       dt = "Unknown            "
250       tb.MoveFirst
260       If Not IsNull(tb!DateTimeDemographics) Then
270           If IsDate(tb!DateTimeDemographics) Then
280               dt = Format$(tb!DateTimeDemographics, "dd/MM/yyyy HH:nn:ss")
290           End If
300       End If

310       OP = "Unknown"
320       If Trim$(tb!Operator & "") <> "" Then
330           OP = tb!Operator
340       End If

350       rtb.SelBold = True
360       rtb.SelColor = vbBlue
370       rtb.SelFontSize = 12
380       rtb.SelUnderline = True
390       rtb.SelText = "Current Record entered by " & OP & " at " & dt & vbCrLf & vbCrLf

400       For X = 1 To UBound(titles) - 2
410           FieldChanged = False
420           For Y = 1 To UBound(Records)
430               If Trim$(Records(Y, X) & "") & "" <> Trim$(Records(0, X) & "") Then
440                   FieldChanged = True
450                   Exit For
460               End If
470           Next
480           If optShow(0) Or (optShow(1) And FieldChanged) Then
490               rtb.SelFontName = "Courier New"
500               rtb.SelBold = True
510               rtb.SelColor = vbBlue
520               rtb.SelFontSize = 12
530               rtb.SelText = Left$(titles(X) & Space$(20), 20)
540               rtb.SelFontSize = 12
550               rtb.SelBold = True
560           End If
570           If FieldChanged Then
580               rtb.SelColor = vbRed
590               rtb.SelText = Left$(Records(0, X) & Space$(25), 25)
600               rtb.SelText = " See below for changes"
610           Else
620               If optShow(0) Then
630                   rtb.SelText = Left$(Records(0, X) & Space$(25), 25)
640               End If
650           End If
660           If optShow(0) Or (optShow(1) And FieldChanged) Then
670               rtb.SelText = vbCrLf
680           End If
690       Next

700       rtb.SelText = vbCrLf
710       rtb.SelText = vbCrLf
720       rtb.SelFontSize = 16
730       rtb.SelColor = vbBlack
740       rtb.SelText = String(40, "-") & vbCrLf
750       rtb.SelFontSize = 16
760       rtb.SelColor = vbBlack
770       rtb.SelText = "Audit Records:" & vbCrLf


780       If UBound(Records) = 0 Then
790           rtb.SelFontSize = 16
800           rtb.SelColor = vbBlack
810           rtb.SelText = "No Changes Made"
820       Else
830           For X = UBound(titles) - 1 To UBound(titles)
840               If UCase$(titles(X)) = "ARCHIVEDATETIME" Then
850                   adt = X
860               End If
870               If UCase$(titles(X)) = "ARCHIVEDBY" Then
880                   AB = X
890               End If
900           Next

910           For X = 1 To UBound(titles)
920               FieldChanged = False
930               If X <> adt And X <> AB Then
940                   For Y = 1 To UBound(Records)
950                       If Trim$(Records(Y, X) & "") <> Trim$(Records(0, X) & "") Then
960                           FieldChanged = True
970                           Exit For
980                       End If
990                   Next
1000              End If
1010              If FieldChanged Then
1020                  Previous = Trim$(Records(0, X) & "")
1030                  NameDisplayed = False
1040                  For yy = 1 To UBound(Records)
1050                      If Previous <> Trim$(Records(yy, X) & "") Then
1060                          If Not NameDisplayed Then
1070                              rtb.SelText = vbCrLf
1080                              rtb.SelBold = True
1090                              rtb.SelColor = vbBlue
1100                              rtb.SelFontSize = 12
1110                              rtb.SelUnderline = True
1120                              rtb.SelText = titles(X) & vbCrLf
1130                              NameDisplayed = True
1140                          End If

1150                          rtb.SelFontSize = 12
1160                          rtb.SelColor = vbBlack
1170                          rtb.SelText = Records(yy, adt) & " "
1180                          rtb.SelColor = vbBlue
1190                          rtb.SelText = Records(yy, AB) & ""
1200                          rtb.SelColor = vbBlack
1210                          rtb.SelText = " Changed "
1220                          rtb.SelColor = vbRed
1230                          rtb.SelBold = True
1240                          If Records(yy, X) = "" Then
1250                              rtb.SelText = "<Blank> "
1260                          Else
1270                              rtb.SelText = Records(yy, X) & ""
1280                          End If
1290                          rtb.SelColor = vbBlack
1300                          rtb.SelBold = False
1310                          rtb.SelText = " to "
1320                          rtb.SelBold = True
1330                          rtb.SelColor = vbRed
1340                          If Trim$(Previous) = "" Then
1350                              rtb.SelText = "<Blank>" & vbCrLf
1360                          Else
1370                              rtb.SelText = Previous & vbCrLf
1380                          End If
1390                          Previous = Trim$(Records(yy, X) & "")
1400                      End If
1410                  Next
1420              End If
1430          Next

1440          If NameDisplayed Then
1450              rtb.SelText = vbCrLf
1460          End If

1470      End If

1480      Exit Sub

FillCurrent_Error:

          Dim strES As String
          Dim intEL As Integer

1490      intEL = Erl
1500      strES = Err.Description
1510      LogError "frmArchive", "FillCurrent", intEL, strES, sql

End Sub

Private Sub cmdCancel_Click()

10        Unload Me

End Sub


Private Sub cmdPrint_Click()

10        On Error GoTo cmdPrint_Click_Error

20        rtb.SelStart = 0
30        rtb.SelLength = 10000000#
40        rtb.SelPrint Printer.hDC

50        Exit Sub

cmdPrint_Click_Error:

          Dim strES As String
          Dim intEL As Integer

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmArchive", "cmdPrint_Click", intEL, strES

End Sub

Private Sub cmdStart_Click()

10        FillCurrent

End Sub

Public Property Let TableName(ByVal sNewValue As String)

10        pTableName = sNewValue
20        pTableNameAudit = sNewValue & "Audit"

End Property
Public Property Let SampleID(ByVal sNewValue As String)

10        txtSampleID = sNewValue

End Property

Private Sub Form_Activate()
10        FillCurrent
End Sub


Private Sub optSearch_Click(Index As Integer)

10        If Index = 0 Then
20            pTableName = "PatientDetails"
30            lblTitle = "SampleID"
40        Else
50            pTableName = "Latest"
60            lblTitle = "Pack Number"
70        End If

80        pTableNameAudit = pTableName & "Audit"

90        rtb.Text = ""
100       txtSampleID = ""

End Sub


Private Sub optShow_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

10        FillCurrent

End Sub

