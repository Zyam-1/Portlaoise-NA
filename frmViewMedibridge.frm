VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmViewMedibridge 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - External Reports"
   ClientHeight    =   6000
   ClientLeft      =   165
   ClientTop       =   600
   ClientWidth     =   10980
   Icon            =   "frmViewMedibridge.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   10980
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Exit"
      Height          =   840
      Left            =   9300
      Picture         =   "frmViewMedibridge.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4980
      Width           =   1605
   End
   Begin RichTextLib.RichTextBox rtbResult 
      Height          =   4005
      Left            =   60
      TabIndex        =   10
      Top             =   870
      Width           =   10845
      _ExtentX        =   19129
      _ExtentY        =   7064
      _Version        =   393217
      BackColor       =   -2147483624
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmViewMedibridge.frx":0614
   End
   Begin VB.Label lblSampleID 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1230
      TabIndex        =   9
      Top             =   150
      Width           =   1395
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "SampleID"
      Height          =   195
      Left            =   510
      TabIndex        =   8
      Top             =   180
      Width           =   690
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Report Time"
      Height          =   195
      Left            =   7440
      TabIndex        =   7
      Top             =   180
      Width           =   870
   End
   Begin VB.Label lblMessageTime 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8370
      TabIndex        =   6
      Top             =   150
      Width           =   2415
   End
   Begin VB.Label lblSex 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5910
      TabIndex        =   5
      Top             =   480
      Width           =   1365
   End
   Begin VB.Label lblDoB 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3420
      TabIndex        =   4
      Top             =   480
      Width           =   1545
   End
   Begin VB.Label lblPatName 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3420
      TabIndex        =   3
      Top             =   240
      Width           =   3855
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Sex"
      Height          =   195
      Left            =   5610
      TabIndex        =   2
      Top             =   540
      Width           =   270
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "DoB"
      Height          =   195
      Left            =   3030
      TabIndex        =   1
      Top             =   510
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      Height          =   195
      Left            =   2970
      TabIndex        =   0
      Top             =   180
      Width           =   420
   End
End
Attribute VB_Name = "frmViewMedibridge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pSampleID As String

Private Sub cmdCancel_Click()

10        Unload Me

End Sub

Private Sub Form_Activate()

          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo Form_Activate_Error

20        sql = "SELECT * FROM Demographics WHERE " & _
                "SampleID = '" & pSampleID & "'"

30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If Not tb.EOF Then
60            lblSampleID = Val(tb!SampleID)
70            lblPatName = tb!PatName & ""
80            lblDoB = tb!Dob & ""
90            Select Case UCase$(Left$(Trim$(tb!sex & ""), 1))
              Case "M": lblSex = "Male"
100           Case "F": lblSex = "Female"
110           Case "U": lblSex = "Unknown"
120           Case Else: lblSex = "Not Given"
130           End Select
140       End If

150       rtbResult.SelText = ""

160       sql = "SELECT * from MedibridgeResults WHERE " & _
                "SampleID = '" & pSampleID & "'"
170       Set tb = New Recordset
180       RecOpenServer 0, tb, sql
190       Do While Not tb.EOF
200           lblMessageTime = tb!MessageTime

210           With rtbResult
220               .SelIndent = 0
230               .SelColor = vbBlue
240               .SelBold = False
250               .SelText = "Request: "
260               .SelBold = False
270               .SelText = .SelText & tb!Request & vbCrLf
280               .SelColor = vbBlack
290               .SelBold = True
300               .SelIndent = 200
310               .SelText = .SelText & tb!Result & ""
320           End With

330           tb.MoveNext
340       Loop

350       sql = "SELECT * from EResults WHERE " & _
                "SampleID = '" & pSampleID & "'"
360       Set tb = New Recordset
370       RecOpenServer 0, tb, sql
380       Do While Not tb.EOF

390           With rtbResult
400               .SelIndent = 0
410               .SelColor = vbBlue
                  '.SelBold = False
                  '.SelText = "Analyte: "
420               .SelBold = True
430               .SelText = .SelText & tb!Analyte & ": "
440               .SelColor = vbBlack
450               .SelBold = True
460               .SelIndent = 200
470               If Trim$(tb!Result & "") <> "" Then
480                   .SelText = .SelText & tb!Result & " " & tb!Units & ""
490               Else
500                   .SelText = .SelText & "Not yet Available."
510               End If
520               .SelText = .SelText & vbCrLf
530           End With

540           tb.MoveNext
550       Loop

560       Exit Sub

Form_Activate_Error:

          Dim strES As String
          Dim intEL As Integer



570       intEL = Erl
580       strES = Err.Description
590       LogError "frmViewMedibridge", "Form_Activate", intEL, strES, sql

End Sub


Public Property Let SampleID(ByRef NewValue As String)

10        On Error GoTo SampleID_Error

20        pSampleID = Val(NewValue)
          'pSampleID = Val(NewValue) + SysOptMicroOffset

30        Exit Property

SampleID_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmViewMedibridge", "SampleID", intEL, strES


End Property
