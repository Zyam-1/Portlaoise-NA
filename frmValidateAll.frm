VERSION 5.00
Begin VB.Form frmValidateAll 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetAcquire --- Validate Confirmation"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   7170
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkSections 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Blood Culture"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   300
      TabIndex        =   14
      Tag             =   "BLOODCULTURE"
      Top             =   1515
      Width           =   4935
   End
   Begin VB.CommandButton bCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   1100
      Left            =   5700
      Picture         =   "frmValidateAll.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4270
      Width           =   1200
   End
   Begin VB.CommandButton cmdValidateMicro 
      Caption         =   "&Validate"
      Height          =   1100
      Left            =   5700
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmValidateAll.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3120
      Width           =   1200
   End
   Begin VB.CheckBox chkSections 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "OP"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   300
      TabIndex        =   9
      Tag             =   "OP"
      Top             =   4995
      Width           =   4935
   End
   Begin VB.CheckBox chkSections 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "HPYLORI"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   300
      TabIndex        =   8
      Tag             =   "HPYLORI"
      Top             =   4560
      Width           =   4935
   End
   Begin VB.CheckBox chkSections 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "C. Diff"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   300
      TabIndex        =   7
      Tag             =   "CDIFF"
      Top             =   4125
      Width           =   4935
   End
   Begin VB.CheckBox chkSections 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "CSF"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   300
      TabIndex        =   6
      Tag             =   "CSF"
      Top             =   3690
      Width           =   4935
   End
   Begin VB.CheckBox chkSections 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "RSV"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   300
      TabIndex        =   5
      Tag             =   "RSV"
      Top             =   3255
      Width           =   4935
   End
   Begin VB.CheckBox chkSections 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Red Sub"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   300
      TabIndex        =   4
      Tag             =   "REDSUB"
      Top             =   2820
      Width           =   4935
   End
   Begin VB.CheckBox chkSections 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Rota / Adeno"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   300
      TabIndex        =   3
      Tag             =   "ROTAADENO"
      Top             =   2385
      Width           =   4935
   End
   Begin VB.CheckBox chkSections 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "FOB"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   300
      TabIndex        =   2
      Tag             =   "FOB"
      Top             =   1950
      Width           =   4935
   End
   Begin VB.CheckBox chkSections 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "C && S"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   300
      TabIndex        =   1
      Tag             =   "CANDS"
      Top             =   1515
      Width           =   4935
   End
   Begin VB.CheckBox chkSections 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Urine"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   300
      TabIndex        =   0
      Tag             =   "URINE"
      Top             =   1080
      Width           =   4935
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Please select the sections you want to validate."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   300
      TabIndex        =   11
      Top             =   555
      Width           =   4920
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "The following sections are not validated."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   660
      TabIndex        =   10
      Top             =   300
      Width           =   4200
   End
End
Attribute VB_Name = "frmValidateAll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public SampleIDToValidate As Double

Private Sub bcancel_Click()
10        Unload Me
End Sub

Private Sub cmdValidateMicro_Click()

          Dim i As Integer

10        On Error GoTo cmdValidateMicro_Click_Error

20        For i = 0 To chkSections.Count - 1
30            If chkSections(i).Enabled = True And chkSections(i).Value = 1 Then
40                UpdatePrintValidLog SampleIDToValidate, chkSections(i).Tag, 1, 0
50            End If
60        Next i

70        Unload Me

80        Exit Sub

cmdValidateMicro_Click_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmValidateAll", "cmdValidateMicro_Click", intEL, strES

End Sub

Private Sub Form_Activate()

          Dim i As Integer
          Dim ValidationRequired As Boolean


10        On Error GoTo Form_Activate_Error

20        ValidationRequired = False

30        For i = 0 To chkSections.Count - 1
40            If chkSections(i).Enabled = True Then
50                ValidationRequired = True
60            End If
70        Next i

80        If Not ValidationRequired Then Unload Me

90        Exit Sub

Form_Activate_Error:

          Dim strES As String
          Dim intEL As Integer

100       intEL = Erl
110       strES = Err.Description
120       LogError "frmValidateAll", "Form_Activate", intEL, strES


End Sub

Private Sub Form_Load()

          Dim tb As Recordset
          Dim sql As String
          Dim i As Integer

10        On Error GoTo Form_Load_Error


          'CHECK IF URINE IS PRESENT
20        sql = "Select * From Urine Where SampleID = '%sampleid'"
30        sql = Replace(sql, "%sampleid", SampleIDToValidate)
40        Set tb = New Recordset
50        RecOpenClient 0, tb, sql
60        If Not tb.EOF Then
70            For i = 0 To tb.Fields.Count - 1
80                If tb.Fields(i).Name <> "Valid" And tb.Fields(i).Name <> "Healthlink" And _
                     tb.Fields(i).Name <> "Printed" And tb.Fields(i).Name <> "UserName" And _
                     tb.Fields(i).Name <> "rowguid" And tb.Fields(i).Name <> "SampleID" Then
90                    If tb.Fields(i).Value & "" <> "" Then
100                       chkSections(0).Enabled = True
110                   End If
120               End If
130           Next i
140       End If

          'CHECK IF C&S IS PRESENT
150       sql = "Select Count(*) As Cnt From Isolates Where SampleID = '%sampleid'"
160       sql = Replace(sql, "%sampleid", SampleIDToValidate)
170       Set tb = New Recordset
180       RecOpenClient 0, tb, sql
190       If tb!Cnt > 0 Then
200           If InStr(1, UCase$(GetMicroSite(SampleIDToValidate)), "BLOOD CULTURE") > 0 And _
                 frmEditMicrobiologyNew.SSTab1.TabVisible(12) = True Then

210               chkSections(10).Enabled = True
220               chkSections(10).Visible = True
230               chkSections(1).Visible = False
240           Else
250               chkSections(1).Enabled = True
260               chkSections(1).Visible = True
270               chkSections(10).Visible = False
280           End If
290       End If

          'CHECK IF FAECES RESULTS IS PRESENT
300       sql = "Select * From Faeces Where SampleID = '%sampleid'"
310       sql = Replace(sql, "%sampleid", SampleIDToValidate)
320       Set tb = New Recordset
330       RecOpenClient 0, tb, sql
340       If Not tb.EOF Then
350           If (tb!OB0 & tb!OB1 & tb!OB2) & "" <> "" Then
360               chkSections(2).Enabled = True
370           End If
380           If tb!Rota & "" <> "" Or tb!Adeno & "" <> "" Then
390               chkSections(3).Enabled = True
400           End If
410           If tb!CDiffCulture & "" <> "" Or tb!ToxinAB <> "" Then
420               chkSections(7).Enabled = True
430           End If
440           If tb!HPylori & "" <> "" Then
450               chkSections(8).Enabled = True
460           End If
470           If tb!OP & "" <> "" Or tb!OP0 & "" <> "" Or tb!OP1 & "" <> "" Or tb!OP2 & "" <> "" Or tb!Cryptosporidium & "" <> "" Then
480               chkSections(9).Enabled = True
490           End If
500       End If

          'CHECK IF GENERIC RESULTS ARE PRESENT
510       sql = "Select Distinct TestName From GenericResults Where SampleID = '%sampleid'"
520       sql = Replace(sql, "%sampleid", SampleIDToValidate)
530       Set tb = New Recordset
540       RecOpenClient 0, tb, sql
550       If Not tb.EOF Then
560           While Not tb.EOF
570               If InStr(1, tb!TestName & "", "Fluid") > 0 Or _
                     InStr(1, tb!TestName & "", "CSF") > 0 Or _
                     InStr(1, tb!TestName & "", "Fungal") > 0 Or _
                     InStr(1, tb!TestName & "", "BATScreen") > 0 Or _
                     tb!TestName & "" = "PneumococcalAT" Or _
                     tb!TestName & "" = "LegionellaAT" Then
580                   chkSections(6).Enabled = True
590               ElseIf UCase(tb!TestName & "") = "REDSUB" Then
600                   chkSections(4).Enabled = True
610               ElseIf UCase(tb!TestName & "") = "RSV" Then
620                   chkSections(5).Enabled = True
630               End If

640               tb.MoveNext
650           Wend
660       End If

          'CHECK IF ENTRY IS ALREADY THERE IN PRINTVALIDLOG
670       sql = "Select * From PrintValidLog Where SampleID = '%sampleid'"
680       sql = Replace(sql, "%sampleid", SampleIDToValidate)
690       Set tb = New Recordset
700       RecOpenClient 0, tb, sql
710       If Not tb.EOF Then
720           While Not tb.EOF
730               Select Case tb!Department
                  Case "U":
740                   If tb!Valid = 1 Then
750                       chkSections(0).Value = 1
760                       chkSections(0).Enabled = False
770                   End If
780               Case "D":
790                   If tb!Valid = 1 Then
800                       chkSections(1).Value = 1
810                       chkSections(1).Enabled = False
820                   End If
830               Case "F":
840                   If tb!Valid = 1 Then
850                       chkSections(2).Value = 1
860                       chkSections(2).Enabled = False
870                   End If
880               Case "A":

890                   If tb!Valid = 1 Then
900                       chkSections(3).Value = 1
910                       chkSections(3).Enabled = False
920                   End If
930               Case "R":
940                   If tb!Valid = 1 Then
950                       chkSections(4).Value = 1
960                       chkSections(4).Enabled = False
970                   End If
980               Case "V":
990                   If tb!Valid = 1 Then
1000                      chkSections(5).Value = 1
1010                      chkSections(5).Enabled = False
1020                  End If
1030              Case "C":
1040                  If tb!Valid = 1 Then
1050                      chkSections(6).Value = 1
1060                      chkSections(6).Enabled = False
1070                  End If
1080              Case "G":
1090                  If tb!Valid = 1 Then
1100                      chkSections(7).Value = 1
1110                      chkSections(7).Enabled = False
1120                  End If
1130              Case "Y":
1140                  If tb!Valid = 1 Then
1150                      chkSections(8).Value = 1
1160                      chkSections(8).Enabled = False
1170                  End If
1180              Case "O":
1190                  If tb!Valid = 1 Then
1200                      chkSections(9).Value = 1
1210                      chkSections(9).Enabled = False
1220                  End If
1230              Case "B":
1240                  If tb!Valid = 1 Then
1250                      chkSections(10).Value = 1
1260                      chkSections(10).Enabled = False
1270                  End If
1280              End Select
1290              tb.MoveNext
1300          Wend
1310      End If

1320      Exit Sub


1330      Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

1340      intEL = Erl
1350      strES = Err.Description
1360      LogError "frmValidateAll", "Form_Load", intEL, strES, sql

End Sub
