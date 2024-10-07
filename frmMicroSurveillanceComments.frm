VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMicroSurveillanceComments 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Micro Surveillance Comments"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   9405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab tabComments 
      Height          =   5355
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   9446
      _Version        =   393216
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      BackColor       =   -2147483644
      TabCaption(0)   =   "Demographic"
      TabPicture(0)   =   "frmMicroSurveillanceComments.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraComments(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "General"
      TabPicture(1)   =   "frmMicroSurveillanceComments.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraComments(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Scientist"
      TabPicture(2)   =   "frmMicroSurveillanceComments.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraComments(2)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Consultant"
      TabPicture(3)   =   "frmMicroSurveillanceComments.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraComments(3)"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "CDiff"
      TabPicture(4)   =   "frmMicroSurveillanceComments.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "fraComments(4)"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "CSFFluid"
      TabPicture(5)   =   "frmMicroSurveillanceComments.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "fraComments(5)"
      Tab(5).ControlCount=   1
      Begin VB.Frame fraComments 
         BackColor       =   &H00404000&
         Height          =   4935
         Index           =   5
         Left            =   -74940
         TabIndex        =   11
         Top             =   360
         Width           =   8895
         Begin VB.TextBox txtComments 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   4395
            Index           =   5
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   12
            Top             =   300
            Width           =   8415
         End
      End
      Begin VB.Frame fraComments 
         BackColor       =   &H00404000&
         Height          =   4935
         Index           =   4
         Left            =   -74940
         TabIndex        =   9
         Top             =   360
         Width           =   8895
         Begin VB.TextBox txtComments 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   4395
            Index           =   4
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   10
            Top             =   300
            Width           =   8415
         End
      End
      Begin VB.Frame fraComments 
         BackColor       =   &H00404000&
         Height          =   4935
         Index           =   3
         Left            =   -74940
         TabIndex        =   7
         Top             =   360
         Width           =   8895
         Begin VB.TextBox txtComments 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   4395
            Index           =   3
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   8
            Top             =   300
            Width           =   8415
         End
      End
      Begin VB.Frame fraComments 
         BackColor       =   &H00404000&
         Height          =   4935
         Index           =   2
         Left            =   -74940
         TabIndex        =   5
         Top             =   360
         Width           =   8895
         Begin VB.TextBox txtComments 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   4395
            Index           =   2
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   6
            Top             =   300
            Width           =   8415
         End
      End
      Begin VB.Frame fraComments 
         BackColor       =   &H00404000&
         Height          =   4935
         Index           =   1
         Left            =   -74940
         TabIndex        =   3
         Top             =   360
         Width           =   8895
         Begin VB.TextBox txtComments 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   4395
            Index           =   1
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   4
            Top             =   300
            Width           =   8415
         End
      End
      Begin VB.Frame fraComments 
         BackColor       =   &H00404000&
         Height          =   4935
         Index           =   0
         Left            =   60
         TabIndex        =   1
         Top             =   360
         Width           =   8895
         Begin VB.TextBox txtComments 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   4395
            Index           =   0
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   2
            Top             =   300
            Width           =   8415
         End
      End
   End
End
Attribute VB_Name = "frmMicroSurveillanceComments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private m_SampleID As String

Private Sub Form_Load()

10    On Error GoTo Form_Load_Error

20    With frmMicroSurveillanceComments
30        .Top = (Screen.Height - .Height) / 2
40        .Left = (Screen.Width - .Width) / 2
50    End With
60    ShowComments

70    Exit Sub

Form_Load_Error:

      Dim strES As String
      Dim intEL As Integer

80    intEL = Erl
90    strES = Err.Description
100   LogError "frmMicroSurveillanceComments", "Form_Load", intEL, strES

End Sub

Private Sub ShowComments()
      Dim OB As Observation
      Dim OBS As New Observations
      Dim i As Integer
10    On Error GoTo ShowComments_Error

20    Set OBS = OBS.Load(Me.SampleID, _
                         "Demographic", "MicroGeneral", "MicroCS", _
                         "MicroConsultant", "MicroCDiff", "CSFFluid")
30    If Not OBS Is Nothing Then
40        If OBS.Count > 0 Then
50            For Each OB In OBS
60                Select Case OB.Discipline
                  Case "Demographic"
70                    txtComments(0).Text = OB.Comment
80                Case "MicroGeneral"
90                    txtComments(1).Text = OB.Comment
100               Case "MicroCS"
110                   txtComments(2).Text = OB.Comment
                      
120               Case "MicroConsultant"
130                   txtComments(3).Text = OB.Comment
                      
140               Case "MICROCDIFF"
150                   txtComments(4).Text = OB.Comment
160               Case "CSFFluid"
170                   txtComments(5).Text = OB.Comment
180               End Select
190           Next
200       End If
210   End If
220   For i = 0 To 5
230       If txtComments(i).Text = "" Then
240           tabComments.TabVisible(i) = False
250       End If
260   Next i

270   Exit Sub

ShowComments_Error:

      Dim strES As String
      Dim intEL As Integer

280   intEL = Erl
290   strES = Err.Description
300   LogError "frmMicroSurveillanceComments", "ShowComments", intEL, strES
End Sub

Public Property Get SampleID() As String

10    SampleID = m_SampleID

End Property

Public Property Let SampleID(ByVal sSampleID As String)

10    m_SampleID = sSampleID

End Property

