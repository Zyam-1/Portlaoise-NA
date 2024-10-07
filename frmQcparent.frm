VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.MDIForm frmQCparent 
   BackColor       =   &H8000000F&
   Caption         =   "Quality Control"
   ClientHeight    =   5820
   ClientLeft      =   405
   ClientTop       =   1965
   ClientWidth     =   11625
   Icon            =   "frmQcparent.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Panel3D1 
      Align           =   1  'Align Top
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   0
      ScaleHeight     =   960
      ScaleWidth      =   11565
      TabIndex        =   2
      Top             =   0
      Width           =   11625
      Begin VB.Frame Frame1 
         Caption         =   "Discipline"
         Height          =   825
         Left            =   45
         TabIndex        =   7
         Top             =   45
         Width           =   2625
         Begin VB.OptionButton optCoag 
            Caption         =   "Coagulation"
            Height          =   255
            Left            =   300
            TabIndex        =   9
            Top             =   510
            Width           =   1845
         End
         Begin VB.OptionButton optBio 
            Caption         =   "Biochemistry"
            Height          =   255
            Left            =   300
            TabIndex        =   8
            Top             =   270
            Width           =   1845
         End
      End
      Begin VB.ComboBox lstpara 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3600
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   495
         Width           =   2655
      End
      Begin VB.TextBox tname 
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
         Left            =   3600
         TabIndex        =   3
         Top             =   180
         Width           =   2655
      End
      Begin MSComCtl2.DTPicker calFromDate 
         Height          =   345
         Left            =   7425
         TabIndex        =   10
         Top             =   180
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   59179009
         CurrentDate     =   37112
      End
      Begin MSComCtl2.DTPicker calToDate 
         Height          =   345
         Left            =   9765
         TabIndex        =   11
         Top             =   180
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   59179009
         CurrentDate     =   37112
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Date to"
         Height          =   195
         Left            =   9135
         TabIndex        =   6
         Top             =   225
         Width           =   525
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Date From"
         Height          =   195
         Left            =   6615
         TabIndex        =   5
         Top             =   225
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Parameter"
         Height          =   195
         Left            =   2835
         TabIndex        =   0
         Top             =   540
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Chart"
         Height          =   195
         Left            =   3150
         TabIndex        =   1
         Top             =   225
         Width           =   375
      End
   End
   Begin VB.Menu mwindow 
      Caption         =   "&Window"
      Begin VB.Menu mnew 
         Caption         =   "&New Window"
      End
      Begin VB.Menu m 
         Caption         =   "-"
      End
      Begin VB.Menu marrange 
         Caption         =   "&Cascade"
         Index           =   0
      End
      Begin VB.Menu marrange 
         Caption         =   "Arrange &Horizontal"
         Index           =   1
      End
      Begin VB.Menu marrange 
         Caption         =   "Arrange &Vertical"
         Index           =   2
      End
      Begin VB.Menu marrange 
         Caption         =   "Arrange &Icons"
         Index           =   3
      End
      Begin VB.Menu m1 
         Caption         =   "-"
      End
      Begin VB.Menu mexit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmQCparent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub lstpara_Click()

          Dim f As Form

10        On Error GoTo lstpara_Click_Error

20        If tName = "" Then Exit Sub

30        Set f = New frcControl
40        f.Show

50        Exit Sub

lstpara_Click_Error:

          Dim strES As String
          Dim intEL As Integer



60        intEL = Erl
70        strES = Err.Description
80        LogError "frmQCparent", "lstpara_Click", intEL, strES


End Sub

Private Sub marrange_Click(Index As Integer)

10        On Error GoTo marrange_Click_Error

20        frmQCparent.Arrange Index

30        Exit Sub

marrange_Click_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmQCparent", "marrange_Click", intEL, strES


End Sub

Private Sub MDIForm_Load()

10        On Error GoTo MDIForm_Load_Error

20        calToDate = Format(Now, "dd/MMM/yyyy")
30        calFromDate = Format(Now - 7, "dd/MMM/yyyy")


40        Exit Sub

MDIForm_Load_Error:

          Dim strES As String
          Dim intEL As Integer



50        intEL = Erl
60        strES = Err.Description
70        LogError "frmQCparent", "MDIForm_Load", intEL, strES


End Sub

Private Sub mexit_Click()

10        Unload Me

End Sub

Private Sub mnew_Click()

          Dim f As Form
          Dim n As Long

10        On Error GoTo mnew_Click_Error

20        If lstpara = "" Or tName = "" Then
30            n = iMsg("Both Analyte and Chart number needed.", 48, "NetAcquire")
40            Exit Sub
50        End If

60        Set f = New frcControl
70        f.Show

80        Exit Sub

mnew_Click_Error:

          Dim strES As String
          Dim intEL As Integer



90        intEL = Erl
100       strES = Err.Description
110       LogError "frmQCparent", "mnew_Click", intEL, strES


End Sub


Private Sub Fill_Bio()
          Dim tb As New Recordset
          Dim sql As String



10        On Error GoTo Fill_Bio_Error

20        sql = "SELECT distinct(longname) from biotestdefinitions WHERE inuse = 1"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql

50        lstpara.Clear

60        Do While Not tb.EOF
70            lstpara.AddItem tb!LongName
80            tb.MoveNext
90        Loop

100       tb.Close

110       calToDate = Format(Now, "dd/mmm/yyyy")
120       calFromDate = calToDate - 7




130       Exit Sub

Fill_Bio_Error:

          Dim strES As String
          Dim intEL As Integer



140       intEL = Erl
150       strES = Err.Description
160       LogError "frmQCparent", "Fill_Bio", intEL, strES, sql


End Sub

Private Sub Fill_Coag()
          Dim tb As New Recordset
          Dim sql As String



10        On Error GoTo Fill_Coag_Error

20        sql = "SELECT distinct(testname) from Coagtestdefinitions"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql

50        lstpara.Clear

60        Do While Not tb.EOF
70            lstpara.AddItem tb!TestName
80            tb.MoveNext
90        Loop

100       tb.Close

110       calToDate = Format(Now, "dd/mmm/yyyy")
120       calFromDate = calToDate - 7



130       Exit Sub

Fill_Coag_Error:

          Dim strES As String
          Dim intEL As Integer



140       intEL = Erl
150       strES = Err.Description
160       LogError "frmQCparent", "Fill_Coag", intEL, strES, sql


End Sub

Private Sub optBio_Click()

10        On Error GoTo optBio_Click_Error

20        Fill_Bio

30        Exit Sub

optBio_Click_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmQCparent", "optBio_Click", intEL, strES


End Sub

Private Sub optCoag_Click()

10        On Error GoTo optCoag_Click_Error

20        Fill_Coag

30        Exit Sub

optCoag_Click_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmQCparent", "optCoag_Click", intEL, strES


End Sub
