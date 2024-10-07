VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmWardViewBloodGas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire 6 - Blood Gas Results"
   ClientHeight    =   5760
   ClientLeft      =   1665
   ClientTop       =   1020
   ClientWidth     =   4890
   Icon            =   "frmWardViewBloodGas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   4890
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCum 
      Caption         =   "Cumulative"
      Height          =   825
      Left            =   270
      Picture         =   "frmWardViewBloodGas.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4770
      Width           =   1365
   End
   Begin VB.CommandButton bPrint 
      Caption         =   "&Print"
      Height          =   825
      Left            =   1710
      Picture         =   "frmWardViewBloodGas.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4770
      Width           =   1365
   End
   Begin VB.CommandButton bexit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   825
      Left            =   3150
      Picture         =   "frmWardViewBloodGas.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4770
      Width           =   1365
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   2205
      Left            =   420
      TabIndex        =   0
      Top             =   2130
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   3889
      _Version        =   393216
      Cols            =   3
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      FormatString    =   "<Analyte     |<Result     |^Normal Range     "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   1875
      Left            =   120
      TabIndex        =   2
      Top             =   150
      Width           =   4485
      _Version        =   65536
      _ExtentX        =   7911
      _ExtentY        =   3307
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   2
      BorderWidth     =   2
      BevelInner      =   1
      Begin VB.ComboBox cmbNo 
         Height          =   315
         Left            =   1080
         TabIndex        =   19
         Top             =   90
         Width           =   3120
      End
      Begin VB.Label Sample 
         AutoSize        =   -1  'True
         Caption         =   "Sample Id"
         Height          =   195
         Left            =   300
         TabIndex        =   17
         Top             =   180
         Width           =   705
      End
      Begin VB.Label lNOPAS 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2970
         TabIndex        =   16
         Top             =   480
         Width           =   1245
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "NOPAS"
         Height          =   195
         Left            =   2400
         TabIndex        =   15
         Top             =   540
         Width           =   555
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Sex"
         Height          =   195
         Left            =   2250
         TabIndex        =   14
         Top             =   1140
         Width           =   270
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Age"
         Height          =   195
         Left            =   3360
         TabIndex        =   13
         Top             =   1110
         Width           =   285
      End
      Begin VB.Label lsex 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2610
         TabIndex        =   12
         Top             =   1110
         Width           =   540
      End
      Begin VB.Label lage 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3675
         TabIndex        =   11
         Top             =   1110
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Address"
         Height          =   195
         Left            =   480
         TabIndex        =   10
         Top             =   1440
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Date of Birth"
         Height          =   195
         Left            =   180
         TabIndex        =   9
         Top             =   1140
         Width           =   885
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "MRN"
         Height          =   195
         Left            =   675
         TabIndex        =   8
         Top             =   510
         Width           =   375
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Left            =   630
         TabIndex        =   7
         Top             =   810
         Width           =   420
      End
      Begin VB.Label lAddress 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1095
         TabIndex        =   6
         Top             =   1440
         Width           =   3120
      End
      Begin VB.Label lDoB 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1080
         TabIndex        =   5
         Top             =   1110
         Width           =   990
      End
      Begin VB.Label lMRN 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1080
         TabIndex        =   4
         Top             =   480
         Width           =   990
      End
      Begin VB.Label lName 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1065
         TabIndex        =   3
         Top             =   780
         Width           =   3150
      End
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1800
      TabIndex        =   22
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "Time Analysed"
      Height          =   255
      Left            =   480
      TabIndex        =   21
      Top             =   4440
      Width           =   1215
   End
End
Attribute VB_Name = "frmWardViewBloodGas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Activated As Boolean

Private Sub bexit_Click()

Unload Me

End Sub


Private Sub cmdCum_Click()

With frmBGHist
  .lblNOPAS = lNOPAS
  .lblName = lName
  .lblMrn = lMRN
  .Show 1
End With

End Sub

Private Sub Form_Activate()

If Activated Then Exit Sub
Activated = True

FillDateTime

End Sub

Private Sub FillDateTime()

Dim tb As New Recordset
Dim SQL As String


SQL = "SELECT distinct sampleid from BGAResults " & _
      "WHERE nopas = '" & lNOPAS & "' " & _
      "order by sampleid desc"
Set tb = New Recordset
RecOpenClient 0, tb, SQL

Do While Not tb.EOF
  cmbNo.AddItem Format(tb!SampleID)
  tb.MoveNext
Loop
FillG

Exit Sub


End Sub

Private Sub Form_Load()

Activated = False

End Sub


Private Sub Form_Unload(Cancel As Integer)

Activated = False

End Sub


Private Sub FillG()

Dim tb As New Recordset
Dim SQL As String
Dim BG As BGAResult
Dim BGs As New BGAResults


bprint.Enabled = True
g.Rows = 2
g.AddItem ""
g.RemoveItem 1

Set BG = BGs.LoadResults(0, cmbNo)
If Not BG Is Nothing Then
  SQL = "SELECT * from BGDefinitions"
  Set tb = New Recordset
  RecOpenClient 0, tb, SQL
  g.AddItem "pH"
  g.AddItem "PO2"
  g.AddItem "PCO2"
  g.AddItem "HCO3"
  g.AddItem "BE"
  g.AddItem "O2SAT"
  g.AddItem "Tot CO2"
  If BG.Valid Then
    g.TextMatrix(2, 1) = BG.pH
    g.TextMatrix(3, 1) = BG.PO2
    g.TextMatrix(4, 1) = BG.PCO2
    g.TextMatrix(5, 1) = BG.HCO3
    g.TextMatrix(6, 1) = BG.BE
    g.TextMatrix(7, 1) = BG.O2SAT
    g.TextMatrix(8, 1) = BG.TotCO2
    If Not tb.EOF Then
      g.TextMatrix(2, 2) = tb!pH & ""
      g.TextMatrix(3, 2) = tb!PO2 & ""
      g.TextMatrix(4, 2) = tb!PCO2 & ""
      g.TextMatrix(5, 2) = tb!HCO3 & ""
      g.TextMatrix(6, 2) = tb!BE & ""
      g.TextMatrix(7, 2) = tb!O2SAT & ""
      g.TextMatrix(8, 2) = tb!TotCO2 & ""
    End If
  Else
    g.TextMatrix(2, 1) = "Not"
    g.TextMatrix(2, 2) = "Validated"
    lblTime = tb!RunDateTime
    bprint.Enabled = False
  End If
End If

If g.Rows > 2 Then
  g.RemoveItem 1
End If


Exit Sub

End Sub




