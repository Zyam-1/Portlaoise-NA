VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmMicroFluidStats 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Fluids"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame6 
      Caption         =   "KOH Preparation"
      Height          =   1425
      Left            =   240
      TabIndex        =   10
      Top             =   1140
      Width           =   4155
      Begin VB.Label lblNoFungalSeen 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2460
         TabIndex        =   14
         Top             =   810
         Width           =   750
      End
      Begin VB.Label lblFungalSeen 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2460
         TabIndex        =   13
         Top             =   420
         Width           =   750
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "No Fungal Elements Seen"
         Height          =   195
         Left            =   525
         TabIndex        =   12
         Top             =   840
         Width           =   1845
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fungal Elements Seen"
         Height          =   195
         Left            =   780
         TabIndex        =   11
         Top             =   450
         Width           =   1590
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Antigen Tests"
      Height          =   1125
      Left            =   240
      TabIndex        =   5
      Top             =   2790
      Width           =   4155
      Begin VB.Label lblPneuATPos 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2820
         TabIndex        =   18
         Top             =   330
         Width           =   1035
      End
      Begin VB.Label lblLegionellaATPos 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2820
         TabIndex        =   17
         Top             =   600
         Width           =   1035
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Positive"
         Height          =   195
         Left            =   3030
         TabIndex        =   16
         Top             =   0
         Width           =   555
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Negative"
         Height          =   195
         Left            =   1620
         TabIndex        =   15
         Top             =   0
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Pneumococcal AT"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1320
      End
      Begin VB.Label lblPneuATNeg 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1470
         TabIndex        =   8
         Top             =   330
         Width           =   1035
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Legionella AT"
         Height          =   195
         Left            =   450
         TabIndex        =   7
         Top             =   630
         Width           =   975
      End
      Begin VB.Label lblLegionellaATNeg 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1470
         TabIndex        =   6
         Top             =   600
         Width           =   1035
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   690
      Left            =   4650
      Picture         =   "frmMicroFluidStats.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3210
      Width           =   1425
   End
   Begin VB.CommandButton bCalc 
      Caption         =   "Start"
      Height          =   690
      Left            =   4650
      Picture         =   "frmMicroFluidStats.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   300
      Width           =   1425
   End
   Begin VB.Frame Frame1 
      Caption         =   "Between Dates"
      Height          =   795
      Left            =   240
      TabIndex        =   0
      Top             =   210
      Width           =   3375
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   315
         Left            =   1800
         TabIndex        =   2
         Top             =   300
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         Format          =   59179009
         CurrentDate     =   39780
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   315
         Left            =   210
         TabIndex        =   1
         Top             =   300
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         Format          =   59179009
         CurrentDate     =   39780
      End
   End
End
Attribute VB_Name = "frmMicroFluidStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Calculate()

          Dim tb As Recordset
          Dim sql As String
          Dim FromDate As String
          Dim ToDate As String

10        On Error GoTo Calculate_Error

20        FromDate = Format$(dtFrom, "dd/MMM/yyyy")
30        ToDate = Format$(dtTo, "dd/MMM/yyyy") & " 23:59"

40        sql = "SELECT COUNT(*) AS Tot FROM GenericResults G, Demographics D WHERE " & _
                "G.SampleID = D.SampleID " & _
                "AND D.RunDate BETWEEN '" & FromDate & "' AND '" & ToDate & "' " & _
                "AND G.TestName = 'FungalElements' " & _
                "AND G.Result = 'Seen'"
50        Set tb = New Recordset
60        RecOpenServer 0, tb, sql
70        lblFungalSeen = tb!Tot

80        sql = "SELECT COUNT(*) AS Tot FROM GenericResults G, Demographics D WHERE " & _
                "G.SampleID = D.SampleID " & _
                "AND D.RunDate BETWEEN '" & FromDate & "' AND '" & ToDate & "' " & _
                "AND G.TestName = 'FungalElements' " & _
                "AND G.Result = 'Not Seen'"
90        Set tb = New Recordset
100       RecOpenServer 0, tb, sql
110       lblNoFungalSeen = tb!Tot

120       sql = "SELECT COUNT(*) AS Tot FROM GenericResults G, Demographics D WHERE " & _
                "G.SampleID = D.SampleID " & _
                "AND D.RunDate BETWEEN '" & FromDate & "' AND '" & ToDate & "' " & _
                "AND G.TestName = 'PneumococcalAT' " & _
                "AND G.Result = 'Positive'"
130       Set tb = New Recordset
140       RecOpenServer 0, tb, sql
150       lblPneuATPos = tb!Tot

160       sql = "SELECT COUNT(*) AS Tot FROM GenericResults G, Demographics D WHERE " & _
                "G.SampleID = D.SampleID " & _
                "AND D.RunDate BETWEEN '" & FromDate & "' AND '" & ToDate & "' " & _
                "AND G.TestName = 'PneumococcalAT' " & _
                "AND G.Result = 'Negative'"
170       Set tb = New Recordset
180       RecOpenServer 0, tb, sql
190       lblPneuATNeg = tb!Tot

200       sql = "SELECT COUNT(*) AS Tot FROM GenericResults G, Demographics D WHERE " & _
                "G.SampleID = D.SampleID " & _
                "AND D.RunDate BETWEEN '" & FromDate & "' AND '" & ToDate & "' " & _
                "AND G.TestName = 'LegionellaAT' " & _
                "AND G.Result = 'Negative'"
210       Set tb = New Recordset
220       RecOpenServer 0, tb, sql
230       lblLegionellaATNeg = tb!Tot

240       sql = "SELECT COUNT(*) AS Tot FROM GenericResults G, Demographics D WHERE " & _
                "G.SampleID = D.SampleID " & _
                "AND D.RunDate BETWEEN '" & FromDate & "' AND '" & ToDate & "' " & _
                "AND G.TestName = 'LegionellaAT' " & _
                "AND G.Result = 'Positive'"
250       Set tb = New Recordset
260       RecOpenServer 0, tb, sql
270       lblLegionellaATPos = tb!Tot

280       Exit Sub

Calculate_Error:

          Dim strES As String
          Dim intEL As Integer

290       intEL = Erl
300       strES = Err.Description
310       LogError "frmMicroFluidStats", "Calculate", intEL, strES, sql

End Sub


Private Sub bCalc_Click()

10        Calculate

End Sub

Private Sub cmdCancel_Click()

10        Unload Me

End Sub

Private Sub Form_Load()

10        dtFrom = Format(Now - 30, "dd/MM/yyyy")
20        dtTo = Format(Now, "dd/MM/yyyy")

End Sub


Private Sub lblLegionellaAT_Click()

End Sub


