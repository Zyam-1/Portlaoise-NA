VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmMicroTotals 
   Caption         =   "NetAcquire - Faeces Totals"
   ClientHeight    =   4110
   ClientLeft      =   1935
   ClientTop       =   2115
   ClientWidth     =   8220
   Icon            =   "frmMicroTotals.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   4110
   ScaleWidth      =   8220
   Begin VB.Frame Frame2 
      Caption         =   "Total Faeces Samples"
      Height          =   975
      Left            =   420
      TabIndex        =   29
      Top             =   870
      Width           =   2085
      Begin VB.Label lTotFaeces 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   390
         TabIndex        =   30
         Top             =   390
         Width           =   1260
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Occult Blood"
      Height          =   975
      Left            =   3300
      TabIndex        =   22
      Top             =   870
      Width           =   2715
      Begin VB.Label lOB 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Index           =   0
         Left            =   120
         TabIndex        =   28
         Top             =   540
         Width           =   780
      End
      Begin VB.Label lOB 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Index           =   1
         Left            =   960
         TabIndex        =   27
         Top             =   540
         Width           =   780
      End
      Begin VB.Label lOB 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Index           =   2
         Left            =   1800
         TabIndex        =   26
         Top             =   540
         Width           =   780
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "First"
         Height          =   195
         Left            =   330
         TabIndex        =   25
         Top             =   270
         Width           =   285
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Second"
         Height          =   195
         Left            =   1020
         TabIndex        =   24
         Top             =   270
         Width           =   555
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Third"
         Height          =   195
         Left            =   1950
         TabIndex        =   23
         Top             =   270
         Width           =   360
      End
   End
   Begin VB.CommandButton bCalc 
      Caption         =   "Start"
      Height          =   705
      Left            =   5040
      Picture         =   "frmMicroTotals.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   690
      Left            =   6510
      Picture         =   "frmMicroTotals.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2730
      Width           =   1425
   End
   Begin MSComCtl2.DTPicker calTo 
      Height          =   315
      Left            =   3240
      TabIndex        =   9
      Top             =   180
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   59637761
      CurrentDate     =   36951
   End
   Begin MSComCtl2.DTPicker calFrom 
      Height          =   315
      Left            =   1290
      TabIndex        =   8
      Top             =   180
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   59637761
      CurrentDate     =   36951
   End
   Begin VB.Label lblGiardia 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4950
      TabIndex        =   32
      Top             =   3480
      Width           =   1080
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Giardia Lambila"
      Height          =   195
      Left            =   3795
      TabIndex        =   31
      Top             =   3540
      Width           =   1080
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "H. Pylori"
      Height          =   195
      Left            =   360
      TabIndex        =   21
      Top             =   3180
      Width           =   585
   End
   Begin VB.Label lblHPylori 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1020
      TabIndex        =   20
      Top             =   3120
      Width           =   1485
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Cypto"
      Height          =   195
      Left            =   4470
      TabIndex        =   19
      Top             =   3150
      Width           =   405
   End
   Begin VB.Label lblCrypto 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4950
      TabIndex        =   18
      Top             =   3090
      Width           =   1080
   End
   Begin VB.Label lblRedSub 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1020
      TabIndex        =   16
      Top             =   2370
      Width           =   1485
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Red Sub"
      Height          =   195
      Left            =   330
      TabIndex        =   15
      Top             =   2430
      Width           =   630
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Adeno Virus"
      Height          =   195
      Left            =   4020
      TabIndex        =   13
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label lblAdeno 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4950
      TabIndex        =   12
      Top             =   2340
      Width           =   1080
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "To"
      Height          =   195
      Left            =   3000
      TabIndex        =   11
      Top             =   240
      Width           =   195
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "From"
      Height          =   195
      Left            =   870
      TabIndex        =   10
      Top             =   240
      Width           =   345
   End
   Begin VB.Label lblCultured 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1020
      TabIndex        =   7
      Top             =   1980
      Width           =   1485
   End
   Begin VB.Label lblRota 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4950
      TabIndex        =   6
      Top             =   1950
      Width           =   1080
   End
   Begin VB.Label lOP 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4950
      TabIndex        =   5
      Top             =   2715
      Width           =   1080
   End
   Begin VB.Label lblCDiff 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1020
      TabIndex        =   4
      Top             =   2745
      Width           =   1485
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "C Difficile"
      Height          =   195
      Left            =   285
      TabIndex        =   3
      Top             =   2790
      Width           =   660
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Ova/Parasites"
      Height          =   195
      Left            =   3855
      TabIndex        =   2
      Top             =   2790
      Width           =   1020
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Rota Virus"
      Height          =   195
      Left            =   4140
      TabIndex        =   1
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cultured"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   2025
      Width           =   585
   End
End
Attribute VB_Name = "frmMicroTotals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Calc()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo Calc_Error

20        lblCDiff = "0"
30        lOP = "0"
40        lOB(0) = "0"
50        lOB(1) = "0"
60        lOB(2) = "0"
70        lblRota = "0"
80        lblAdeno = 20
90        lblCultured = "0"
100       lblRedSub = "0"
110       lblCrypto = "0"
120       lblGiardia = "0"

130       sql = "SELECT COUNT(F.SampleID) AS Tot FROM Demographics AS D, Faeces AS F WHERE " & _
                "RunDate BETWEEN '" & Format(calFrom, "dd/MMM/yyyy") & "' AND '" & Format(calTo, "dd/MMM/yyyy") & "' " & _
                "AND D.SampleID = F.SampleID"
140       Set tb = New Recordset
150       RecOpenServer 0, tb, sql
160       lTotFaeces = Format(tb!Tot)

170       sql = "SELECT COUNT(F.OB0) AS Tot FROM Demographics D, Faeces F WHERE " & _
                "RunDate BETWEEN '" & Format(calFrom, "dd/MMM/yyyy") & "' AND '" & Format(calTo, "dd/MMM/yyyy") & "' " & _
                "AND D.SampleID = F.SampleID " & _
                "AND (COALESCE(F.OB0, '') <> '')"
180       Set tb = New Recordset
190       RecOpenServer 0, tb, sql
200       lOB(0) = Format(tb!Tot)

210       sql = "SELECT COUNT(F.OB1) AS Tot FROM Demographics D, Faeces F WHERE " & _
                "RunDate BETWEEN '" & Format(calFrom, "dd/MMM/yyyy") & "' AND '" & Format(calTo, "dd/MMM/yyyy") & "' " & _
                "AND D.SampleID = F.SampleID " & _
                "AND (COALESCE(F.OB1, '') <> '')"
220       Set tb = New Recordset
230       RecOpenServer 0, tb, sql
240       lOB(1) = Format(tb!Tot)

250       sql = "SELECT COUNT(F.OB2) AS Tot FROM Demographics D, Faeces F WHERE " & _
                "RunDate BETWEEN '" & Format(calFrom, "dd/MMM/yyyy") & "' AND '" & Format(calTo, "dd/MMM/yyyy") & "' " & _
                "AND D.SampleID = F.SampleID " & _
                "AND (COALESCE(F.OB2, '') <> '')"
260       Set tb = New Recordset
270       RecOpenServer 0, tb, sql
280       lOB(2) = Format(tb!Tot)

290       sql = "SELECT COUNT(F.ToxinAB) AS Tot FROM Demographics D, Faeces F WHERE " & _
                "RunDate BETWEEN '" & Format(calFrom, "dd/MMM/yyyy") & "' AND '" & Format(calTo, "dd/MMM/yyyy") & "' " & _
                "AND D.SampleID = F.SampleID " & _
                "AND (COALESCE(F.ToxinAB, '') <> '')"
300       Set tb = New Recordset
310       RecOpenServer 0, tb, sql
320       lblCDiff = Format(tb!Tot)

330       sql = "SELECT COUNT(F.SampleID) AS Tot FROM Demographics D, Faeces F WHERE " & _
                "RunDate BETWEEN '" & Format(calFrom, "dd/MMM/yyyy") & "' AND '" & Format(calTo, "dd/MMM/yyyy") & "' " & _
                "AND D.SampleID = F.SampleID " & _
                "AND ( (F.OP0 IS NOT NULL AND CAST(F.OP0 AS nvarchar(50)) <> '') " & _
                "   OR (F.OP1 IS NOT NULL AND CAST(F.OP1 AS nvarchar(50)) <> '') " & _
                "   OR (F.OP2 IS NOT NULL AND CAST(F.OP2 AS nvarchar(50)) <> '') )"
340       Set tb = New Recordset
350       RecOpenServer 0, tb, sql
360       lOP = Format(tb!Tot)

370       sql = "SELECT COUNT(F.SampleID) AS Tot FROM Demographics D, Faeces F WHERE " & _
                "RunDate BETWEEN '" & Format(calFrom, "dd/MMM/yyyy") & "' AND '" & Format(calTo, "dd/MMM/yyyy") & "' " & _
                "AND D.SampleID = F.SampleID " & _
                "AND (COALESCE(F.Rota, '') <> '')"
380       Set tb = New Recordset
390       RecOpenServer 0, tb, sql
400       lblRota = Format(tb!Tot)

410       sql = "SELECT COUNT(F.SampleID) AS Tot FROM Demographics D, Faeces F WHERE " & _
                "RunDate BETWEEN '" & Format(calFrom, "dd/MMM/yyyy") & "' AND '" & Format(calTo, "dd/MMM/yyyy") & "' " & _
                "AND D.SampleID = F.SampleID " & _
                "AND (COALESCE(F.Adeno, '') <> '')"
420       Set tb = New Recordset
430       RecOpenServer 0, tb, sql
440       lblAdeno = Format(tb!Tot)

450       sql = "SELECT COUNT(F.SampleID) AS Tot FROM Demographics D, Faeces F WHERE " & _
                "RunDate BETWEEN '" & Format(calFrom, "dd/MMM/yyyy") & "' AND '" & Format(calTo, "dd/MMM/yyyy") & "' " & _
                "AND D.SampleID = F.SampleID " & _
                "AND (COALESCE(F.Cryptosporidium, '') <> '') "
460       Set tb = New Recordset
470       RecOpenServer 0, tb, sql
480       lblCrypto = Format(tb!Tot)

490       sql = "SELECT COUNT(F.SampleID) AS Tot FROM Demographics D, Faeces F WHERE " & _
                "RunDate BETWEEN '" & Format(calFrom, "dd/MMM/yyyy") & "' AND '" & Format(calTo, "dd/MMM/yyyy") & "' " & _
                "AND D.SampleID = F.SampleID " & _
                "AND (COALESCE(F.GiardiaLambila, '') <> '') "
500       Set tb = New Recordset
510       RecOpenServer 0, tb, sql
520       lblGiardia = Format(tb!Tot)

530       sql = "SELECT COUNT(F.SampleID) AS Tot FROM Demographics D, Faeces F WHERE " & _
                "RunDate BETWEEN '" & Format(calFrom, "dd/MMM/yyyy") & "' AND '" & Format(calTo, "dd/MMM/yyyy") & "' " & _
                "AND D.SampleID = F.SampleID " & _
                "AND (COALESCE(F.HPylori, '') <> '')"
540       Set tb = New Recordset
550       RecOpenServer 0, tb, sql
560       lblHPylori = Format(tb!Tot)

570       sql = "SELECT COUNT(G.SampleID) AS Tot FROM Demographics D, GenericResults G WHERE " & _
                "RunDate BETWEEN '" & Format(calFrom, "dd/MMM/yyyy") & "' AND '" & Format(calTo, "dd/MMM/yyyy") & "' " & _
                "AND D.SampleID = G.SampleID " & _
                "AND G.TestName = 'RedSub'"
580       Set tb = New Recordset
590       RecOpenServer 0, tb, sql
600       lblRedSub = Format(tb!Tot)

610       sql = "SELECT COUNT(DISTINCT(I.SampleID)) AS Tot FROM Demographics D, Isolates I, Faeces F WHERE " & _
                "RunDate BETWEEN '" & Format(calFrom, "dd/MMM/yyyy") & "' AND '" & Format(calTo, "dd/MMM/yyyy") & "' " & _
                "AND D.SampleID = I.SampleID AND F.SampleID = I.SampleID"
620       Set tb = New Recordset
630       RecOpenServer 0, tb, sql
640       lblCultured = Format(tb!Tot)

650       Exit Sub

Calc_Error:

          Dim strES As String
          Dim intEL As Integer

660       intEL = Erl
670       strES = Err.Description
680       LogError "frmMicroTotals", "Calc", intEL, strES, sql

End Sub

Private Sub bCalc_Click()

10        Calc

End Sub

Private Sub calFrom_CloseUp()

10        Calc

End Sub

Private Sub calTo_CloseUp()

10        Calc

End Sub

Private Sub cmdCancel_Click()

10        Unload Me

End Sub

Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        calTo = Format(Now, "dd/mmm/yyyy")
30        calFrom = Format(Now - 365, "dd/mmm/yyyy")

40        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer



50        intEL = Erl
60        strES = Err.Description
70        LogError "frmMicroTotals", "Form_Load", intEL, strES


End Sub


Private Sub Label5_Click()

End Sub
