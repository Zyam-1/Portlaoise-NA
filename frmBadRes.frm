VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmBadRes 
   Caption         =   "NetAcquire - Bad Results"
   ClientHeight    =   6960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10875
   Icon            =   "frmBadRes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   10875
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid grdBad 
      Height          =   5280
      Left            =   135
      TabIndex        =   17
      Top             =   1620
      Width           =   10725
      _ExtentX        =   18918
      _ExtentY        =   9313
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      AllowUserResizing=   1
      FormatString    =   $"frmBadRes.frx":030A
   End
   Begin VB.Frame Frame2 
      Caption         =   "Discipline"
      Height          =   1455
      Left            =   5220
      TabIndex        =   11
      Top             =   90
      Width           =   5505
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Exit"
         Height          =   735
         Left            =   3915
         Picture         =   "frmBadRes.frx":03CC
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   630
         Width           =   1500
      End
      Begin VB.OptionButton optDept 
         Caption         =   "Haematology"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   16
         Top             =   315
         Value           =   -1  'True
         Width           =   1275
      End
      Begin VB.OptionButton optDept 
         Caption         =   "Biochemistry"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   15
         Tag             =   "Bio"
         Top             =   555
         Width           =   1320
      End
      Begin VB.OptionButton optDept 
         Caption         =   "Coagulation"
         Height          =   195
         Index           =   2
         Left            =   1665
         TabIndex        =   14
         Tag             =   "Coag"
         Top             =   315
         Width           =   1695
      End
      Begin VB.OptionButton optDept 
         Caption         =   "Endocrinology"
         Height          =   195
         Index           =   3
         Left            =   1665
         TabIndex        =   13
         Tag             =   "End"
         Top             =   540
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.OptionButton optDept 
         Caption         =   "Immunology"
         Height          =   195
         Index           =   4
         Left            =   1665
         TabIndex        =   12
         Tag             =   "Imm"
         Top             =   765
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.OptionButton optDept 
         Caption         =   "External"
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   19
         Tag             =   "e"
         Top             =   810
         Width           =   1725
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Between Dates"
      Height          =   1470
      Left            =   135
      TabIndex        =   0
      Top             =   90
      Width           =   5175
      Begin VB.CommandButton cmdRecalc 
         Caption         =   "Start"
         Height          =   660
         Left            =   3630
         MaskColor       =   &H8000000F&
         Picture         =   "frmBadRes.frx":06D6
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   165
         Width           =   1275
      End
      Begin Threed.SSOption optBetween 
         Height          =   225
         Index           =   6
         Left            =   405
         TabIndex        =   2
         Top             =   870
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Today"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Value           =   -1  'True
      End
      Begin Threed.SSOption optBetween 
         Height          =   225
         Index           =   3
         Left            =   2370
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   885
         Width           =   1155
         _Version        =   65536
         _ExtentX        =   2037
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Last Quarter"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin Threed.SSOption optBetween 
         Height          =   225
         Index           =   4
         Left            =   3600
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   885
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Last Full Quarter"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption optBetween 
         Height          =   225
         Index           =   2
         Left            =   1215
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1170
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Last Full Month"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption optBetween 
         Height          =   225
         Index           =   5
         Left            =   3600
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1125
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Year to Date"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption optBetween 
         Height          =   225
         Index           =   1
         Left            =   75
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1170
         Width           =   1065
         _Version        =   65536
         _ExtentX        =   1879
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Last Month"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin Threed.SSOption optBetween 
         Height          =   225
         Index           =   0
         Left            =   1215
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   900
         Width           =   1065
         _Version        =   65536
         _ExtentX        =   1879
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Last Week"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   285
         Left            =   300
         TabIndex        =   9
         Top             =   300
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   503
         _Version        =   393216
         Format          =   118947841
         CurrentDate     =   37019
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   315
         Left            =   2010
         TabIndex        =   10
         Top             =   300
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         Format          =   118947841
         CurrentDate     =   37019
      End
   End
End
Attribute VB_Name = "frmBadRes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim calFrom As String
Dim calTo As String
Dim Dept As String

Private Sub cmdCancel_Click()

10        Unload Me

End Sub

Private Sub cmdRecalc_Click()

10        On Error GoTo cmdRecalc_Click_Error

20        calFrom = Format(dtFrom, "dd/MMM/yyyy 00:00:00")
30        calTo = Format(dtTo, "dd/MMM/yyyy 23:59:59")


40        FillG

50        cmdRecalc.Visible = False

60        Exit Sub

cmdRecalc_Click_Error:

          Dim strES As String
          Dim intEL As Integer

70        intEL = Erl
80        strES = Err.Description
90        LogError "frmBadRes", "cmdRecalc_Click", intEL, strES


End Sub

Private Sub dtFrom_Change()

10        On Error GoTo dtFrom_Change_Error

20        cmdRecalc.Visible = True

30        Exit Sub

dtFrom_Change_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmBadRes", "dtFrom_Change", intEL, strES


End Sub

Private Sub dtTo_Change()

10        On Error GoTo dtTo_Change_Error

20        cmdRecalc.Visible = True

30        Exit Sub

dtTo_Change_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmBadRes", "dtTo_Change", intEL, strES


End Sub

Private Sub FillG()

          Dim sql As String
          Dim tb As Recordset
          Dim sn As Recordset
          Dim Comment As String
          Dim Str As String
          Dim Obs As Observations
          Dim Disc As String

10        On Error GoTo FillG_Error

20        ClearFGrid grdBad

30        If optDept(0) Then
40            Disc = "Haematology"
50            sql = "SELECT * from haemresults WHERE cbad = 1 and " & _
                    "Rundatetime between '" & calFrom & "' and '" & calTo & "'"
60        ElseIf optDept(1) Then
70            Disc = "Biochemistry"
80            sql = "SELECT * from Bioresults WHERE code = '" & SysOptBioCodeForBad(0) & "' and " & _
                    "Runtime between '" & calFrom & "' and '" & calTo & "'"
90        ElseIf optDept(2) Then
100           Disc = "Coagulation"
110           sql = "SELECT * from coagresults WHERE code = '" & SysOptCBad(0) & "' and " & _
                    "rundate between '" & calFrom & "' and '" & calTo & "'"
120       ElseIf optDept(5) Then
130           Disc = ""
140           sql = "SELECT * from extresults WHERE analyte = '" & SysOptEBad(0) & "' and " & _
                    "sentdate between '" & calFrom & "' and '" & calTo & "'"
150       End If
160       Set tb = New Recordset
170       RecOpenServer 0, tb, sql
180       Do While Not tb.EOF
190           Comment = ""

200           Set Obs = New Observations
210           Set Obs = Obs.Load(tb!SampleID, Disc)
220           If Not Obs Is Nothing Then
230               Comment = Obs.Item(1).Comment
240           End If
250           sql = "SELECT * FROM Demographics WHERE " & _
                    "SampleID = '" & tb!SampleID & "'"

260           Set sn = New Recordset
270           RecOpenServer 0, sn, sql
280           If Not sn.EOF Then
290               Str = tb!SampleID & vbTab & Trim(sn!PatName & "") & vbTab & _
                        "BAD" & vbTab & Comment & vbTab & Trim(sn!Ward & "") & _
                        vbTab & Trim(sn!Clinician & "") & vbTab & Trim(sn!GP) & ""
300               grdBad.AddItem Str
310           End If
320           tb.MoveNext
330       Loop

340       FixG grdBad

350       If grdBad.TextMatrix(1, 0) <> "" Then
360           grdBad.AddItem ""
370           grdBad.AddItem "Total" & vbTab & Val(grdBad.Rows - 2)
380       End If


390       Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

400       intEL = Erl
410       strES = Err.Description
420       LogError "frmBadRes", "FillG", intEL, strES, sql


End Sub

Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        dtTo = Format(Now, "dd/MM/yyyy")
30        dtFrom = Format(Now, "dd/MM/yyyy")
40        optDept_Click (0)

50        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmBadRes", "Form_Load", intEL, strES


End Sub



Private Sub optBetween_Click(Index As Integer, Value As Integer)
          Dim upto As String

10        On Error GoTo optBetween_Click_Error

20        dtFrom = BetweenDates(Index, upto)
30        dtTo = upto

40        cmdRecalc.Visible = True


50        Exit Sub

optBetween_Click_Error:

          Dim strES As String
          Dim intEL As Integer

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmBadRes", "optBetween_Click", intEL, strES


End Sub

Private Sub optDept_Click(Index As Integer)

10        On Error GoTo optDept_Click_Error

20        cmdRecalc.Visible = True
30        Dept = optDept(Index).Tag

40        Exit Sub

optDept_Click_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmBadRes", "optDept_Click", intEL, strES


End Sub
