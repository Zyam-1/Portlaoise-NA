VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmMicroGeneral 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdXL 
      Caption         =   "Export to Excel"
      Height          =   735
      Left            =   4650
      Picture         =   "frmMicroGeneral.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   825
      Left            =   4560
      Picture         =   "frmMicroGeneral.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3720
      Width           =   1155
   End
   Begin MSComCtl2.DTPicker dtTo 
      Height          =   315
      Left            =   2820
      TabIndex        =   2
      Top             =   270
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   59179009
      CurrentDate     =   39733
   End
   Begin MSComCtl2.DTPicker dtFrom 
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Top             =   270
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   59179009
      CurrentDate     =   39733
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   3825
      Left            =   390
      TabIndex        =   0
      Top             =   720
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   6747
      _Version        =   393216
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      AllowUserResizing=   1
      FormatString    =   "<Site                                                   |<Samples        "
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4470
      TabIndex        =   6
      Top             =   1470
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Between Dates"
      Height          =   195
      Left            =   150
      TabIndex        =   3
      Top             =   330
      Width           =   1095
   End
End
Attribute VB_Name = "frmMicroGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub DTPicker2_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub

Private Sub cmdCancel_Click()

10        Unload Me

End Sub

Private Sub cmdXL_Click()

10        ExportFlexGrid g, Me

End Sub


Private Sub dtFrom_CloseUp()

10        FillSites

End Sub

Private Sub dtTo_CloseUp()

10        FillSites

End Sub

Private Sub Form_Activate()

10        FillSites

End Sub

Private Sub FillSites()

          Dim sql As String
          Dim tbSite As Recordset
          Dim tb As Recordset
          Dim FromDate As String
          Dim ToDate As String

10        On Error GoTo FillSites_Error

20        FromDate = Format$(dtFrom, "dd/MMM/yyyy")
30        ToDate = Format$(dtTo, "dd/MMM/yyyy") & " 23:59:59"

40        g.Rows = 2
50        g.AddItem ""
60        g.RemoveItem 1

70        sql = "SELECT Text FROM Lists WHERE " & _
                "ListType = 'SI' " & _
                "ORDER BY ListOrder"
80        Set tbSite = New Recordset
90        RecOpenServer 0, tbSite, sql
100       Do While Not tbSite.EOF
110           sql = "SELECT COUNT(D.SampleID) AS Tot FROM " & _
                    "Demographics D, MicroSiteDetails M WHERE " & _
                    "D.SampleID = M.SampleID " & _
                    "AND D.RunDate BETWEEN '" & FromDate & "' AND '" & ToDate & "' " & _
                    "AND M.Site LIKE '" & tbSite!Text & "'"
120           Set tb = New Recordset
130           RecOpenServer 0, tb, sql
140           If tb!Tot <> 0 Then
150               g.AddItem tbSite!Text & vbTab & tb!Tot
160           End If
170           tbSite.MoveNext
180       Loop

190       If g.Rows > 2 Then
200           g.RemoveItem 1
210       End If

220       Exit Sub

FillSites_Error:

          Dim strES As String
          Dim intEL As Integer

230       intEL = Erl
240       strES = Err.Description
250       LogError "frmMicroGeneral", "FillSites", intEL, strES, sql

End Sub


Private Sub Form_Load()

10        dtFrom = Format$(Now - 7, "dd/MM/yyyy")
20        dtTo = Format$(Now, "dd/MM/yyyy")

End Sub


