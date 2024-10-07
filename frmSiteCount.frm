VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmSiteCount 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire"
   ClientHeight    =   7020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   6405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton bcancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   810
      Left            =   4830
      Picture         =   "frmSiteCount.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5850
      Width           =   1155
   End
   Begin VB.CommandButton bCalc 
      Caption         =   "Start"
      Height          =   705
      Left            =   3840
      Picture         =   "frmSiteCount.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   180
      Width           =   735
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   5685
      Left            =   330
      TabIndex        =   3
      Top             =   990
      Width           =   4245
      _ExtentX        =   7488
      _ExtentY        =   10028
      _Version        =   393216
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      ScrollBars      =   2
      AllowUserResizing=   1
      FormatString    =   "<Site                                                    |<Samples          "
   End
   Begin MSComCtl2.DTPicker dtStart 
      Height          =   315
      Left            =   360
      TabIndex        =   2
      Top             =   390
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      _Version        =   393216
      Format          =   59637761
      CurrentDate     =   39731
   End
   Begin MSComCtl2.DTPicker dtStop 
      Height          =   315
      Left            =   1980
      TabIndex        =   1
      Top             =   390
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      _Version        =   393216
      Format          =   59637761
      CurrentDate     =   39731
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Between Dates"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   150
      Width           =   1095
   End
End
Attribute VB_Name = "frmSiteCount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub bCalc_Click()

          Dim sql As String
          Dim tb As Recordset
          Dim n As Integer
          Dim StartDate As String
          Dim StopDate As String

10        On Error GoTo bCalc_Click_Error

20        StartDate = Format(dtStart, "dd/MMM/yyyy")
30        StopDate = Format$(dtStop, "dd/MMM/yyyy") & " 23:59:59"

40        g.Rows = 2
50        g.AddItem ""
60        g.RemoveItem 1
70        Screen.MousePointer = vbHourglass

80        sql = "SELECT DISTINCT CAST(Site AS nvarchar(50)) AS Site FROM MicroSiteDetails"
90        Set tb = New Recordset
100       RecOpenServer 0, tb, sql
110       Do While Not tb.EOF
120           g.AddItem tb!Site & ""
130           tb.MoveNext
140       Loop

150       For n = 1 To g.Rows - 1
160           sql = "SELECT COUNT(M.SampleID) AS Tot FROM MicroSiteDetails M, Demographics D WHERE " & _
                    "Site LIKE '" & g.TextMatrix(n, 0) & "' " & _
                    "AND D.RunDate BETWEEN '" & StartDate & "' AND '" & StopDate & "' " & _
                    "AND D.SampleID = M.SampleID"
170           Set tb = New Recordset
180           RecOpenServer 0, tb, sql
190           g.TextMatrix(n, 1) = Format$(tb!Tot)
200       Next

210       If g.Rows > 2 Then
220           g.RemoveItem 1
230       End If
240       Screen.MousePointer = vbNormal

250       Exit Sub

bCalc_Click_Error:

          Dim strES As String
          Dim intEL As Integer

260       intEL = Erl
270       strES = Err.Description
280       LogError "frmSiteCount", "bCalc_Click", intEL, strES, sql
290       Screen.MousePointer = vbNormal

End Sub

Private Sub bcancel_Click()

10        Unload Me

End Sub


Private Sub Form_Load()

10        dtStart = Format(Now - 30, "dd/MM/yyyy")
20        dtStop = Format(Now, "dd/MM/yyyy")

End Sub


