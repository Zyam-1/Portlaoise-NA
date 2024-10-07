VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmFrozen 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Frozen Section Count"
   ClientHeight    =   4170
   ClientLeft      =   2580
   ClientTop       =   1695
   ClientWidth     =   5130
   ControlBox      =   0   'False
   Icon            =   "frmFrozen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   705
      Left            =   3780
      Picture         =   "frmFrozen.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1710
      Width           =   1245
   End
   Begin VB.Frame Frame2 
      Height          =   1545
      Left            =   405
      TabIndex        =   7
      Top             =   2340
      Width           =   2565
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Number of Blocks between the above dates"
         Height          =   435
         Left            =   240
         TabIndex        =   9
         Top             =   270
         Width           =   2055
      End
      Begin VB.Label lblBCount 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   660
         TabIndex        =   8
         Top             =   840
         Width           =   1125
      End
   End
   Begin MSComCtl2.DTPicker CalTo 
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   240
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
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
      CustomFormat    =   "dd/MMM/yyyy"
      Format          =   59965443
      CurrentDate     =   37643
   End
   Begin MSComCtl2.DTPicker CalFrom 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
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
      CustomFormat    =   "dd/MMM/yyyy"
      Format          =   59965443
      CurrentDate     =   37643
   End
   Begin VB.Frame Frame1 
      Height          =   1545
      Left            =   390
      TabIndex        =   2
      Top             =   840
      Width           =   2565
      Begin VB.Label lblCount 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   660
         TabIndex        =   4
         Top             =   840
         Width           =   1125
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Number of Frozen sections between the above dates"
         Height          =   435
         Left            =   240
         TabIndex        =   3
         Top             =   270
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "C&alculate"
      Height          =   705
      Left            =   3780
      Picture         =   "frmFrozen.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2520
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   705
      Left            =   3780
      Picture         =   "frmFrozen.frx":0A56
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3330
      Width           =   1245
   End
End
Attribute VB_Name = "frmFrozen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Calculate()

          Dim sn As Recordset
          Dim sql As String


10        On Error GoTo Calculate_Error

20        lblCount = "0"
30        lblCount.Refresh
40        lblBCount = "0"
50        lblBCount.Refresh

60        sql = "select * from histospecimen, demographics where " & _
                "demographics.rundate between '" & _
                Format(calFrom, "dd/mmm/yyyy") & " 00:00:00' and '" & _
                Format(calTo, "dd/mmm/yyyy") & " 23:59:59' and histospecimen.sampleid = demographics.sampleid"
70        Set sn = New Recordset
80        RecOpenServer 0, sn, sql
90        Do While Not sn.EOF
100           lblCount = lblCount + Val(sn!fs & "")
110           lblBCount = lblBCount + Val(sn!blocks & "")
120           sn.MoveNext
130       Loop

140       Exit Sub

Calculate_Error:

          Dim strES As String
          Dim intEL As Integer

150       intEL = Erl
160       strES = Err.Description
170       LogError "frmFrozen", "calculate", intEL, strES, sql

End Sub

Private Sub cmdCalc_Click()

10        On Error GoTo cmdCalc_Click_Error

20        Calculate

30        Exit Sub

cmdCalc_Click_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmFrozen", "cmdCalc_Click", intEL, strES


End Sub

Private Sub cmdCancel_Click()

10        Unload Me

End Sub


Private Sub cmdPrint_Click()
          Dim n As Integer


10        On Error GoTo cmdPrint_Click_Error

20        Printer.Print Tab(20); "Report between " & Format(calFrom, "dd/MMM/yyyy") & " and " & Format(calTo, "dd/MMM/yyyy")
30        For n = 0 To 3
40            Printer.Print
50        Next
60        Printer.Font.Bold = True
70        Printer.Print Tab(20); "No. of Frozen sections : " & lblCount
80        For n = 0 To 3
90            Printer.Print
100       Next
110       Printer.Print Tab(20); "No. of Blocks          : " & lblBCount
120       Printer.Font.Bold = False
130       For n = 0 To 3
140           Printer.Print
150       Next
160       Printer.Print Tab(20); " ------ End of Report -------"
170       Printer.EndDoc


180       Exit Sub

cmdPrint_Click_Error:

          Dim strES As String
          Dim intEL As Integer



190       intEL = Erl
200       strES = Err.Description
210       LogError "frmFrozen", "cmdPrint_Click", intEL, strES


End Sub

Private Sub Form_Activate()

10        On Error GoTo Form_Activate_Error

20        Calculate

30        Exit Sub

Form_Activate_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmFrozen", "Form_Activate", intEL, strES


End Sub

'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' DateTime  : 29/06/2007 16:32
' Author    : Myles
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        calFrom = DateAdd("m", -1, Now)
30        calTo = Now

40        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer



50        intEL = Erl
60        strES = Err.Description
70        LogError "frmFrozen", "Form_Load", intEL, strES


End Sub

