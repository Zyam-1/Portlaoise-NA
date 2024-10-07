VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMicroExtRequestTo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10065
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   10065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   3825
      Left            =   150
      TabIndex        =   9
      Top             =   1920
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   6747
      _Version        =   393216
      Cols            =   4
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      FormatString    =   $"frmMicroExtRequestTo.frx":0000
   End
   Begin VB.Frame Frame1 
      Caption         =   "External Laboratory"
      Height          =   1575
      Left            =   150
      TabIndex        =   1
      Top             =   150
      Width           =   5895
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   765
         Left            =   5160
         Picture         =   "frmMicroExtRequestTo.frx":00AE
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   585
      End
      Begin VB.TextBox txtAddress 
         Height          =   285
         Index           =   2
         Left            =   1020
         MaxLength       =   50
         TabIndex        =   7
         Top             =   1140
         Width           =   4005
      End
      Begin VB.TextBox txtAddress 
         Height          =   285
         Index           =   1
         Left            =   1020
         MaxLength       =   50
         TabIndex        =   6
         Top             =   840
         Width           =   4005
      End
      Begin VB.TextBox txtAddress 
         Height          =   285
         Index           =   0
         Left            =   1020
         MaxLength       =   50
         TabIndex        =   5
         Top             =   540
         Width           =   4005
      End
      Begin VB.TextBox txtLabName 
         Height          =   285
         Left            =   1020
         MaxLength       =   30
         TabIndex        =   4
         Top             =   240
         Width           =   4005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Address"
         Height          =   195
         Left            =   390
         TabIndex        =   3
         Top             =   570
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Lab Name"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   270
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   825
      Left            =   9060
      Picture         =   "frmMicroExtRequestTo.frx":1A30
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4905
      Width           =   795
   End
End
Attribute VB_Name = "frmMicroExtRequestTo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub FillG()

          Dim s As String
          Dim tb As Recordset
          Dim sql As String

10        g.Rows = 2
20        g.AddItem ""
30        g.RemoveItem 1

40        sql = "SELECT * FROM MicroExtLabName"
50        Set tb = New Recordset
60        RecOpenServer 0, tb, sql
70        Do While Not tb.EOF

80            s = tb!LabName & vbTab & _
                  tb!Address0 & vbTab & _
                  tb!Address1 & vbTab & _
                  tb!Address2 & ""
90            g.AddItem s

100           tb.MoveNext
110       Loop

120       If g.Rows > 2 Then
130           g.RemoveItem 1
140       End If

End Sub

Private Sub cmdadd_Click()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo cmdadd_Click_Error

20        If Trim$(txtLabName) = "" Then
30            iMsg "Enter Laboratory Name", vbCritical
40            Exit Sub
50        End If

60        sql = "SELECT * FROM MicroExtLabName WHERE " & _
                "LabName = '" & txtLabName & "'"
70        Set tb = New Recordset
80        RecOpenServer 0, tb, sql
90        If Not tb.EOF Then
100           iMsg "Lab Name already entered!", vbExclamation
110           txtLabName = ""
120           Exit Sub
130       Else
140           tb.AddNew
150           tb!LabName = txtLabName
160           tb!Address0 = txtAddress(0)
170           tb!Address1 = txtAddress(1)
180           tb!Address2 = txtAddress(2)
190           tb.Update
200       End If

210       txtLabName = ""
220       txtAddress(0) = ""
230       txtAddress(1) = ""
240       txtAddress(2) = ""

250       FillG

260       Exit Sub

cmdadd_Click_Error:

          Dim strES As String
          Dim intEL As Integer

270       intEL = Erl
280       strES = Err.Description
290       LogError "frmMicroExtRequestTo", "cmdAdd_Click", intEL, strES, sql

End Sub

Private Sub cmdCancel_Click()

10        Unload Me

End Sub


Private Sub cmdSave_Click()

End Sub


Private Sub Form_Load()

10        FillG

End Sub


