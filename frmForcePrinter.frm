VERSION 5.00
Begin VB.Form frmForcePrinter 
   Caption         =   "NetAcquire - Printer Options"
   ClientHeight    =   4425
   ClientLeft      =   2085
   ClientTop       =   2415
   ClientWidth     =   6195
   Icon            =   "frmForcePrinter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   6195
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   825
      Left            =   810
      Picture         =   "frmForcePrinter.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3300
      Width           =   1275
   End
   Begin VB.CommandButton bcancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   825
      Left            =   4110
      Picture         =   "frmForcePrinter.frx":1C8C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3300
      Width           =   1275
   End
   Begin VB.ListBox lAvailable 
      Height          =   2325
      IntegralHeight  =   0   'False
      Left            =   600
      TabIndex        =   0
      Top             =   750
      Width           =   4965
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmForcePrinter.frx":360E
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   150
      Width           =   4755
   End
End
Attribute VB_Name = "frmForcePrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public f As Form
Private Sub SaveSelection()

10        On Error GoTo SaveSelection_Error

20        If f.Name = "frmEditHisto" Then
30            SaveSetting "NetAcquire", "Histology", "Printer", lAvailable
40        End If

50        f.PrintToPrinter = lAvailable

60        Unload Me

70        Exit Sub

SaveSelection_Error:

          Dim strES As String
          Dim intEL As Integer

80        intEL = Erl
90        strES = Err.Description
100       LogError "frmForcePrinter", "SaveSelection", intEL, strES

End Sub

Private Sub bcancel_Click()

10        Unload Me

End Sub

Private Sub cmdSave_Click()

10        SaveSelection

End Sub

Private Sub Form_Load()

          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo Form_Load_Error

20        lAvailable.Clear
30        lAvailable.AddItem "Automatic Selection"
40        lAvailable.AddItem ""

50        sql = "SELECT * FROM InstalledPrinters WHERE " & _
                "Location = 'Lab'"
60        Set tb = New Recordset
70        RecOpenServer 0, tb, sql

80        Do While Not tb.EOF
90            lAvailable.AddItem tb!PrinterName & ""
100           tb.MoveNext
110       Loop

120       Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer



130       intEL = Erl
140       strES = Err.Description
150       LogError "frmForcePrinter", "Form_Load", intEL, strES, sql


End Sub


Private Sub lAvailable_DblClick()

10        SaveSelection

End Sub


