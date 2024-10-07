VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmPrinters 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Printers"
   ClientHeight    =   8475
   ClientLeft      =   630
   ClientTop       =   1095
   ClientWidth     =   10110
   Icon            =   "frmPrinters.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8475
   ScaleWidth      =   10110
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCheck 
      Caption         =   "Check Printers"
      Height          =   825
      Left            =   8010
      TabIndex        =   16
      Top             =   4095
      Width           =   780
   End
   Begin VB.CommandButton bprint 
      Appearance      =   0  'Flat
      Caption         =   "&Print"
      Height          =   825
      Left            =   8010
      Picture         =   "frmPrinters.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5085
      Width           =   795
   End
   Begin VB.CommandButton bcancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   825
      Left            =   8010
      Picture         =   "frmPrinters.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7065
      Width           =   795
   End
   Begin VB.Frame FrameAdd 
      Caption         =   "Add New Printer"
      Height          =   1515
      Left            =   150
      TabIndex        =   1
      Top             =   120
      Width           =   9735
      Begin VB.ListBox lAvailable 
         Height          =   1185
         IntegralHeight  =   0   'False
         Left            =   4620
         TabIndex        =   13
         Top             =   240
         Width           =   4965
      End
      Begin VB.TextBox tMappedTo 
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
         Left            =   180
         MaxLength       =   7
         TabIndex        =   4
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox tPrinterName 
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
         Left            =   180
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1050
         Width           =   3495
      End
      Begin VB.CommandButton bAdd 
         Caption         =   "Add"
         Height          =   780
         Left            =   2655
         Picture         =   "frmPrinters.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   210
         Width           =   960
      End
      Begin VB.Label lCopy 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Copy"
         Height          =   285
         Left            =   4140
         TabIndex        =   15
         Top             =   1140
         Width           =   480
      End
      Begin VB.Image iCopy 
         Height          =   480
         Left            =   3690
         Picture         =   "frmPrinters.frx":0C28
         Top             =   1020
         Width           =   480
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Available Printers"
         Height          =   195
         Left            =   4680
         TabIndex        =   14
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mapped To"
         Height          =   195
         Left            =   210
         TabIndex        =   6
         Top             =   270
         Width           =   825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Printer Name"
         Height          =   195
         Left            =   210
         TabIndex        =   5
         Top             =   870
         Width           =   915
      End
   End
   Begin VB.CommandButton bSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   825
      Left            =   8010
      Picture         =   "frmPrinters.frx":106A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6075
      Width           =   795
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   6525
      Left            =   150
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1770
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   11509
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      FormatString    =   "^Mapped To |<Printer Name                                                                    "
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
   Begin VB.Image Image1 
      Height          =   480
      Left            =   6870
      Picture         =   "frmPrinters.frx":1374
      Top             =   2910
      Width           =   480
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Click on Specific Printer Name to Edit"
      Height          =   375
      Left            =   7350
      TabIndex        =   12
      Top             =   2970
      Width           =   1545
   End
   Begin VB.Label lCurrent 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6960
      TabIndex        =   11
      Top             =   1980
      Width           =   2925
   End
   Begin VB.Label Label3 
      Caption         =   "Current Default Printer"
      Height          =   195
      Left            =   6990
      TabIndex        =   10
      Top             =   1770
      Width           =   1695
   End
End
Attribute VB_Name = "frmPrinters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Activated As Boolean
Private Sub CopyToName()

          Dim n As Long
          Dim Found As Boolean

10        On Error GoTo CopyToName_Error

20        For n = 0 To lAvailable.ListCount - 1
30            If lAvailable.Selected(n) Then
40                tPrinterName = lAvailable.List(n)
50                lAvailable.Selected(n) = False
60                Found = True
70                Exit For
80            End If
90        Next

100       If Not Found Then
110           MsgBox "Make a SELECTion from the available printers.", vbInformation
120       End If

130       Exit Sub

CopyToName_Error:

          Dim strES As String
          Dim intEL As Integer



140       intEL = Erl
150       strES = Err.Description
160       LogError "frmPrinters", "CopyToName", intEL, strES


End Sub

Private Sub bAdd_Click()

10        On Error GoTo bAdd_Click_Error

20        tMappedTo = Trim$(UCase$(tMappedTo))
30        tPrinterName = Trim$(UCase$(tPrinterName))

40        If tMappedTo = "" Then
50            Exit Sub
60        End If
70        If tPrinterName = "" Then
80            Exit Sub
90        End If

100       g.AddItem tMappedTo & vbTab & tPrinterName

110       tMappedTo = ""
120       tPrinterName = ""

130       bsave.Enabled = True

140       Exit Sub

bAdd_Click_Error:

          Dim strES As String
          Dim intEL As Integer



150       intEL = Erl
160       strES = Err.Description
170       LogError "frmPrinters", "bAdd_Click", intEL, strES


End Sub

Private Sub bcancel_Click()

10        Unload Me

End Sub


Private Sub bSave_Click()

          Dim Y As Long
          Dim sql As String
          Dim sn As New Recordset

10        On Error GoTo bSave_Click_Error

20        For Y = 1 To g.Rows - 1
30            sql = "SELECT * from printers WHERE mappedto = '" & UCase$(g.TextMatrix(Y, 0)) & "'"
40            Set sn = New Recordset
50            RecOpenServer 0, sn, sql
60            If sn.EOF Then sn.AddNew
70            sn!MappedTo = UCase$(g.TextMatrix(Y, 0))
80            sn!PrinterName = UCase$(g.TextMatrix(Y, 1))
90            sn.Update
100       Next

110       FillG

120       tMappedTo = ""
130       tPrinterName = ""
140       tMappedTo.SetFocus
150       bsave.Enabled = False

160       Exit Sub

bSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer



170       intEL = Erl
180       strES = Err.Description
190       LogError "frmPrinters", "bSave_Click", intEL, strES, sql


End Sub


Private Sub FillG()

          Dim s As String
          Dim sql As String
          Dim tb As New Recordset


10        On Error GoTo FillG_Error

20        ClearFGrid g

30        sql = "SELECT * from printers"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        Do While Not tb.EOF
70            s = Trim(tb!MappedTo) & vbTab & Trim(tb!PrinterName)
80            g.AddItem s
90            tb.MoveNext
100       Loop

110       FixG g

120       Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer



130       intEL = Erl
140       strES = Err.Description
150       LogError "frmPrinters", "FillG", intEL, strES, sql


End Sub




Private Sub cmdCheck_Click()
          Dim Px As Printer

10        On Error GoTo cmdCheck_Click_Error

20        For Each Px In Printers
30            Set Printer = Px
40            If PrinterInstalled = False Then
50                iMsg "problem with " & Px.DeviceName
60            End If
70        Next


80        Exit Sub

cmdCheck_Click_Error:

          Dim strES As String
          Dim intEL As Integer



90        intEL = Erl
100       strES = Err.Description
110       LogError "frmPrinters", "cmdCheck_Click", intEL, strES


End Sub

Private Sub Form_Activate()

10        On Error GoTo Form_Activate_Error

20        If Activated Then Exit Sub

30        FillG

40        Activated = True

50        Exit Sub

Form_Activate_Error:

          Dim strES As String
          Dim intEL As Integer



60        intEL = Erl
70        strES = Err.Description
80        LogError "frmPrinters", "Form_Activate", intEL, strES


End Sub

Private Sub Form_Load()

          Dim Px As Printer

10        On Error GoTo Form_Load_Error

20        lCurrent = Printer.DeviceName

30        lAvailable.Clear
40        For Each Px In Printers
50            lAvailable.AddItem Px.DeviceName
60        Next

70        Activated = False

80        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer



90        intEL = Erl
100       strES = Err.Description
110       LogError "frmPrinters", "Form_Load", intEL, strES


End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

10        On Error GoTo Form_QueryUnload_Error

20        If bsave.Enabled Then
30            If MsgBox("Cancel without saving?", vbQuestion + vbYesNo) = vbNo Then
40                Cancel = True
50            End If
60        End If

70        Exit Sub

Form_QueryUnload_Error:

          Dim strES As String
          Dim intEL As Integer



80        intEL = Erl
90        strES = Err.Description
100       LogError "frmPrinters", "Form_QueryUnload", intEL, strES


End Sub

Private Sub g_Click()

          Dim OldName As String
          Dim NewName As String

10        On Error GoTo g_Click_Error

20        If g.MouseRow = 0 Then Exit Sub
30        If g.MouseCol = 0 Then Exit Sub

40        OldName = g.TextMatrix(g.Row, 1)
50        NewName = iBOX("PROCEED WITH CAUTION" & vbCrLf & vbCrLf & "New Printer Name?", , OldName)
60        If Trim$(NewName) = "" Then
70            Exit Sub
80        End If

90        If MsgBox("Change " & vbCrLf & OldName & vbCrLf & "to" & vbCrLf & NewName, vbQuestion + vbYesNo) = vbNo Then Exit Sub

100       g.TextMatrix(g.Row, 1) = NewName
110       bsave.Enabled = True

120       Exit Sub

g_Click_Error:

          Dim strES As String
          Dim intEL As Integer



130       intEL = Erl
140       strES = Err.Description
150       LogError "frmPrinters", "g_Click", intEL, strES


End Sub



Private Sub iCopy_Click()

10        On Error GoTo iCopy_Click_Error

20        CopyToName

30        Exit Sub

iCopy_Click_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmPrinters", "iCopy_Click", intEL, strES


End Sub
Private Sub lCopy_Click()

10        On Error GoTo lCopy_Click_Error

20        CopyToName

30        Exit Sub

lCopy_Click_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmPrinters", "lCopy_Click", intEL, strES


End Sub



Public Function PrinterInstalled() As Boolean
10        On Error Resume Next

          Dim strDummy As String
20        strDummy = Printer.DeviceName

30        If Err.Number Then
40            PrinterInstalled = False
50        Else
60            PrinterInstalled = True
70        End If

End Function
