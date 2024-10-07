VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmAddCoagTest 
   Caption         =   "NetAcquire - Add Coagulation Test"
   ClientHeight    =   3555
   ClientLeft      =   2400
   ClientTop       =   3615
   ClientWidth     =   7290
   Icon            =   "frmAddCoagTest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   7290
   Begin VB.Frame Frame1 
      Height          =   2265
      Left            =   45
      TabIndex        =   8
      Top             =   45
      Width           =   4425
      Begin VB.ComboBox cmbUnits 
         Height          =   315
         Left            =   1170
         TabIndex        =   2
         Text            =   "cunits"
         Top             =   1275
         Width           =   1965
      End
      Begin VB.TextBox txtCode 
         Height          =   285
         Left            =   1170
         MaxLength       =   4
         TabIndex        =   0
         Top             =   225
         Width           =   825
      End
      Begin VB.TextBox txtTestName 
         Height          =   285
         Left            =   1155
         MaxLength       =   40
         TabIndex        =   1
         Top             =   765
         Width           =   3105
      End
      Begin VB.CommandButton cmdAddUnit 
         Caption         =   "Add New Unit"
         Height          =   855
         Left            =   3195
         Picture         =   "frmAddCoagTest.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1125
         Width           =   1125
      End
      Begin VB.TextBox txtAnalyser 
         Height          =   285
         Left            =   1170
         MaxLength       =   4
         TabIndex        =   4
         Top             =   1770
         Width           =   825
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Units"
         Height          =   195
         Left            =   645
         TabIndex        =   12
         Top             =   1335
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Code"
         Height          =   195
         Left            =   630
         TabIndex        =   11
         Top             =   255
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Test Name"
         Height          =   195
         Left            =   165
         TabIndex        =   10
         Top             =   795
         Width           =   780
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Analyser Code"
         Height          =   195
         Left            =   135
         TabIndex        =   9
         Top             =   1800
         Width           =   1470
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdCoag 
      Height          =   3390
      Left            =   4545
      TabIndex        =   7
      Top             =   45
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   5980
      _Version        =   393216
      Cols            =   3
      FormatString    =   "Test Name    | Code   |Units   "
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   885
      Left            =   135
      Picture         =   "frmAddCoagTest.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2430
      Width           =   1560
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   885
      Left            =   2295
      Picture         =   "frmAddCoagTest.frx":0A56
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2430
      Width           =   1470
   End
End
Attribute VB_Name = "frmAddCoagTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAddUnit_Click()

10        On Error GoTo cmdAddUnit_Click_Error

20        With frmLists
30            .o(5) = True
40            .Show 1
50        End With
60        FillLists

70        Exit Sub

cmdAddUnit_Click_Error:

          Dim strES As String
          Dim intEL As Integer

80        intEL = Erl
90        strES = Err.Description
100       LogError "frmAddCoagTest", "cmdAddUnit_Click", intEL, strES


End Sub

Private Sub cmdCancel_Click()

10        Unload Me

End Sub

Private Sub cmdSave_Click()

          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo cmdSave_Click_Error

20        If Trim$(txtCode) = "" Then
30            iMsg "Enter Code.", vbCritical
40            Exit Sub
50        End If

60        If Trim$(txtTestName) = "" Then
70            iMsg "Enter Test Name.", vbCritical
80            Exit Sub
90        End If

100       If Len(cmbUnits) = 0 Then
110           iMsg "SELECT Units.", vbCritical
120           Exit Sub
130       End If

140       sql = "SELECT * from coagtestdefinitions WHERE code = '" & txtCode & "' " & _
                "and agefromdays = '0' and AGETODAYS = '" & MaxAgeToDays & "'"
150       Set tb = New Recordset
160       RecOpenServer 0, tb, sql
170       If Not tb.EOF Then
180           iMsg "Code already used.", vbCritical
190           Exit Sub
200       Else
210           tb.AddNew
220           tb!Code = UCase(txtCode)
230           tb!TestName = txtTestName
240           tb!DoDelta = False
250           tb!DeltaLimit = 0
260           tb!PrintPriority = 999
270           tb!DP = 1
280           tb!Units = cmbUnits
290           tb!MaleLow = 0
300           tb!MaleHigh = 9999
310           tb!FemaleLow = 0
320           tb!FemaleHigh = 9999
330           tb!FlagMaleLow = 0
340           tb!FlagMaleHigh = 9999
350           tb!FlagFemaleLow = 0
360           tb!FlagFemaleHigh = 9999
370           tb!Category = ""
380           tb!Printable = True
390           tb!PlausibleLow = 0
400           tb!PlausibleHigh = 9999
410           tb!InUse = True
420           tb!AgeFromDays = 0
430           tb!AgeToDays = 43820
440           tb!Hospital = HospName(0)
450           tb.Update
460       End If

470       sql = "Select * from Coagtranslation"
480       Set tb = New Recordset
490       RecOpenServer 0, tb, sql
500       tb.AddNew
510       tb!trancode = txtAnalyser
520       tb!Code = txtCode
530       tb!Units = cmbUnits
540       tb.Update

550       Unload Me

560       Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

570       intEL = Erl
580       strES = Err.Description
590       LogError "frmAddCoagTest", "cmdsave_Click", intEL, strES, sql

End Sub

Private Sub FillG()


          Dim sql As String
          Dim tb As New Recordset


10        On Error GoTo FillG_Error

20        ClearFGrid grdCoag

30        sql = "SELECT * from coagtestdefinitions"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        Do While Not tb.EOF
70            grdCoag.AddItem Trim(tb!TestName) & vbTab & Trim(tb!Code) & vbTab & Trim(tb!Units)
80            tb.MoveNext
90        Loop

100       FixG grdCoag




110       Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

120       intEL = Erl
130       strES = Err.Description
140       LogError "frmAddCoagTest", "FillG", intEL, strES, sql


End Sub

Private Sub FillLists()

          Dim tb As New Recordset
          Dim sql As String



10        On Error GoTo FillLists_Error

20        cmbUnits.Clear

30        sql = "SELECT * from lists WHERE listtype = 'UN'"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql

60        Do While Not tb.EOF
70            cmbUnits.AddItem Trim(tb!Text & "")
80            tb.MoveNext
90        Loop




100       Exit Sub

FillLists_Error:

          Dim strES As String
          Dim intEL As Integer

110       intEL = Erl
120       strES = Err.Description
130       LogError "frmAddCoagTest", "FillLists", intEL, strES


End Sub

Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        FillLists

30        FillG

40        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmAddCoagTest", "Form_Load", intEL, strES


End Sub

Private Sub txtCode_LostFocus()

10        On Error GoTo txtCode_LostFocus_Error

20        If txtCode <> "" Then
30            If Not IsNumeric(txtCode) Then
40                iMsg "Code must be Numeric"
50                txtCode = ""
60            End If
70        End If

80        Exit Sub

txtCode_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmAddCoagTest", "txtCode_LostFocus", intEL, strES


End Sub
