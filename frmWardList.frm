VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWardList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Ward List"
   ClientHeight    =   7140
   ClientLeft      =   375
   ClientTop       =   705
   ClientWidth     =   14415
   Icon            =   "frmWardList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   14415
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraGPPrinting 
      Caption         =   "DR. Name"
      Height          =   3150
      Left            =   11160
      TabIndex        =   18
      Top             =   1575
      Visible         =   0   'False
      Width           =   3180
      Begin VB.ListBox LstGpPrinting 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2220
         ItemData        =   "frmWardList.frx":030A
         Left            =   135
         List            =   "frmWardList.frx":0326
         Style           =   1  'Checkbox
         TabIndex        =   21
         Top             =   270
         Width           =   2910
      End
      Begin VB.CommandButton cmdSaveGpPrinting 
         Appearance      =   0  'Flat
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   390
         TabIndex        =   20
         Top             =   2610
         Width           =   930
      End
      Begin VB.CommandButton cmdExitGpPrinting 
         Appearance      =   0  'Flat
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1845
         TabIndex        =   19
         Top             =   2610
         Width           =   930
      End
   End
   Begin VB.CommandButton cmdMoveDown 
      Caption         =   "Move Down"
      Enabled         =   0   'False
      Height          =   825
      Left            =   11160
      Picture         =   "frmWardList.frx":0396
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6045
      Width           =   795
   End
   Begin VB.CommandButton cmdMoveUp 
      Caption         =   "Move Up"
      Enabled         =   0   'False
      Height          =   825
      Left            =   11160
      Picture         =   "frmWardList.frx":07D8
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5220
      Width           =   795
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   225
      Left            =   12975
      TabIndex        =   16
      Top             =   45
      Visible         =   0   'False
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H0000FFFF&
      Caption         =   "Save"
      Height          =   705
      Left            =   12975
      Picture         =   "frmWardList.frx":0C1A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   270
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.ListBox lHospital 
      Height          =   1230
      Left            =   7110
      TabIndex        =   14
      Top             =   150
      Width           =   1995
   End
   Begin VB.Frame Frame2 
      Caption         =   "Add New Ward"
      Height          =   1365
      Left            =   120
      TabIndex        =   11
      Top             =   60
      Width           =   5925
      Begin VB.TextBox txtPrinter 
         Height          =   285
         Left            =   810
         TabIndex        =   3
         Top             =   960
         Width           =   4755
      End
      Begin VB.TextBox txtFAX 
         Height          =   285
         Left            =   2220
         MaxLength       =   20
         TabIndex        =   1
         Top             =   240
         Width           =   2235
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   765
         Left            =   4725
         Picture         =   "frmWardList.frx":0F24
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   150
         Width           =   930
      End
      Begin VB.TextBox txtText 
         Height          =   285
         Left            =   810
         MaxLength       =   50
         TabIndex        =   2
         Top             =   600
         Width           =   3645
      End
      Begin VB.TextBox txtCode 
         Height          =   285
         Left            =   810
         MaxLength       =   12
         TabIndex        =   0
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Printer"
         Height          =   195
         Left            =   330
         TabIndex        =   17
         Top             =   990
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "FAX"
         Height          =   195
         Left            =   1860
         TabIndex        =   15
         Top             =   270
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ward"
         Height          =   195
         Left            =   390
         TabIndex        =   13
         Top             =   630
         Width           =   390
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Code"
         Height          =   195
         Left            =   420
         TabIndex        =   12
         Top             =   270
         Width           =   375
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdWardList 
      Height          =   5445
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   9604
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      AllowUserResizing=   1
      FormatString    =   $"frmWardList.frx":122E
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   705
      Left            =   13050
      Picture         =   "frmWardList.frx":12F1
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6345
      Width           =   1290
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   735
      Left            =   13050
      MaskColor       =   &H8000000F&
      Picture         =   "frmWardList.frx":15FB
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5265
      Width           =   1290
   End
End
Attribute VB_Name = "frmWardList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

10    Unload Me

End Sub


Private Sub cmdExitGpPrinting_Click()
10    fraGPPrinting.Visible = False
End Sub

Private Sub cmdMoveDown_Click()

      Dim n As Long
      Dim s As String
      Dim X As Long

10    On Error GoTo cmdMoveDown_Click_Error

20    If grdWardList.row = grdWardList.Rows - 1 Then Exit Sub
30    n = grdWardList.row

40    s = ""
50    For X = 0 To grdWardList.Cols - 1
60        s = s & grdWardList.TextMatrix(n, X) & vbTab
70    Next
80    s = Left$(s, Len(s) - 1)

90    grdWardList.RemoveItem n
100   If n < grdWardList.Rows Then
110       grdWardList.AddItem s, n + 1
120       grdWardList.row = n + 1
130   Else
140       grdWardList.AddItem s
150       grdWardList.row = grdWardList.Rows - 1
160   End If

170   For X = 0 To grdWardList.Cols - 1
180       grdWardList.Col = X
190       grdWardList.CellBackColor = vbYellow
200   Next

210   cmdSave.Visible = True

220   Exit Sub

cmdMoveDown_Click_Error:

      Dim strES As String
      Dim intEL As Integer


230   intEL = Erl
240   strES = Err.Description
250   LogError "frmWardList", "cmdMoveDown_Click", intEL, strES


End Sub

Private Sub cmdMoveUp_Click()

      Dim n As Long
      Dim s As String
      Dim X As Long

10    On Error GoTo cmdMoveUp_Click_Error

20    If grdWardList.row = 1 Then Exit Sub

30    n = grdWardList.row

40    s = ""
50    For X = 0 To grdWardList.Cols - 1
60        s = s & grdWardList.TextMatrix(n, X) & vbTab
70    Next
80    s = Left$(s, Len(s) - 1)

90    grdWardList.RemoveItem n
100   grdWardList.AddItem s, n - 1

110   grdWardList.row = n - 1
120   For X = 0 To grdWardList.Cols - 1
130       grdWardList.Col = X
140       grdWardList.CellBackColor = vbYellow
150   Next

160   cmdSave.Visible = True

170   Exit Sub

cmdMoveUp_Click_Error:

      Dim strES As String
      Dim intEL As Integer


180   intEL = Erl
190   strES = Err.Description
200   LogError "frmWardList", "cmdMoveUp_Click", intEL, strES


End Sub


Private Sub cmdPrint_Click()

      Dim Y As Long

10    On Error GoTo cmdPrint_Click_Error

20    Printer.Print
30    Printer.Font.Name = "Courier New"
40    Printer.Font.Size = 12

50    Printer.Print "List of Wards. ("; lHospital; ")"
60    Printer.Print

70    With grdWardList
80        For Y = 0 To .Rows - 1
90            Printer.Print .TextMatrix(Y, 0); Tab(7);    'InUse
100           Printer.Print .TextMatrix(Y, 1); Tab(15);    'Code
110           Printer.Print .TextMatrix(Y, 2)    'Text
120       Next
130   End With

140   Printer.EndDoc

150   Exit Sub

cmdPrint_Click_Error:

      Dim strES As String
      Dim intEL As Integer

160   intEL = Erl
170   strES = Err.Description
180   LogError "frmWardList", "cmdPrint_Click", intEL, strES

End Sub

Private Sub cmdadd_Click()

10    On Error GoTo cmdadd_Click_Error

20    txtCode = UCase$(Trim$(txtCode))
30    If txtCode = "" Then
40        iMsg "Enter Code.", vbCritical
50        Exit Sub
60    End If


70    If Len(txtCode) <> 3 Then
80        iMsg "Code Length Must be 3 Chars.", vbCritical
90        Exit Sub
100   End If


110   txtText = Trim$(txtText)
120   If txtText = "" Then
130       iMsg "Enter Ward.", vbCritical
140       Exit Sub
150   End If

160   grdWardList.AddItem "Yes" & vbTab & _
                          txtCode & vbTab & _
                          txtText & vbTab & _
                          txtFAX & vbTab & _
                          txtPrinter

170   txtCode = ""
180   txtText = ""
190   txtFAX = ""
200   txtPrinter = ""
210   txtCode.Locked = False

220   FixG grdWardList

230   cmdSave.Visible = True

240   Exit Sub

cmdadd_Click_Error:

      Dim strES As String
      Dim intEL As Integer


250   intEL = Erl
260   strES = Err.Description
270   LogError "frmWardList", "cmdAdd_Click", intEL, strES


End Sub

Private Sub cmdSave_Click()
      Dim Hosp As String
      Dim Y As Long
      Dim sql As String
      Dim tb As New Recordset

10    On Error GoTo cmdSave_Click_Error

20    Hosp = ListCodeFor("HO", lHospital)

30    pb.Max = grdWardList.Rows - 1
40    pb.Visible = True
50    cmdSave.Caption = "Saving..."

60    For Y = 1 To grdWardList.Rows - 1
70        pb = Y
80        sql = "SELECT * from wards WHERE hospitalcode = '" & Hosp & "' and " & _
                "code = '" & grdWardList.TextMatrix(Y, 1) & "'"
90        Set tb = New Recordset
100       RecOpenServer 0, tb, sql
110       If tb.EOF Then tb.AddNew
120       tb!Code = grdWardList.TextMatrix(Y, 1)
130       tb!HospitalCode = Hosp
140       If grdWardList.TextMatrix(Y, 0) = "Yes" Then tb!InUse = 1 Else tb!InUse = 0
150       tb!Text = grdWardList.TextMatrix(Y, 2)
160       tb!FAX = grdWardList.TextMatrix(Y, 3)
170       tb!PrinterAddress = grdWardList.TextMatrix(Y, 4)
180       tb!ListOrder = Y
190       tb.Update
200   Next




210   pb.Visible = False
220   cmdSave.Visible = False
230   cmdSave.Caption = "Save"

240   Exit Sub

cmdSave_Click_Error:

      Dim strES As String
      Dim intEL As Integer


250   intEL = Erl
260   strES = Err.Description
270   LogError "frmWardList", "cmdsave_Click", intEL, strES, sql


End Sub


Private Sub cmdSaveGpPrinting_Click()
      Dim Y As Long
      Dim sql As String
      Dim tb As Recordset
      Dim Dept As String

10    On Error GoTo cmdSaveGpPrinting_Click_Error

20    For Y = 0 To LstGpPrinting.ListCount - 1
30        Dept = LstGpPrinting.List(Y)

40        If LstGpPrinting.Selected(Y) = True Then

50            sql = "IF EXISTS (SELECT * FROM DisablePrinting WHERE " & _
                    "           Department = '" & LstGpPrinting.List(Y) & "' " & _
                    "           AND GPCode = '" & grdWardList.TextMatrix(grdWardList.row, 1) & "' )" & _
                    "    UPDATE DisablePrinting " & _
                    "    SET Department = '" & LstGpPrinting.List(Y) & "', " & _
                    "    GPCode = '" & grdWardList.TextMatrix(grdWardList.row, 1) & "', " & _
                    "    GPName = '" & grdWardList.TextMatrix(grdWardList.row, 2) & "', " & _
                    "    Type = 'WARD', " & _
                    "    Disable = '1' " & _
                    "    WHERE Department = '" & LstGpPrinting.List(Y) & "' " & _
                    "    AND GPCode = '" & grdWardList.TextMatrix(grdWardList.row, 1) & "' " & _
                    "ELSE " & _
                    "    INSERT INTO DisablePrinting " & _
                    "    (GPCode, GPName, Type, Disable, Department) "
60            sql = sql & _
                    "    VALUES ( " & _
                    "    '" & grdWardList.TextMatrix(grdWardList.row, 1) & "', " & _
                    "    '" & grdWardList.TextMatrix(grdWardList.row, 2) & "', " & _
                    "    'WARD', " & _
                    "    '1', " & _
                    "    '" & LstGpPrinting.List(Y) & "'" & _
                    "     )"
70            Cnxn(0).Execute sql
80        Else
90            sql = "DELETE FROM DisablePrinting WHERE " & _
                    "Department = '" & LstGpPrinting.List(Y) & "' " & _
                    "AND GPCode = '" & grdWardList.TextMatrix(grdWardList.row, 0) & "'"
100           Cnxn(0).Execute sql
110       End If
120   Next

130   fraGPPrinting.Visible = False
      'MsgBox "Updated!", vbOKOnly, "Disable Printing"

140   Exit Sub

cmdSaveGpPrinting_Click_Error:
      Dim strES As String
      Dim intEL As Integer

150   intEL = Erl
160   strES = Err.Description
170   LogError "frmWardList", "cmdSaveGpPrinting_Click", intEL, strES, sql

End Sub

Private Sub Form_Load()

      Dim tb As New Recordset
      Dim sql As String

10    On Error GoTo Form_Load_Error

20    lHospital.Clear

30    sql = "SELECT * from lists WHERE listtype = 'HO'"
40    Set tb = New Recordset
50    RecOpenServer 0, tb, sql

60    Do While Not tb.EOF
70        lHospital.AddItem Trim(tb!Text)
80        tb.MoveNext
90    Loop

100   If lHospital.ListCount > 0 Then
110       lHospital.ListIndex = 0
120   End If
130   grdWardList.ColWidth(5) = 0

140   If GetOptionSetting("DisableWardPrinting", 0) = True Then
150       grdWardList.ColWidth(5) = 1200
160   End If
170   FillG

180   Set_Font Me

190   Exit Sub

Form_Load_Error:

      Dim strES As String
      Dim intEL As Integer


200   intEL = Erl
210   strES = Err.Description
220   LogError "frmWardList", "Form_Load", intEL, strES, sql


End Sub

Private Sub FillG()

      Dim s As String
      Dim Hosp As String
      Dim tb As New Recordset
      Dim sql As String


10    On Error GoTo FillG_Error

20    Hosp = ListCodeFor("HO", lHospital)

30    ClearFGrid grdWardList

40    sql = "SELECT * from wards WHERE hospitalcode = '" & Hosp & "' order by listorder"
50    Set tb = New Recordset
60    RecOpenServer 0, tb, sql
70    Do While Not tb.EOF
80        If Val(tb!InUse) = 1 Then s = "Yes" Else s = "No"
90        s = s & vbTab & _
              Trim(tb!Code) & vbTab & _
              Trim(tb!Text) & vbTab & _
              Trim(tb!FAX) & vbTab & _
              Trim(tb!PrinterAddress) & vbTab & _
              "Click To View"
100       grdWardList.AddItem s
110       tb.MoveNext
120   Loop

130   FixG grdWardList



140   Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer


150   intEL = Erl
160   strES = Err.Description
170   LogError "frmWardList", "FillG", intEL, strES, sql


End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

10    On Error GoTo Form_QueryUnload_Error

20    If cmdSave.Visible Then
30        If iMsg("Cancel without saving?", vbQuestion + vbYesNo) = vbNo Then
40            Cancel = True
50        End If
60    End If

70    Exit Sub

Form_QueryUnload_Error:

      Dim strES As String
      Dim intEL As Integer


80    intEL = Erl
90    strES = Err.Description
100   LogError "frmWardList", "Form_QueryUnload", intEL, strES


End Sub

Private Sub grdWardList_Click()

      Static SortOrder As Boolean
      Dim X As Long
      Dim Y As Long
      Dim ySave As Long

10    On Error GoTo grdWardList_Click_Error
      'Debug.Print grdWardList.Col
20    ySave = grdWardList.row

30    If grdWardList.MouseRow = 0 Then
40        If SortOrder Then
50            grdWardList.Sort = flexSortGenericAscending
60        Else
70            grdWardList.Sort = flexSortGenericDescending
80        End If
90        SortOrder = Not SortOrder
100       cmdMoveUp.Enabled = False
110       cmdMoveDown.Enabled = False
120       cmdSave.Visible = True
130       Exit Sub
140   End If

150   If grdWardList.Col = 0 Then
160       grdWardList = IIf(grdWardList = "No", "Yes", "No")
170       cmdSave.Visible = True
180       Exit Sub
190   End If


200   If grdWardList.Col = 3 Then
210       grdWardList = iBOX("Enter Fax Number ", , grdWardList, False)
220       cmdSave.Visible = True
230       Exit Sub
240   End If
250   If grdWardList.Col = 1 Then
260       grdWardList.Enabled = False
270       If iMsg("Edit this line?", vbQuestion + vbYesNo) = vbYes Then
280           txtCode = grdWardList.TextMatrix(grdWardList.row, 1)
290           txtCode.Locked = True
300           txtText = grdWardList.TextMatrix(grdWardList.row, 2)
310           txtFAX = grdWardList.TextMatrix(grdWardList.row, 3)
320           txtPrinter = grdWardList.TextMatrix(grdWardList.row, 4)
330           grdWardList.RemoveItem grdWardList.row
340           cmdSave.Visible = True
350       End If
360       grdWardList.Enabled = True
370       Exit Sub
380   End If
390   If grdWardList.Col = 5 Then
400       fraGPPrinting.Visible = True
410       fraGPPrinting.Caption = grdWardList.TextMatrix(grdWardList.row, 2)
420       FillLstGpPrinting
430   End If
440   grdWardList.Visible = False
450   grdWardList.Col = 0
460   For Y = 1 To grdWardList.Rows - 1
470       grdWardList.row = Y
480       If grdWardList.CellBackColor = vbYellow Then
490           For X = 0 To grdWardList.Cols - 1
500               grdWardList.Col = X
510               grdWardList.CellBackColor = 0
520           Next
530           Exit For
540       End If
550   Next

560   grdWardList.row = ySave
570   grdWardList.Visible = True

580   For X = 0 To grdWardList.Cols - 1
590       grdWardList.Col = X
600       grdWardList.CellBackColor = vbYellow
610   Next

620   cmdMoveUp.Enabled = True
630   cmdMoveDown.Enabled = True

640   Exit Sub

grdWardList_Click_Error:

      Dim strES As String
      Dim intEL As Integer


650   intEL = Erl
660   strES = Err.Description
670   LogError "frmWardList", "grdWardList_Click", intEL, strES


End Sub
Private Sub FillLstGpPrinting()

      Dim tb As Recordset
      Dim sql As String
      'Dim dept As String
      Dim Y As Integer

10    On Error GoTo FillLstGpPrinting_Error

20    sql = "SELECT * FROM DisablePrinting WHERE " & _
            " GPCode = '" & grdWardList.TextMatrix(grdWardList.row, 1) & "'"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql

50    For Y = 0 To LstGpPrinting.ListCount - 1
          'dept = LstGpPrinting.List(Y)
60        LstGpPrinting.Selected(Y) = False
70    Next Y

80    Do While Not tb.EOF
90        For Y = 0 To LstGpPrinting.ListCount - 1
100           If LstGpPrinting.List(Y) = tb!Department Then
110               LstGpPrinting.Selected(Y) = True
120           End If
130       Next
140       tb.MoveNext
150   Loop


160   Exit Sub

FillLstGpPrinting_Error:
      Dim strES As String
      Dim intEL As Integer

170   intEL = Erl
180   strES = Err.Description
190   LogError "frmWardList", "FillLstGpPrinting", intEL, strES, sql

End Sub

Private Sub grdWardList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

10    On Error GoTo grdWardList_MouseMove_Error

20    If grdWardList.MouseRow = 0 Then
30        grdWardList.ToolTipText = ""
40    ElseIf grdWardList.MouseCol = 0 Then
50        grdWardList.ToolTipText = "Click to Toggle 1 for Inuse, 0 for not Inuse"
60    ElseIf grdWardList.MouseCol = 1 Then
70        grdWardList.ToolTipText = "Click to Edit"
80    Else
90        grdWardList.ToolTipText = "Click to Move"
100   End If

110   Exit Sub

grdWardList_MouseMove_Error:

      Dim strES As String
      Dim intEL As Integer


120   intEL = Erl
130   strES = Err.Description
140   LogError "frmWardList", "grdWardList_MouseMove", intEL, strES


End Sub


Private Sub lHospital_Click()

10    On Error GoTo lHospital_Click_Error

20    FillG

30    Exit Sub

lHospital_Click_Error:

      Dim strES As String
      Dim intEL As Integer


40    intEL = Erl
50    strES = Err.Description
60    LogError "frmWardList", "lHospital_Click", intEL, strES


End Sub


Private Sub txtCode_LostFocus()

10    On Error GoTo txtCode_LostFocus_Error

20    txtCode = UCase(txtCode)

30    Exit Sub

txtCode_LostFocus_Error:

      Dim strES As String
      Dim intEL As Integer


40    intEL = Erl
50    strES = Err.Description
60    LogError "frmWardList", "txtCode_LostFocus", intEL, strES


End Sub

