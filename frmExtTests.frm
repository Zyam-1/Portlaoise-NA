VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form frmExtTests 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - External Test List"
   ClientHeight    =   9795
   ClientLeft      =   240
   ClientTop       =   1455
   ClientWidth     =   15495
   ForeColor       =   &H80000008&
   Icon            =   "frmExtTests.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9795
   ScaleWidth      =   15495
   Begin VB.CommandButton cmdExcel 
      Caption         =   "Export to Excel"
      Height          =   1155
      Left            =   14040
      Picture         =   "frmExtTests.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save Changes"
      Height          =   1155
      Left            =   14040
      Picture         =   "frmExtTests.frx":5F28
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   7080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddToAddress 
      Appearance      =   0  'Flat
      Caption         =   "&Add Address"
      Height          =   1155
      Left            =   14040
      Picture         =   "frmExtTests.frx":6DF2
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   300
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Add Test"
      Height          =   1455
      Left            =   180
      TabIndex        =   4
      Top             =   120
      Width           =   12135
      Begin VB.TextBox txtSampleType 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7080
         TabIndex        =   22
         Top             =   945
         Width           =   2055
      End
      Begin VB.TextBox txtComment 
         Height          =   315
         Left            =   7080
         TabIndex        =   20
         Top             =   420
         Width           =   3675
      End
      Begin VB.ComboBox cmbDepartment 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   17
         Text            =   "Combo1"
         Top             =   405
         Width           =   2595
      End
      Begin VB.TextBox txtCode 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2850
         TabIndex        =   16
         Top             =   405
         Width           =   1125
      End
      Begin VB.TextBox txtNormalRange 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4080
         TabIndex        =   9
         Top             =   945
         Width           =   2895
      End
      Begin VB.TextBox txtUnits 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2850
         MaxLength       =   20
         TabIndex        =   8
         Top             =   945
         Width           =   1125
      End
      Begin VB.TextBox txtTestName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4080
         MaxLength       =   40
         TabIndex        =   7
         Top             =   405
         Width           =   2895
      End
      Begin VB.CommandButton cmdAddToList 
         Appearance      =   0  'Flat
         Caption         =   "Add to &List"
         Height          =   855
         Left            =   10920
         Picture         =   "frmExtTests.frx":70FC
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   420
         Width           =   1065
      End
      Begin VB.ComboBox cmbSendTo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   5
         Text            =   "cmbSendTo"
         Top             =   960
         Width           =   2595
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "SampleType"
         Height          =   195
         Left            =   7590
         TabIndex        =   23
         Top             =   750
         Width           =   885
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Comments"
         Height          =   195
         Left            =   8700
         TabIndex        =   19
         Top             =   210
         Width           =   735
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Department"
         Height          =   195
         Left            =   1020
         TabIndex        =   18
         Top             =   210
         Width           =   825
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Test Code"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3060
         TabIndex        =   14
         Top             =   210
         Width           =   735
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Normal Range"
         Height          =   195
         Left            =   5010
         TabIndex        =   13
         Top             =   750
         Width           =   1020
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Units"
         Height          =   195
         Left            =   3210
         TabIndex        =   12
         Top             =   750
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Test Name"
         Height          =   195
         Left            =   5130
         TabIndex        =   11
         Top             =   210
         Width           =   780
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Send to"
         Height          =   195
         Left            =   1110
         TabIndex        =   10
         Top             =   750
         Width           =   555
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   7485
      Left            =   180
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2160
      Width           =   13725
      _ExtentX        =   24209
      _ExtentY        =   13203
      _Version        =   393216
      Cols            =   8
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
      AllowUserResizing=   1
      FormatString    =   $"frmExtTests.frx":7406
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
   Begin VB.CommandButton cmdDelete 
      Appearance      =   0  'Flat
      Caption         =   "&Delete"
      Height          =   1155
      Left            =   14040
      Picture         =   "frmExtTests.frx":74F5
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5100
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Appearance      =   0  'Flat
      Caption         =   "&Print"
      Height          =   1155
      Left            =   14040
      Picture         =   "frmExtTests.frx":77FF
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   1155
      Left            =   14040
      Picture         =   "frmExtTests.frx":7B09
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8430
      Width           =   1215
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
      Left            =   14040
      TabIndex        =   26
      Top             =   3360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblABC 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   480
      Left            =   180
      TabIndex        =   24
      Top             =   1680
      Width           =   480
   End
End
Attribute VB_Name = "frmExtTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pDepartment As String

Private Sub cmbDepartment_Click()

10        pDepartment = cmbDepartment

20        FillG

End Sub

Private Sub cmbDepartment_KeyPress(KeyAscii As Integer)
10        KeyAscii = 0
End Sub


Private Sub cmdaddtoaddress_Click()

10        frmAddress.Show 1
20        FillAddress

End Sub

Private Sub cmdAddToList_Click()

          Dim sql As String
          Dim tb As New Recordset

10        On Error GoTo cmdAddToList_Click_Error

20        txtTestName = Trim$(txtTestName)
30        txtCode = Trim$(txtCode)
40        cmbSendTo = Trim$(cmbSendTo)

50        If txtTestName = "" Then
60            iMsg "Enter Test Name", vbCritical
70            Exit Sub
80        End If

90        If txtCode = "" Then
100           iMsg "Enter Test Code", vbCritical
110           Exit Sub
120       End If

130       If cmbSendTo = "" Then
140           iMsg "Enter 'Send To' Address", vbCritical
150           Exit Sub
160       End If

170       sql = "SELECT MBCode FROM ExternalDefinitions WHERE " & _
                "MBCode = '" & txtCode & "' " & _
                "AND Department = '" & pDepartment & "'"
180       Set tb = New Recordset
190       RecOpenServer 0, tb, sql
200       If Not tb.EOF Then
210           iMsg "Code already used!" & vbCrLf & "Enter new Code", vbCritical
220           txtCode = ""
230           Exit Sub
240       End If

250       sql = "SELECT * FROM ExternalDefinitions WHERE 0 = 1"
260       Set tb = New Recordset
270       RecOpenServer 0, tb, sql
280       tb.AddNew
290       tb!AnalyteName = txtTestName
300       tb!MBCode = txtCode
310       tb!SendTo = cmbSendTo
320       tb!Units = txtUnits
330       tb!NormalRange = txtNormalRange
340       tb!Department = pDepartment
350       tb!Comment = txtComment
360       tb!InUse = 1
370       tb!SampleType = txtSampleType
380       tb.Update

390       FillG

400       txtTestName = ""
410       cmbSendTo = ""
420       txtUnits = ""
430       txtNormalRange = ""
440       txtCode = ""
450       txtComment = ""
460       txtSampleType = ""

470       Exit Sub

cmdAddToList_Click_Error:

          Dim strES As String
          Dim intEL As Integer

480       intEL = Erl
490       strES = Err.Description
500       LogError "frmExtTests", "cmdAddToList_Click", intEL, strES, sql

End Sub

Private Sub cmdCancel_Click()

10        Unload Me

End Sub

Private Sub cmdDelete_Click()

          Dim sql As String

10        On Error GoTo cmdDelete_Click_Error

20        g.Col = 1
30        If Trim(g) = "" Then Exit Sub

40        sql = "DELETE FROM ExternalDefinitions WHERE EntryNum = " & g
50        Cnxn(0).Execute sql

60        FillG

70        Exit Sub

cmdDelete_Click_Error:

          Dim strES As String
          Dim intEL As Integer

80        intEL = Erl
90        strES = Err.Description
100       LogError "frmExtTests", "cmdDelete_Click", intEL, strES, sql

End Sub


Private Sub cmdExcel_Click()

          Dim s As String

10        On Error GoTo cmdExcel_Click_Error

20        If g.Rows = 2 Then
30            iMsg "Nothing to export", vbInformation
40            Exit Sub
50        End If

60        s = "External Tests List" & vbCr

70        ExportFlexGrid g, Me, s

80        Exit Sub

cmdExcel_Click_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmExtTests", "cmdExcel_Click", intEL, strES

End Sub

Private Sub cmdPrint_Click()

          Dim Num As Long

10        On Error GoTo cmdPrint_Click_Error

20        Printer.Print
30        Printer.Font.Name = "Courier New"
40        Printer.Font.Size = 12

50        For Num = 0 To g.Rows - 1

60            Printer.Print g.TextMatrix(Num, 0);
70            Printer.Print Tab(5); g.TextMatrix(Num, 1);
80            Printer.Print Tab(15); g.TextMatrix(Num, 2);
90            Printer.Print Tab(35); g.TextMatrix(Num, 3);
100           Printer.Print Tab(50); g.TextMatrix(Num, 4);
110           Printer.Print Tab(70); g.TextMatrix(Num, 5);
120           Printer.Print Tab(6); g.TextMatrix(Num, 6)

130       Next

140       Printer.EndDoc

150       Exit Sub

cmdPrint_Click_Error:

          Dim strES As String
          Dim intEL As Integer

160       intEL = Erl
170       strES = Err.Description
180       LogError "frmExtTests", "cmdPrint_Click", intEL, strES

End Sub

Private Sub FillAddress()

          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo FillAddress_Error

20        sql = "SELECT Addr0 FROM ExtAddress"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql

50        cmbSendTo.Clear
60        Do While Not tb.EOF
70            cmbSendTo.AddItem tb!Addr0 & ""
80            tb.MoveNext
90        Loop

100       Exit Sub

FillAddress_Error:

          Dim strES As String
          Dim intEL As Integer

110       intEL = Erl
120       strES = Err.Description
130       LogError "frmExtTests", "FillAddress", intEL, strES, sql

End Sub

Private Sub FillG()

          Dim sql As String
          Dim tb As New Recordset

10        On Error GoTo FillG_Error

20        ClearFGrid g

30        sql = "SELECT " & _
                "(ISNULL(MBCode, '') + CHAR(9) + " & _
                "AnalyteName + CHAR(9) + " & _
                "ISNULL(NormalRange, '') + CHAR(9) + " & _
                "ISNULL(Units, '') + CHAR(9) + " & _
                "SendTo + CHAR(9) + " & _
                "ISNULL(Comment, '')  + CHAR(9) + " & _
                "CASE COALESCE(InUse, 0) WHEN 0 THEN 'No' WHEN 1 THEN 'Yes' END + CHAR(9) + " & _
                "COALESCE(SampleType, '') ) R " & _
                "FROM ExternalDefinitions WHERE " & _
                "Department = '" & pDepartment & "' " & _
                "ORDER BY AnalyteName"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        Do While Not tb.EOF
70            g.AddItem tb!R
80            tb.MoveNext
90        Loop

100       FixG g
110       If g.Rows > 1 Then
120           lblABC = UCase(Left(g.TextMatrix(g.TopRow, 1), 1))
130       End If

140       Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

150       intEL = Erl
160       strES = Err.Description
170       LogError "frmExtTests", "FillG", intEL, strES, sql

End Sub

Private Sub cmdSave_Click()

          Dim sql As String
          Dim Y As Integer

10        On Error GoTo cmdSave_Click_Error

20        cmdSave.Caption = "Saving..."
30        cmdSave.Refresh

40        For Y = 1 To g.Rows - 1
50            If g.TextMatrix(Y, 0) <> "" Then
60                sql = "IF EXISTS (SELECT * FROM ExternalDefinitions WHERE MBCode = '" & g.TextMatrix(Y, 0) & "' " & _
                        "AND Department = '" & cmbDepartment & "') " & _
                        "  UPDATE ExternalDefinitions " & _
                        "  SET AnalyteName = '" & AddTicks(g.TextMatrix(Y, 1)) & "', " & _
                        "  NormalRange = '" & g.TextMatrix(Y, 2) & "', " & _
                        "  Units = '" & g.TextMatrix(Y, 3) & "', " & _
                        "  SendTo = '" & AddTicks(g.TextMatrix(Y, 4)) & "', " & _
                        "  Comment = '" & AddTicks(g.TextMatrix(Y, 5)) & "', " & _
                        "  PrintPriority = " & Y & ", " & _
                        "  InUse = " & IIf(g.TextMatrix(Y, 6) = "Yes", 1, 0) & ", " & _
                        "  SampleType = '" & g.TextMatrix(Y, 7) & "' " & _
                        "  WHERE MBCode = '" & g.TextMatrix(Y, 0) & "' " & _
              "AND Department = '" & cmbDepartment & "' "
70                sql = sql & _
                        "ELSE " & _
                        "  INSERT INTO ExternalDefinitions (MBCode, AnalyteName, PrintPriority, Units, NormalRange, SendTo, SampleType, Comment, Department, InUse) " & _
                        "  VALUES " & _
                        "  ('" & g.TextMatrix(Y, 0) & "', " & _
                        "  '" & AddTicks(g.TextMatrix(Y, 1)) & "', " & _
                        "  " & Y & ", " & _
                        "  '" & g.TextMatrix(Y, 3) & "', " & _
                        "  '" & g.TextMatrix(Y, 2) & "', " & _
                        "  '" & AddTicks(g.TextMatrix(Y, 4)) & "', " & _
                        "  '" & g.TextMatrix(Y, 7) & "', " & _
                        "  '" & AddTicks(g.TextMatrix(Y, 5)) & "', " & _
                        "  '" & cmbDepartment & "', " & _
                        IIf(g.TextMatrix(Y, 6) = "Yes", "1", "0") & " )"
80                Cnxn(0).Execute sql
90            End If
100       Next

110       cmdSave.Caption = "&Save Changes"
120       cmdSave.Visible = False

130       Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

140       intEL = Erl
150       strES = Err.Description
160       LogError "frmExtTests", "cmdSave_Click", intEL, strES, sql

End Sub

Private Sub Form_Load()

          Dim sql As String

10        On Error GoTo Form_Load_Error

20        PopulateDepartments
      '    cmbDepartment.Clear
      '    cmbDepartment.AddItem "General"
      '    cmbDepartment.AddItem "Haematology"
      '    cmbDepartment.AddItem "Micro"
      '    cmbDepartment.Text = pDepartment

30        FillAddress
40        FillG

50        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmExtTests", "Form_Load", intEL, strES, sql


End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

10        If cmdSave.Visible Then
20            If iMsg("Cancel without Saving?", vbYesNo) = vbNo Then
30                Cancel = True
40            End If
50        End If

End Sub


Private Sub g_Click()

          Dim sql As String
          Dim tb As New Recordset
          Static SortOrder As Boolean
          Dim s As String
          Dim TestName As String
          Dim ip As String
10        ReDim Options(0 To 0) As String
          Dim n As Integer
          Dim f As Form
          Dim gRow As Integer

20        On Error GoTo g_Click_Error

30        If g.MouseRow = 0 Then
40            If SortOrder Then
50                g.Sort = flexSortGenericAscending
60            Else
70                g.Sort = flexSortGenericDescending
80            End If
90            SortOrder = Not SortOrder
100           Exit Sub
110       End If

120       gRow = g.MouseRow

130       TestName = g.TextMatrix(gRow, 1)

140       Select Case g.Col
          Case 0, 1:
150           s = "Remove " & TestName & " from list?"
160           If iMsg(s, vbQuestion + vbYesNo) = vbNo Then
170               Exit Sub
180           End If

190           sql = "SELECT * FROM ExtPanels WHERE " & _
                    "Content = '" & TestName & "' " & _
                    "AND Department = '" & cmbDepartment.Text & "'"
200           Set tb = New Recordset
210           RecOpenServer 0, tb, sql
220           If Not tb.EOF Then
230               s = "Cannot Remove " & TestName & " from list." & vbCrLf & _
                      "It is used in " & tb!PanelName & " Panel."
240               iMsg s, vbCritical
250               Exit Sub
260           End If

270           sql = "SELECT TOP 1 * FROM ExtResults WHERE Analyte = '" & TestName & "'"
280           Set tb = New Recordset
290           RecOpenServer 0, tb, sql
300           If Not tb.EOF Then
310               s = "Cannot Remove " & TestName & " from list." & vbCrLf & "Results are present."
320               iMsg s, vbCritical
330               Exit Sub
340           End If

350           sql = "SELECT * FROM ExternalDefinitions WHERE AnalyteName = '" & TestName & "'"
360           Set tb = New Recordset
370           RecOpenServer 0, tb, sql

380           If Not tb.EOF Then
390               txtCode = Trim(tb!MBCode & "")
400               txtTestName = Trim(tb!AnalyteName & "")
410               txtNormalRange = Trim(tb!NormalRange & "")
420               cmbSendTo = Trim(tb!SendTo & "")
430               txtUnits = Trim(tb!Units & "")
440               txtComment = Trim$(tb!Comment & "")
450               txtSampleType = Trim$(tb!SampleType & "")
460           End If

470           sql = "DELETE from ExternalDefinitions WHERE AnalyteName = '" & TestName & "'"
480           Cnxn(0).Execute sql

490       Case 2:
500           ip = iBOX("Enter Normal Range for " & TestName, , g.TextMatrix(gRow, 2))
510           If Trim(ip) = "" Then
520               If iMsg("Blank Normal Range?", vbYesNo) = vbNo Then
530                   ip = g.TextMatrix(gRow, 2)
540               End If
550           End If
560           g.TextMatrix(gRow, 2) = ip
570           cmdSave.Visible = True

580       Case 3:
590           sql = "SELECT Text FROM Lists WHERE ListType = 'UN'"
600           Set tb = New Recordset
610           RecOpenServer 0, tb, sql
620           n = -1
630           Do While Not tb.EOF
640               n = n + 1
650               ReDim Preserve Options(0 To n) As String
660               Options(n) = tb!Text & ""
670               tb.MoveNext
680           Loop
690           Set f = New fcdrDBox
700           With f
710               .Options = Options
720               .Prompt = "Select Units"
730               .Show 1
740               ip = .ReturnValue
750           End With
760           Unload f
770           Set f = Nothing
780           If Trim(ip) = "" Then
790               If iMsg("Blank Units?", vbYesNo) = vbNo Then
800                   ip = g.TextMatrix(gRow, 3)
810               End If
820           End If
830           g.TextMatrix(gRow, 3) = ip
840           cmdSave.Visible = True

850       Case 4
860           sql = "SELECT Code FROM ExtAddress"
870           Set tb = New Recordset
880           RecOpenServer 0, tb, sql
890           n = -1
900           Do While Not tb.EOF
910               n = n + 1
920               ReDim Preserve Options(0 To n) As String
930               Options(n) = tb!Code & ""
940               tb.MoveNext
950           Loop
960           Set f = New fcdrDBox
970           With f
980               .Options = Options
990               .Prompt = "Select Address"
1000              .Show 1
1010              ip = .ReturnValue
1020          End With
1030          Unload f
1040          Set f = Nothing
1050          If Trim(ip) = "" Then
1060              If iMsg("Blank Address?", vbYesNo) = vbNo Then
1070                  ip = g.TextMatrix(gRow, 4)
1080              End If
1090          End If
1100          g.TextMatrix(gRow, 4) = ip
1110          cmdSave.Visible = True

1120      Case 5:
1130          ip = iBOX("Enter Comment for " & TestName, , g.TextMatrix(gRow, 5))
1140          If Trim(ip) = "" Then
1150              If iMsg("Blank Comment?", vbYesNo) = vbNo Then
1160                  ip = g.TextMatrix(gRow, 5)
1170              End If
1180          End If
1190          g.TextMatrix(gRow, 5) = ip
1200          cmdSave.Visible = True
1210      Case 6:
1220          g.TextMatrix(gRow, 6) = IIf(g.TextMatrix(gRow, 6) = "Yes", "No", "Yes")
1230          cmdSave.Visible = True
1240      Case 7:
1250          ip = iBOX("Enter sample type for " & TestName, , g.TextMatrix(gRow, 7))
1260          If Trim(ip) = "" Then
1270              If iMsg("Blank sample type?", vbYesNo) = vbNo Then
1280                  ip = g.TextMatrix(gRow, 7)
1290              End If
1300          End If
1310          g.TextMatrix(gRow, 7) = ip
1320          cmdSave.Visible = True
              '        Select Case Trim$(UCase$(g.TextMatrix(gRow, 7)))
              '            Case "": g.TextMatrix(gRow, 7) = "SER^SERUM"
              '            Case "SER^SERUM": g.TextMatrix(gRow, 7) = "BLD^WHOLE BLOOD"
              '            Case "BLD^WHOLE BLOOD": g.TextMatrix(gRow, 7) = "CSF^CSF"
              '            Case "CSF^CSF": g.TextMatrix(gRow, 7) = ""
              '        End Select
              '        cmdSave.Visible = True
1330      End Select

1340      Exit Sub

g_Click_Error:

          Dim strES As String
          Dim intEL As Integer

1350      intEL = Erl
1360      strES = Err.Description
1370      LogError "frmExtTests", "g_Click", intEL, strES, sql

End Sub

Private Sub cmbSendTo_KeyPress(KeyAscii As Integer)

10        KeyAscii = 0

End Sub


Public Property Let Department(ByVal sNewValue As String)

10        pDepartment = sNewValue

End Property





Private Sub g_Scroll()
10        lblABC = UCase(Left(g.TextMatrix(g.TopRow, 1), 1))
End Sub


Private Sub PopulateDepartments()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo PopulateDepartments_Error


20        With cmbDepartment
30            .Clear
              '.AddItem "All Departments"
40            sql = "SELECT DISTINCT Department FROM ExternalDefinitions " & _
                    "ORDER BY Department"
50            Set tb = New Recordset
60            RecOpenClient 0, tb, sql
70            If Not tb.EOF Then
80                While Not tb.EOF
90                    .AddItem tb!Department & ""
100                   tb.MoveNext
110               Wend
120           End If
130           .ListIndex = 0
140       End With

150       Exit Sub

PopulateDepartments_Error:

          Dim strES As String
          Dim intEL As Integer

160       intEL = Erl
170       strES = Err.Description
180       LogError "frmAddToTests", "PopulateDepartments", intEL, strES, sql

End Sub


