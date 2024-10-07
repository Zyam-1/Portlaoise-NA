VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAddToTests 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add to External Tests Requested"
   ClientHeight    =   8490
   ClientLeft      =   960
   ClientTop       =   480
   ClientWidth     =   13980
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8490
   ScaleWidth      =   13980
   Begin MSComctlLib.TreeView tv 
      Height          =   7575
      Left            =   360
      TabIndex        =   11
      Top             =   600
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   13361
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      SingleSel       =   -1  'True
      Appearance      =   1
   End
   Begin VB.ComboBox cmbDepartment 
      BackColor       =   &H0080FFFF&
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
      ItemData        =   "frmAddToTests.frx":0000
      Left            =   9000
      List            =   "frmAddToTests.frx":0002
      TabIndex        =   9
      Text            =   "Combo1"
      Top             =   90
      Width           =   3555
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   12825
      Picture         =   "frmAddToTests.frx":0004
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5040
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   5385
      Left            =   6360
      TabIndex        =   4
      Top             =   690
      Width           =   6165
      _ExtentX        =   10874
      _ExtentY        =   9499
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      FormatString    =   "<Analyte                  |<Sample Type  |<Destination    |<Department    "
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   675
      Index           =   0
      Left            =   4350
      TabIndex        =   2
      Top             =   690
      Width           =   1965
      _Version        =   65536
      _ExtentX        =   3466
      _ExtentY        =   1191
      _StockProps     =   15
      Caption         =   "Test required    "
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   2
      BorderWidth     =   2
      BevelOuter      =   0
      BevelInner      =   1
      Alignment       =   4
      Begin VB.Image Image1 
         Height          =   480
         Left            =   150
         Picture         =   "frmAddToTests.frx":0ECE
         Top             =   90
         Width           =   480
      End
   End
   Begin VB.ListBox lpanels 
      Height          =   5730
      IntegralHeight  =   0   'False
      Left            =   4350
      TabIndex        =   1
      Top             =   2520
      Width           =   1965
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   12825
      Picture         =   "frmAddToTests.frx":1310
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7260
      Width           =   975
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   855
      Index           =   1
      Left            =   4350
      TabIndex        =   3
      Top             =   1560
      Width           =   1965
      _Version        =   65536
      _ExtentX        =   3466
      _ExtentY        =   1508
      _StockProps     =   15
      Caption         =   "Panel required    "
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   2
      BorderWidth     =   2
      BevelOuter      =   0
      BevelInner      =   1
      Alignment       =   6
      Begin VB.Image Image2 
         Height          =   480
         Left            =   630
         Picture         =   "frmAddToTests.frx":21DA
         Top             =   270
         Width           =   480
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Department"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6420
      TabIndex        =   10
      Top             =   90
      Width           =   2550
   End
   Begin VB.Label lblSampleID 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "12345678"
      Height          =   285
      Left            =   12705
      TabIndex        =   7
      Top             =   2850
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "SampleID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   12960
      TabIndex        =   6
      Top             =   2640
      Width           =   690
   End
   Begin VB.Label lblComment 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   2145
      Left            =   6390
      TabIndex        =   5
      Top             =   6120
      Width           =   6165
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmAddToTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim NodX As MSComctlLib.Node
Public FromEdit As Boolean

Private Activated As Boolean

Private pDepartment As String

Private pSampleID As String

Private pClinDetails As String
Private pSampleDate As String
Private pSampleTime As String
Private pSex As String
Private pGP As String
Private pWard As String
Private pClinician As String


Private Sub FillAll()

10    FillTV
20    FillPanels
30    FillOrders
40    FillDSSI

End Sub

Private Sub FillDSSI()
'
'10    cmbDSSI.AddItem "AU -Audiology"
'20    cmbDSSI.AddItem "BG -Blood gases"
'30    cmbDSSI.AddItem "BLB-Blood bank"
'40    cmbDSSI.AddItem "CUS-Cardiac Ultrasound"
'50    cmbDSSI.AddItem "CTH-Cardiac catheterization"
'60    cmbDSSI.AddItem "CT -CAT scan"
'70    cmbDSSI.AddItem "CH -Chemistry"
'80    cmbDSSI.AddItem "CP -Cytopathology"
'90    cmbDSSI.AddItem "EC -Electrocardiac (e.g., EKG, EEC, Holter)"
'100   cmbDSSI.AddItem "EN -Electroneuro(EEG, EMG, EP, PSG)"
'110   cmbDSSI.AddItem "HM -Hematology"
'120   cmbDSSI.AddItem "ICU-Bedside ICU Monitoring"
'130   cmbDSSI.AddItem "IMG-Diagnostic Imaging"
'140   cmbDSSI.AddItem "IMM-Immunology"
'150   cmbDSSI.AddItem "LAB-Laboratory"
'160   cmbDSSI.AddItem "MB -Microbiology"
'170   cmbDSSI.AddItem "MCB-Mycobacteriology"
'180   cmbDSSI.AddItem "MYC-Mycology"
'190   cmbDSSI.AddItem "NMS-Nuclear medicine scan"
'200   cmbDSSI.AddItem "NMR-Nuclear magnetic resonance"
'210   cmbDSSI.AddItem "NRS-Nursing service measures"
'220   cmbDSSI.AddItem "OUS-OB Ultrasound"
'230   cmbDSSI.AddItem "OT -Occupational Therapy"
'240   cmbDSSI.AddItem "OTH-Other"
'250   cmbDSSI.AddItem "OSL-Outside Lab"
'260   cmbDSSI.AddItem "PAR-Parasitology"
'270   cmbDSSI.AddItem "PAT-Pathology(gross & histopath, Not Surgical)"
'280   cmbDSSI.AddItem "PHR-Pharmacy"
'290   cmbDSSI.AddItem "PT -Physical Therapy"
'300   cmbDSSI.AddItem "PHY-Physician (Hx. Dx, admission note, etc.)"
'310   cmbDSSI.AddItem "PF -Pulmonary function"
'320   cmbDSSI.AddItem "RAD-Radiology"
'330   cmbDSSI.AddItem "RX -Radiograph"
'340   cmbDSSI.AddItem "RUS-Radiology ultrasound"
'350   cmbDSSI.AddItem "RC -Respiratory Care (therapy)"
'360   cmbDSSI.AddItem "RT -Radiation therapy"
'370   cmbDSSI.AddItem "SR -Serology"
'380   cmbDSSI.AddItem "SP -Surgical"
'390   cmbDSSI.AddItem "TX -Toxicology"
'400   cmbDSSI.AddItem "URN-Urinalysis"
'410   cmbDSSI.AddItem "VUS-Vascular Ultrasound"
'420   cmbDSSI.AddItem "VR -Virology"
'430   cmbDSSI.AddItem "XRC-Cineradiograph"
'
'440   cmbDSSI = "LAB-Laboratory"

End Sub

Sub FillPanels()
Attribute FillPanels.VB_Description = "Load Panels"

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo FillPanels_Error

20    If pDepartment = "All Departments" Then
30        sql = "SELECT DISTINCT PanelName FROM ExtPanels " & _
                "WHERE LEFT(Department, 5) <> 'Micro' " & _
                "ORDER BY PanelName"
40    Else
50        sql = "SELECT DISTINCT PanelName FROM ExtPanels " & _
                "WHERE Department = '" & pDepartment & "' " & _
                "ORDER BY PanelName"
60    End If


70    Set tb = New Recordset
80    RecOpenServer 0, tb, sql

90    lpanels.Clear

100   Do While Not tb.EOF
110       lpanels.AddItem tb!PanelName
120       tb.MoveNext
130   Loop

140   Exit Sub

FillPanels_Error:

      Dim strES As String
      Dim intEL As Integer

150   intEL = Erl
160   strES = Err.Description
170   LogError "frmAddToTests", "FillPanels", intEL, strES, sql

End Sub

Private Sub cmbDepartment_Click()

10    pDepartment = cmbDepartment
20    FillAll

End Sub


Private Sub cmbDepartment_KeyPress(KeyAscii As Integer)

10    KeyAscii = 0

End Sub


Private Sub cmdCancel_Click()

10    Unload Me

End Sub




Private Sub PlaceOrder()

      Dim tb As Recordset
      Dim sql As String
      Dim n As Integer
      Dim SampleDateTime As String
      Dim Department As String
      Dim SendTo As String
      Dim Units As String
      Dim Analyte As String
      Dim MBCode As String
      Dim HospitalName As String

10    On Error GoTo PlaceOrder_Error

      '20        HospitalName = GetOptionSetting("STJAMESHOSPITAL", "St James Hospital")



20    If IsDate(pSampleTime) Then
30        SampleDateTime = pSampleDate & " " & pSampleTime
40    Else
50        SampleDateTime = pSampleDate & " 00:01"
60    End If
70    SampleDateTime = Format$(SampleDateTime, "dd/mmm/yyyy HH:mm")

80    For n = 1 To g.Rows - 1
90        Analyte = g.TextMatrix(n, 0)
100       sql = "Select * from ExternalDefinitions where " & _
                "AnalyteName = '" & Analyte & "'"
110       Set tb = New Recordset
120       RecOpenServer 0, tb, sql
130       If Not tb.EOF Then
              '150               If UCase(HospitalName) = UCase(tb!SendTo & "") Then
140           If ListCodeFor("MBEnableLabs", UCase(g.TextMatrix(n, 2))) <> "" Then 'ISITEMINLIST(tb!SendTo, "MBEnableLabs") Then
150               SendTo = tb!SendTo & ""
160               Units = tb!Units & ""
170               Department = tb!Department & ""
180               MBCode = tb!MBCode & ""

190               sql = "IF NOT EXISTS ( SELECT * FROM MedibridgeRequests WHERE " & _
                        "                SampleID = '" & pSampleID & "' " & _
                        "                AND TestName = '" & Analyte & "') " & _
                        "  INSERT INTO MedibridgeRequests " & _
                        "  (SampleID, TestCode, TestName, SampleDateTime, ClinDetails, Orderer, Dept, SpecimenSource, Status) " & _
                        "  VALUES " & _
                        "  ('" & pSampleID & "', " & _
                        "  '" & MBCode & "', " & _
                        "  '" & Analyte & "', " & _
                        "  '" & SampleDateTime & "', " & _
                        "  '" & pClinDetails & "', " & _
                        "  '" & UserCode & "^" & UserName & "', " & _
                        "  '" & Department & "', " & _
                        "  '" & g.TextMatrix(n, 1) & "'," & _
                        "  'Requested')"
200               Cnxn(0).Execute sql
210           ElseIf ListCodeFor("BiomnisEnableLabs", UCase(g.TextMatrix(n, 2))) <> "" Then


220               sql = "DECLARE @Code nvarchar(50) " & _
                        "DECLARE @Dept nvarchar(50) " & _
                        "DECLARE @SampleType nvarchar(50) " & _
                        "SET @Code = (SELECT BiomnisCode FROM ExternalDefinitions " & _
                        "             WHERE AnalyteName = '" & g.TextMatrix(n, 0) & "') "

230               sql = sql & _
                        "SET @Dept = (SELECT Department FROM ExternalDefinitions " & _
                        "             WHERE AnalyteName = '" & g.TextMatrix(n, 0) & "') "

240               sql = sql & _
                        "SET @SampleType = (SELECT SampleType FROM ExternalDefinitions " & _
                        "             WHERE AnalyteName = '" & g.TextMatrix(n, 0) & "') "
250               sql = sql & _
                        "IF EXISTS(SELECT * FROM ExternalDefinitions " & _
                        "          WHERE AnalyteName = '" & g.TextMatrix(n, 0) & "' " & _
                        "          AND COALESCE(BiomnisCode, '') <> '') " & _
                        "  BEGIN IF NOT EXISTS(SELECT * FROM BiomnisRequests " & _
                        "                  WHERE SampleID = '" & pSampleID & "' " & _
                        "                  AND TestCode = @Code)" & _
                        "      INSERT INTO BiomnisRequests (SampleID, TestCode, TestName, SampleType, SampleDateTime, Department, RequestedBy, SendTo, Status) " & _
                        "      VALUES " & _
                        "     ('" & pSampleID & "', " & _
                        "      @Code, " & _
                        "      '" & g.TextMatrix(n, 0) & "', " & _
                        "      @SampleType, " & _
                        "      '" & SampleDateTime & "', " & _
                        "      @Dept, " & _
                        "      '" & UserCode & "^" & UserName & "', " & _
                        "      '" & g.TextMatrix(n, 2) & "', " & _
                        "      'OutStanding') END"
260               Cnxn(0).Execute sql



270           End If
280       End If
290   Next

300   cmdSave.Enabled = False

310   SaveRequests

320   If GetOptionSetting("ExtPrintRequests", 0) = 1 Then
330       PrintRequests
340   End If



350   Exit Sub

PlaceOrder_Error:

      Dim strES As String
      Dim intEL As Integer

360   intEL = Erl
370   strES = Err.Description
380   LogError "frmAddToTests", "PlaceOrder", intEL, strES, sql

End Sub

Private Sub PrintRequests()

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo PrintRequests_Error

20    LogTimeOfPrinting pSampleID, "X"
30    sql = "SELECT * FROM PrintPending WHERE " & _
            "Department = 'X' " & _
            "AND SampleID = '" & pSampleID & "' " & _
            "AND (FaxNumber = '' OR FaxNumber IS NULL)"
40    Set tb = New Recordset
50    RecOpenClient 0, tb, sql
60    If tb.EOF Then
70        tb.AddNew
80    End If
90    tb!SampleID = pSampleID
100   tb!Department = "X"
110   tb!Initiator = UserName
120   tb!Ward = pWard
130   tb!Clinician = pClinician
140   tb!GP = pGP
150   tb!UsePrinter = ""
160   tb!pTime = Now
170   tb.Update

180   Exit Sub

PrintRequests_Error:

      Dim strES As String
      Dim intEL As Integer

190   intEL = Erl
200   strES = Err.Description
210   LogError "frmAddToTests", "PrintRequests", intEL, strES, sql

End Sub

Private Sub SaveRequests()

      Dim tb As Recordset
      Dim sql As String
      Dim n As Integer
      Dim SampleDateTime As String
      Dim Department As String
      Dim SendTo As String
      Dim Units As String
      Dim Analyte As String
      Dim MBCode As String

10    On Error GoTo SaveRequests_Error

20    If IsDate(pSampleTime) Then
30        SampleDateTime = pSampleDate & " " & pSampleTime
40    Else
50        SampleDateTime = pSampleDate & " 00:01"
60    End If
70    SampleDateTime = Format$(SampleDateTime, "dd/mmm/yyyy HH:mm")

80    For n = 1 To g.Rows - 1
90        Analyte = g.TextMatrix(n, 0)
100       sql = "Select * from ExternalDefinitions where " & _
                "AnalyteName = '" & Analyte & "'"
110       Set tb = New Recordset
120       RecOpenServer 0, tb, sql
130       If Not tb.EOF Then
140           SendTo = tb!SendTo & ""
150           Units = tb!Units & ""
160           Department = tb!Department & ""
170           MBCode = tb!MBCode & ""


180           sql = "IF NOT EXISTS ( SELECT * FROM ExtResults WHERE " & _
                    "                SampleID = '" & pSampleID & "' " & _
                    "                AND Analyte = '" & Analyte & "') " & _
                    "  INSERT INTO ExtResults " & _
                    "  (SampleID, Analyte, Result, SendTo, Units, Date, RetDate, SentDate, " & _
                    "   SAPCode, HealthLink, OrderList, SaveTime, UserName, Valid, Department) " & _
                    "  VALUES " & _
                    "  ('" & pSampleID & "', " & _
                    "   '" & Analyte & "', " & _
                    "   Null, " & _
                    "   '" & AddTicks(SendTo) & "', " & _
                    "   '" & Units & "', " & _
                    "   Null, " & _
                    "   Null, " & _
                    "   getdate(), " & _
                    "   Null, " & _
                    "   '0', " & _
                    "   '" & n & "', " & _
                    "   Null, " & _
                    "   '" & AddTicks(UserName) & "', " & _
                    "   '0', '" & Department & "')"
190           Cnxn(0).Execute sql

200       End If
210   Next

220   cmdSave.Enabled = False

230   Exit Sub

SaveRequests_Error:

      Dim strES As String
      Dim intEL As Integer

240   intEL = Erl
250   strES = Err.Description
260   LogError "frmAddToTests", "SaveRequests", intEL, strES, sql

End Sub


Private Sub cmdSave_Click()

10    PlaceOrder

End Sub




Private Sub Form_Activate()

10    If Not Activated Then
20        FillAll
30        tv.SetFocus
40        For Each NodX In tv.Nodes
50            NodX.Expanded = False
60        Next
70        Activated = True
80    End If

End Sub
Private Sub FillOrders()

      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo FillOrders_Error

20    g.Rows = 2
30    g.AddItem ""
40    g.RemoveItem 1

50    sql = "Select * from ExtResults where " & _
            "SampleID = '" & Val(pSampleID) & "'"
60    Set tb = New Recordset
70    RecOpenServer 0, tb, sql
80    Do While Not tb.EOF
90        g.AddItem tb!Analyte & "" & vbTab & _
                    "" & vbTab & _
                    tb!SendTo & "" & vbTab & _
                    tb!Department & ""
100       tb.MoveNext
110   Loop

120   If g.Rows > 2 Then
130       g.RemoveItem 1
140   End If

150   Exit Sub

FillOrders_Error:

      Dim strES As String
      Dim intEL As Integer

160   intEL = Erl
170   strES = Err.Description
180   LogError "frmAddToTests", "FillOrders", intEL, strES, sql

End Sub

Sub FillTV()
Attribute FillTV.VB_Description = "Fill Ndal List"
      Dim NodX As MSComctlLib.Node
      Dim n As Integer
      Dim Relative As String
      Dim ThisNode As String
      Dim sql As String
      Dim tb As Recordset
      Dim Key As String
      Dim NodeText As String

10    On Error GoTo FillTV_Error

20    tv.Visible = False

30    tv.Nodes.Clear

40    For n = Asc("A") To Asc("Z")
50        Key = chr$(n)
60        NodeText = chr$(n)
70        Set NodX = tv.Nodes.Add(, , Key, NodeText)
80    Next
90    For n = Asc("0") To Asc("9")
100       Set NodX = tv.Nodes.Add(, , "#" & chr$(n), chr$(n))
110   Next

120   If UCase(pDepartment) = UCase("All Departments") Then
130       sql = "SELECT AnalyteName FROM ExternalDefinitions " & _
                "WHERE InUse = 1 " & _
                "ORDER BY AnalyteName"
140   Else

150       sql = "SELECT AnalyteName FROM ExternalDefinitions " & _
                "WHERE Department = '" & pDepartment & "' AND InUse = 1 " & _
                "ORDER BY AnalyteName"
160   End If

170   Set tb = New Recordset
180   RecOpenServer 0, tb, sql
190   Do While Not tb.EOF
200       If Trim$(tb!AnalyteName & "") <> "" Then
210           Relative = UCase(Left(tb!AnalyteName, 1))
220           If IsNumeric(Relative) Then Relative = "#" & Relative
230           ThisNode = tb!AnalyteName
240           Set NodX = tv.Nodes.Add(Relative, tvwChild, , ThisNode)
250       End If
260       tb.MoveNext
270   Loop
280   tv.Visible = True


'290   tv.SetFocus
'300   tv.SelectedItem.Expanded = False

310   Exit Sub

FillTV_Error:

      Dim strES As String
      Dim intEL As Integer

320   intEL = Erl
330   strES = Err.Description
340   LogError "frmAddToTests", "FillTV", intEL, strES, sql
350   tv.Visible = True
End Sub


Private Sub Form_Load()

10    Activated = False

20    If UCase(pDepartment) = "MICRO" Then
30        lblSampleID = pSampleID - SysOptMicroOffset(0)
40    Else
50        lblSampleID = pSampleID
60    End If

70    PopulateDepartments

      'cmbDepartment.Clear
      'cmbDepartment.AddItem "General"
      'cmbDepartment.AddItem "Immunology"
      'cmbDepartment.AddItem "Haematology"
      'cmbDepartment.AddItem "Micro"

80    cmbDepartment = pDepartment

End Sub

Private Sub PopulateDepartments()

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo PopulateDepartments_Error


20    With cmbDepartment
30        .Clear

40        If UCase(pDepartment) = "MICRO" Then
50            sql = "SELECT DISTINCT Department FROM ExternalDefinitions " & _
                    "WHERE UPPER(LEFT(Department, 5)) = 'MICRO' " & _
                    "ORDER BY Department"
60        Else
70            .AddItem "All Departments"
80            sql = "SELECT DISTINCT Department FROM ExternalDefinitions " & _
                    "WHERE UPPER(LEFT(Department, 5)) <> 'MICRO' " & _
                    "ORDER BY Department"
90        End If
          '    sql = "SELECT DISTINCT Department FROM ExternalDefinitions " & _
               '          "ORDER BY Department"
100       Set tb = New Recordset
110       RecOpenClient 0, tb, sql
120       If Not tb.EOF Then
130           While Not tb.EOF
140               .AddItem tb!Department & ""
150               tb.MoveNext
160           Wend
170       End If
180       .ListIndex = 0
190   End With

200   Exit Sub

PopulateDepartments_Error:

      Dim strES As String
      Dim intEL As Integer

210   intEL = Erl
220   strES = Err.Description
230   LogError "frmAddToTests", "PopulateDepartments", intEL, strES, sql

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)


10    If cmdSave.Enabled Then
20        If iMsg("Cancel without Saving?", vbYesNo) = vbNo Then
30            Cancel = True
40        End If
50    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

10    Activated = False

End Sub

Private Sub g_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
      Dim s As String
      Dim n As Integer
      Dim R As Integer

10    R = g.MouseRow

20    s = "Remove " & g.TextMatrix(R, 0) & " from tests requested?"
30    n = iMsg(s, vbYesNo + vbQuestion)
40    If n = vbYes Then
50        If g.Rows = 2 Then
60            g.AddItem ""
70            g.RemoveItem 1
80        Else
90            g.RemoveItem R
100       End If
110   End If

End Sub


Private Sub lpanels_Click()

      Dim tb As Recordset
      Dim sql As String
      Dim n As Integer
      Dim Found As Boolean
      Dim s As String
      Dim XDep As String

10    On Error GoTo lpanels_Click_Error
20    If pDepartment = "All Departments" Then
30        XDep = "%"
40    Else
50        XDep = pDepartment
60    End If

70    sql = "SELECT P.TestName, D.SampleType, D.SendTo, D.Comment, D.Department " & _
            "FROM ExtPanels P JOIN (SELECT * FROM ExternalDefinitions WHERE InUse = 1 AND Department like '" & XDep & "') D " & _
            "ON P.TestName COLLATE DATABASE_DEFAULT = D.AnalyteName COLLATE DATABASE_DEFAULT " & _
            "WHERE PanelName = '" & lpanels & "'"
80    Set tb = New Recordset
90    RecOpenServer 0, tb, sql
100   Do While Not tb.EOF
110       Found = False
120       For n = 1 To g.Rows - 1
130           If g.TextMatrix(n, 0) = tb!TestName & "" Then
140               Found = True
150               Exit For
160           End If
170       Next
180       If Not Found Then
190           s = tb!TestName & vbTab & tb!SampleType & vbTab & tb!SendTo & "" & vbTab & tb!Department & ""
200           g.AddItem s


210           cmdSave.Enabled = True

220           If Trim$(tb!Comment & "") <> "" Then
230               s = tb!TestName & " : " & tb!Comment & vbCrLf
240               lblComment = lblComment & s
250           End If
260       End If

270       tb.MoveNext
280   Loop

290   If g.Rows > 2 And g.TextMatrix(1, 0) = "" Then
300       g.RemoveItem 1
310   End If

320   Exit Sub

lpanels_Click_Error:

      Dim strES As String
      Dim intEL As Integer

330   intEL = Erl
340   strES = Err.Description
350   LogError "frmAddToTests", "lpanels_Click", intEL, strES, sql

End Sub

Private Sub Treeview1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

'Set nodX = Treeview1.SelectedItem

End Sub


Private Sub Treeview1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

'If Button = vbLeftButton Then
'  Treeview1.Drag vbBeginDrag
'End If

End Sub






Public Property Let Department(ByVal sNewValue As String)

10    pDepartment = sNewValue

End Property

Public Property Let SampleID(ByVal sNewValue As String)
10    pSampleID = sNewValue
End Property

Public Property Let SampleDate(ByVal sNewValue As String)
10    pSampleDate = sNewValue
End Property
Public Property Let SampleTime(ByVal sNewValue As String)
10    pSampleTime = sNewValue
End Property

Public Property Let ClinDetails(ByVal sNewValue As String)
10    pClinDetails = sNewValue
End Property


Public Property Let sex(ByVal sNewValue As String)
10    pSex = sNewValue
End Property

Public Property Let GP(ByVal sNewValue As String)
10    pGP = sNewValue
End Property

Public Property Let Clinician(ByVal sNewValue As String)
10    pClinician = sNewValue
End Property

Public Property Let Ward(ByVal sNewValue As String)
10    pWard = sNewValue
End Property




Private Sub tv_NodeClick(ByVal Node As MSComctlLib.Node)
Dim sql As String
      Dim tb As Recordset
      Dim n As Integer
      Dim Found As Boolean
      Dim s As String

10    On Error GoTo Treeview1_NodeClick_Error

20    sql = "Select * from ExternalDefinitions where " & _
            "AnalyteName = '" & AddTicks(Node.Text) & "' AND COALESCE(InUse, 0) = 1"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    If Not tb.EOF Then
60        Found = False
70        For n = 0 To g.Rows - 1
80            If g.TextMatrix(n, 0) = Node.Text Then
90                Found = True
100               Exit For
110           End If
120       Next
130       If Not Found Then
140           s = Node.Text & vbTab & tb!SampleType & vbTab & tb!SendTo & "" & vbTab & tb!Department & ""
150           g.AddItem s


160           cmdSave.Enabled = True

170           If Trim$(tb!Comment & "") <> "" Then
180               s = Node.Text & " : " & tb!Comment & vbCrLf
190               lblComment = lblComment & s
200           End If
210       End If
220   End If

230   If g.Rows > 2 And g.TextMatrix(1, 0) = "" Then
240       g.RemoveItem 1
250   End If

260   Exit Sub

Treeview1_NodeClick_Error:

      Dim strES As String
      Dim intEL As Integer

270   intEL = Erl
280   strES = Err.Description
290   LogError "frmAddToTests", "Treeview1_NodeClick", intEL, strES, sql
End Sub
