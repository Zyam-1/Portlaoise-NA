VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGps 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - G. P. Entry"
   ClientHeight    =   9960
   ClientLeft      =   90
   ClientTop       =   390
   ClientWidth     =   18720
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
   Icon            =   "frmGps.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9960
   ScaleWidth      =   18720
   Begin VB.Frame fraGPPrinting 
      Caption         =   "DR. Name"
      Height          =   3150
      Left            =   15405
      TabIndex        =   37
      Top             =   2100
      Visible         =   0   'False
      Width           =   3180
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
         TabIndex        =   40
         Top             =   2610
         Width           =   930
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
         TabIndex        =   39
         Top             =   2610
         Width           =   930
      End
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
         ItemData        =   "frmGps.frx":030A
         Left            =   120
         List            =   "frmGps.frx":0326
         Style           =   1  'Checkbox
         TabIndex        =   38
         Top             =   315
         Width           =   2865
      End
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "Export to Excel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1100
      Left            =   17175
      Picture         =   "frmGps.frx":0396
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   555
      Width           =   1200
   End
   Begin VB.Timer tmrDown 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   17820
      Top             =   6060
   End
   Begin VB.Timer tmrUp 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   17820
      Top             =   5490
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1100
      Left            =   16995
      Picture         =   "frmGps.frx":06A0
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   6780
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton cmdMoveDown 
      Enabled         =   0   'False
      Height          =   555
      Left            =   15360
      Picture         =   "frmGps.frx":156A
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   6000
      Width           =   465
   End
   Begin VB.CommandButton cmdMoveUp 
      Enabled         =   0   'False
      Height          =   555
      Left            =   15360
      Picture         =   "frmGps.frx":2EEC
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   5430
      Width           =   465
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1100
      Left            =   16995
      Picture         =   "frmGps.frx":486E
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   8430
      Width           =   1200
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1100
      Left            =   15630
      Picture         =   "frmGps.frx":5738
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   555
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      Caption         =   "Add GP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   180
      TabIndex        =   14
      Top             =   45
      Width           =   15180
      Begin VB.TextBox txtMCNumber 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   13065
         TabIndex        =   34
         Top             =   750
         Width           =   1665
      End
      Begin VB.CommandButton cmdAddToPractice 
         Caption         =   "..."
         Height          =   315
         Left            =   8280
         TabIndex        =   30
         ToolTipText     =   "Add/Edit Practices"
         Top             =   735
         Width           =   405
      End
      Begin VB.ComboBox cmbHospital 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   810
         TabIndex        =   4
         Text            =   "cmbHospital"
         Top             =   345
         Width           =   4650
      End
      Begin VB.ComboBox cmbPractice 
         Height          =   315
         Left            =   2850
         TabIndex        =   1
         Top             =   735
         Width           =   5265
      End
      Begin VB.TextBox txtFAX 
         Height          =   285
         Left            =   11955
         TabIndex        =   11
         Top             =   1620
         Width           =   1635
      End
      Begin VB.TextBox txtPhone 
         Height          =   285
         Left            =   11955
         TabIndex        =   8
         Top             =   1290
         Width           =   1635
      End
      Begin VB.TextBox txtSurname 
         Height          =   285
         Left            =   4470
         TabIndex        =   7
         Top             =   1290
         Width           =   5115
      End
      Begin VB.TextBox txtForeName 
         Height          =   285
         Left            =   2040
         TabIndex        =   6
         Top             =   1290
         Width           =   2415
      End
      Begin VB.TextBox txtTitle 
         Height          =   285
         Left            =   810
         TabIndex        =   5
         Top             =   1290
         Width           =   1185
      End
      Begin VB.TextBox txtAddr1 
         Height          =   285
         Left            =   5250
         TabIndex        =   10
         Top             =   1620
         Width           =   4335
      End
      Begin VB.TextBox txtAddr0 
         Height          =   285
         Left            =   810
         TabIndex        =   9
         Top             =   1620
         Width           =   4425
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   13950
         Picture         =   "frmGps.frx":6602
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1200
         Width           =   780
      End
      Begin VB.TextBox txtCode 
         Height          =   285
         Left            =   810
         MaxLength       =   5
         TabIndex        =   0
         Top             =   750
         Width           =   1185
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "MC Number"
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
         Left            =   12105
         TabIndex        =   33
         Top             =   795
         Width           =   840
      End
      Begin VB.Label lblHL 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9690
         TabIndex        =   2
         Top             =   750
         Width           =   435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hospital"
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
         Left            =   150
         TabIndex        =   31
         Top             =   405
         Width           =   570
      End
      Begin VB.Label Label12 
         Caption         =   "Compiled Report"
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
         Left            =   10245
         TabIndex        =   23
         Top             =   795
         Width           =   1185
      End
      Begin VB.Label lblCompiled 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   11550
         TabIndex        =   3
         Top             =   750
         Width           =   435
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Practice"
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
         Left            =   2145
         TabIndex        =   22
         Top             =   795
         Width           =   585
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "FAX"
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
         Left            =   11625
         TabIndex        =   21
         Top             =   1680
         Width           =   300
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Phone"
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
         Left            =   11475
         TabIndex        =   20
         Top             =   1380
         Width           =   465
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Surname"
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
         Left            =   4470
         TabIndex        =   19
         Top             =   1110
         Width           =   630
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Forename"
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
         Left            =   2070
         TabIndex        =   18
         Top             =   1110
         Width           =   705
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Title"
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
         Left            =   810
         TabIndex        =   17
         Top             =   1110
         Width           =   300
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Address"
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
         Left            =   150
         TabIndex        =   16
         Top             =   1650
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Code"
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
         Left            =   345
         TabIndex        =   15
         Top             =   795
         Width           =   375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Healthlink"
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
         Left            =   8865
         TabIndex        =   32
         Top             =   795
         Width           =   705
      End
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   7425
      Left            =   180
      TabIndex        =   13
      Top             =   2160
      Width           =   15180
      _ExtentX        =   26776
      _ExtentY        =   13097
      _Version        =   393216
      Cols            =   15
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
      FormatString    =   $"frmGps.frx":7F84
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
   Begin MSComctlLib.ProgressBar pb 
      Height          =   225
      Left            =   150
      TabIndex        =   29
      Top             =   9660
      Visible         =   0   'False
      Width           =   15180
      _ExtentX        =   26776
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lblExcelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exporting..."
      Height          =   345
      Left            =   17175
      TabIndex        =   36
      Top             =   1650
      Visible         =   0   'False
      Width           =   1200
   End
End
Attribute VB_Name = "frmGps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Activated As Boolean

Dim strFullName As String

Private FireCounter As Long


Private Sub cmdExcel_Click()

10        ExportFlexGrid g, Me

End Sub

Private Sub FireDown()

          Dim n As Long
          Dim s As String
          Dim X As Long
          Dim VisibleRows As Long

10        On Error GoTo FireDown_Error

20        If g.row = g.Rows - 1 Then Exit Sub
30        n = g.row

40        VisibleRows = g.Height \ g.RowHeight(1) - 1

50        FireCounter = FireCounter + 1
60        If FireCounter > 5 Then
70            tmrDown.Interval = 100
80        End If

90        g.Visible = False

100       s = ""
110       For X = 0 To g.Cols - 1
120           s = s & g.TextMatrix(n, X) & vbTab
130       Next
140       s = Left$(s, Len(s) - 1)

150       g.RemoveItem n
160       If n < g.Rows Then
170           g.AddItem s, n + 1
180           g.row = n + 1
190       Else
200           g.AddItem s
210           g.row = g.Rows - 1
220       End If

230       For X = 0 To g.Cols - 1
240           g.Col = X
250           g.CellBackColor = vbYellow
260       Next

270       If Not g.RowIsVisible(g.row) Or g.row = g.Rows - 1 Then
280           If g.row - VisibleRows + 1 > 0 Then
290               g.TopRow = g.row - VisibleRows + 1
300           End If
310       End If

320       g.Visible = True

330       cmdSave.Visible = True

340       Exit Sub

FireDown_Error:

          Dim strES As String
          Dim intEL As Integer



350       intEL = Erl
360       strES = Err.Description
370       LogError "frmGps", "FireDown", intEL, strES


End Sub

Private Sub FireUp()

          Dim n As Long
          Dim s As String
          Dim X As Long

10        On Error GoTo FireUp_Error

20        If g.row = 1 Then Exit Sub

30        FireCounter = FireCounter + 1
40        If FireCounter > 5 Then
50            tmrUp.Interval = 100
60        End If

70        n = g.row

80        g.Visible = False

90        s = ""
100       For X = 0 To g.Cols - 1
110           s = s & g.TextMatrix(n, X) & vbTab
120       Next
130       s = Left$(s, Len(s) - 1)

140       g.RemoveItem n
150       g.AddItem s, n - 1

160       g.row = n - 1
170       For X = 0 To g.Cols - 1
180           g.Col = X
190           g.CellBackColor = vbYellow
200       Next

210       If Not g.RowIsVisible(g.row) Then
220           g.TopRow = g.row
230       End If

240       g.Visible = True

250       cmdSave.Visible = True

260       Exit Sub

FireUp_Error:

          Dim strES As String
          Dim intEL As Integer



270       intEL = Erl
280       strES = Err.Description
290       LogError "frmGps", "FireUp", intEL, strES


End Sub



Private Sub cmdAddToPractice_Click()

'frmPractice.Show 1

'FillPractices

End Sub

Private Sub cmdCancel_Click()

10        Unload Me

End Sub

Private Sub cmdExitGpPrinting_Click()
    fraGPPrinting.Visible = False
End Sub

Private Sub cmdMoveDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo cmdMoveDown_MouseDown_Error

20        FireDown

30        tmrDown.Interval = 250
40        FireCounter = 0

50        tmrDown.Enabled = True

60        Exit Sub

cmdMoveDown_MouseDown_Error:

          Dim strES As String
          Dim intEL As Integer



70        intEL = Erl
80        strES = Err.Description
90        LogError "frmGps", "cmdMoveDown_MouseDown", intEL, strES


End Sub


Private Sub cmdMoveDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo cmdMoveDown_MouseUp_Error

20        tmrDown.Enabled = False

30        Exit Sub

cmdMoveDown_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmGps", "cmdMoveDown_MouseUp", intEL, strES


End Sub


Private Sub cmdMoveUp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo cmdMoveUp_MouseDown_Error

20        FireUp

30        tmrUp.Interval = 250
40        FireCounter = 0

50        tmrUp.Enabled = True

60        Exit Sub

cmdMoveUp_MouseDown_Error:

          Dim strES As String
          Dim intEL As Integer



70        intEL = Erl
80        strES = Err.Description
90        LogError "frmGps", "cmdMoveUp_MouseDown", intEL, strES


End Sub


Private Sub cmdMoveUp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

10        On Error GoTo cmdMoveUp_MouseUp_Error

20        tmrUp.Enabled = False

30        Exit Sub

cmdMoveUp_MouseUp_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmGps", "cmdMoveUp_MouseUp", intEL, strES


End Sub


Private Sub cmdPrint_Click()

          Dim Y As Long

10        On Error GoTo cmdPrint_Click_Error

20        Printer.Print
30        Printer.Font.Name = "Courier New"
40        Printer.Font.Size = 12

50        Printer.Print "List of G. P.'s."

60        For Y = 0 To g.Rows - 1
70            g.row = Y
80            g.Col = 0
90            Printer.Print g; Tab(5);
100           g.Col = 2
110           Printer.Print g; Tab(50);
120           g.Col = 3
130           Printer.Print g
140       Next

150       Printer.EndDoc



160       Exit Sub

cmdPrint_Click_Error:

          Dim strES As String
          Dim intEL As Integer



170       intEL = Erl
180       strES = Err.Description
190       LogError "frmGps", "cmdPrint_Click", intEL, strES


End Sub

Private Sub FillG()

    Dim tb As Recordset
    Dim sql As String
    Dim s As String
    Dim HospitalCode As String

    On Error GoTo FillG_Error

    g.Visible = False
    g.Rows = 2
    g.AddItem ""
    g.RemoveItem 1

    HospitalCode = ListCodeFor("HO", cmbHospital)

    sql = "SELECT Code, " & _
          "CASE InUse WHEN 1 THEN 'Yes' ELSE 'No' END InUse, " & _
          "COALESCE(Text, '') Text, " & _
          "COALESCE(Addr0, '') Addr0, " & _
          "COALESCE(Addr1, '') Addr1, " & _
          "COALESCE(Title, '') Title, " & _
          "COALESCE(ForeName, '') ForeName, " & _
          "COALESCE(SurName, '') SurName, " & _
          "COALESCE(Phone, '') Phone, " & _
          "COALESCE(FAX, '') FAX, " & _
          "COALESCE(Practice, '') Practice, " & _
          "CASE Compiled WHEN 1 THEN 'Compiled' ELSE 'Full' END Compiled, " & _
          "CASE HealthLink WHEN 1 THEN 'Yes' ELSE 'No' END HealthLink, " & _
          "COALESCE(MCNumber, '') MCNumber " & _
          "FROM GPs WHERE " & _
          "HospitalCode = '" & HospitalCode & "' " & _
          "ORDER BY LISTORDER"
    Set tb = New Recordset
    RecOpenServer 0, tb, sql

    Do While Not tb.EOF
        With tb
            s = !Code & vbTab & _
                !InUse & vbTab & _
                !Text & vbTab & _
                !Addr0 & vbTab & _
                !Addr1 & vbTab & _
                !Title & vbTab & _
                !ForeName & vbTab & _
                !SurName & vbTab & _
                !Phone & vbTab & _
                !FAX & vbTab & _
                !practice & vbTab & _
                !compiled & vbTab & _
                !HealthLink & vbTab & _
                !MCNumber & vbTab & _
                "Click To View"
            g.AddItem s
'            If CheckDisablePrinting(!Code) Then
'                g.RowSel = g.row
'                g.ColSel = 14
'                g.CellBackColor = vbRed
'            End If
        End With
        tb.MoveNext
    Loop
    If g.Rows > 2 Then g.RemoveItem 1
    g.Visible = True
    
    Exit Sub

FillG_Error:

    Dim strES As String
    Dim intEL As Integer

    intEL = Erl
    strES = Err.Description
    LogError "frmGps", "FillG", intEL, strES, sql
    g.Visible = True

End Sub

Private Sub cmdSave_Click()

          Dim HospitalCode As String
          Dim Y As Long
          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo cmdSave_Click_Error

20        HospitalCode = ListCodeFor("HO", cmbHospital)

30        pb.Max = g.Rows - 1
40        pb.Visible = True
50        cmdSave.Caption = "Saving..."

60        For Y = 1 To g.Rows - 1
70            pb = Y
80            sql = "SELECT * from GPs WHERE " & _
                    "Code = '" & g.TextMatrix(Y, 0) & "' " & _
                    "and HospitalCode = '" & HospitalCode & "'"
90            Set tb = New Recordset
100           RecOpenClient 0, tb, sql
110           If tb.EOF Then
120               tb.AddNew
130           End If
140           With tb
150               !Code = g.TextMatrix(Y, 0)
160               If g.TextMatrix(Y, 1) = "Yes" Then !InUse = 1 Else !InUse = 0
170               !Text = g.TextMatrix(Y, 2)
180               !Addr0 = initial2upper(g.TextMatrix(Y, 3))
190               !Addr1 = initial2upper(g.TextMatrix(Y, 4))
200               !Title = initial2upper(g.TextMatrix(Y, 5))
210               !ForeName = initial2upper(g.TextMatrix(Y, 6))
220               !SurName = initial2upper(g.TextMatrix(Y, 7))
230               !Phone = g.TextMatrix(Y, 8)
240               !FAX = g.TextMatrix(Y, 9)
250               !practice = g.TextMatrix(Y, 10)
260               !compiled = g.TextMatrix(Y, 11) = "Compiled"
270               !HospitalCode = HospitalCode
280               !ListOrder = Y
290               If g.TextMatrix(Y, 12) = "Yes" Then !HealthLink = 1 Else !HealthLink = False
300               !MCNumber = g.TextMatrix(Y, 13)
310               .Update
320           End With
330       Next

340       pb.Visible = False
350       cmdSave.Visible = False
360       cmdSave.Caption = "Save"

370       Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

380       intEL = Erl
390       strES = Err.Description
400       LogError "frmGps", "cmdsave_Click", intEL, strES, sql

End Sub

Private Sub cmbHospital_Click()

10        On Error GoTo cmbHospital_Click_Error

20        FillPractices

30        FillG

40        Exit Sub

cmbHospital_Click_Error:

          Dim strES As String
          Dim intEL As Integer



50        intEL = Erl
60        strES = Err.Description
70        LogError "frmGps", "cmbHospital_Click", intEL, strES


End Sub

Private Sub cmbHospital_KeyPress(KeyAscii As Integer)

10        On Error GoTo cmbHospital_KeyPress_Error

20        KeyAscii = 0

30        Exit Sub

cmbHospital_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmGps", "cmbHospital_KeyPress", intEL, strES


End Sub


Private Sub cmdadd_Click()

          Dim strCode As String
          Dim strSurName As String
          Dim s As String
          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo cmdadd_Click_Error

20        strCode = Trim$(UCase$(txtCode))
30        If strCode = "" Then
40            iMsg "Enter Code", vbCritical
50            Exit Sub
60        End If

70        If Len(txtCode) <> 3 Then
80            iMsg "Code Length must be 3 Chars.", vbCritical
90            Exit Sub
100       End If

110       strSurName = Trim$(txtSurname)
120       If strSurName = "" Then
130           iMsg "Enter Surname", vbCritical
140           Exit Sub
150       End If

160       If txtCode.Locked = False Then
170           sql = "SELECT * from gps WHERE code = '" & strCode & "' and hospitalcode = '" & Left(HospName(0), 1) & "'"
180           Set tb = New Recordset
190           RecOpenServer 0, tb, sql
200           If Not tb.EOF Then
210               iMsg "Code Already Exists!"
220               txtCode = ""
230               Exit Sub
240           End If
250       End If

260       s = strCode & vbTab & _
              "Yes" & vbTab & _
              strFullName & vbTab & _
              txtAddr0 & vbTab & _
              txtAddr1 & vbTab & _
              txtTitle & vbTab & _
              txtForeName & vbTab & _
              txtSurname & vbTab & _
              txtPhone & vbTab & _
              txtFAX & vbTab & _
              cmbPractice & vbTab & _
              IIf(lblCompiled = "Yes", "Compiled", "Full") & vbTab & _
              lblHL & vbTab & _
              txtMCNumber
270       g.AddItem s

280       txtCode = ""
290       txtAddr0 = ""
300       txtAddr1 = ""
310       txtTitle = ""
320       txtForeName = ""
330       txtSurname = ""
340       txtPhone = ""
350       txtFAX = ""
360       cmbPractice = ""
370       lblCompiled = "No"
380       lblHL = "No"
390       txtMCNumber = ""
400       txtCode.Locked = False

410       cmdSave.Visible = True

420       Exit Sub

cmdadd_Click_Error:

          Dim strES As String
          Dim intEL As Integer

430       intEL = Erl
440       strES = Err.Description
450       LogError "frmGps", "cmdAdd_Click", intEL, strES, sql

End Sub

Private Sub cmbPractice_Click()


'sql = "SELECT * from Practices WHERE " & _
  '      "Text = '" & cmbPractice.Text & "'"
'RecOpenServer 0, tb, sql
'If tb.EOF Then
'  txtFAX = ""
'Else
'  txtFAX = tb!FAX & ""
'End If

End Sub


Private Sub cmdSaveGpPrinting_Click()

    Dim Y As Long
    Dim sql As String
    Dim tb As Recordset
    Dim Dept As String
    
    For Y = 0 To LstGpPrinting.ListCount - 1
        Dept = LstGpPrinting.List(Y)
        
        If LstGpPrinting.Selected(Y) = True Then
            
            sql = "IF EXISTS (SELECT * FROM DisablePrinting WHERE " & _
                        "           Department = '" & LstGpPrinting.List(Y) & "' " & _
                        "           AND GPCode = '" & g.TextMatrix(g.row, 0) & "' )" & _
                        "    UPDATE DisablePrinting " & _
                        "    SET Department = '" & LstGpPrinting.List(Y) & "', " & _
                        "    GPCode = '" & g.TextMatrix(g.row, 0) & "', " & _
                        "    GPName = '" & g.TextMatrix(g.row, 2) & "', " & _
                        "    Type = 'GP', " & _
                        "    Disable = '1' " & _
                        "    WHERE Department = '" & LstGpPrinting.List(Y) & "' " & _
                        "    AND GPCode = '" & g.TextMatrix(g.row, 0) & "' " & _
                        "ELSE " & _
                        "    INSERT INTO DisablePrinting " & _
                    "    (GPCode, GPName, Type, Disable, Department) "
                   sql = sql & _
                        "    VALUES ( " & _
                        "    '" & g.TextMatrix(g.row, 0) & "', " & _
                        "    '" & g.TextMatrix(g.row, 2) & "', " & _
                        "    'GP', " & _
                        "    '1', " & _
                        "    '" & LstGpPrinting.List(Y) & "'" & _
                        "     )"
                  Cnxn(0).Execute sql
        Else
            sql = "DELETE FROM DisablePrinting WHERE " & _
                    "Department = '" & LstGpPrinting.List(Y) & "' " & _
                    "AND GPCode = '" & g.TextMatrix(g.row, 0) & "'"
                    Cnxn(0).Execute sql
        End If
    Next
    
     fraGPPrinting.Visible = False
     'MsgBox "Updated!", vbOKOnly, "Disable Printing"
     

    
    

End Sub

Private Sub Form_Activate()

10        On Error GoTo Form_Activate_Error

20        If Activated Then
30            Exit Sub
40        End If

50        Activated = True

60        FillG

70        Exit Sub

Form_Activate_Error:

          Dim strES As String
          Dim intEL As Integer

80        intEL = Erl
90        strES = Err.Description
100       LogError "frmGps", "Form_Activate", intEL, strES

End Sub

Private Sub Form_Load()

          Dim tb As Recordset
          Dim sql As String
          Dim n As Long

10        On Error GoTo Form_Load_Error

20        cmbHospital.Clear

30        sql = "SELECT * from Lists WHERE " & _
                "ListType = 'HO' " & _
                "order by ListOrder"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        Do While Not tb.EOF
70            cmbHospital.AddItem tb!Text & ""
80            tb.MoveNext
90        Loop
100       For n = 0 To cmbHospital.ListCount - 1
110           If UCase(cmbHospital.List(n)) = UCase(HospName(0)) Then
120               cmbHospital = Format(HospName(0), vbProperCase)
130           End If
140       Next

150       FillPractices
          'Trevor 18/11/15
160       g.ColWidth(14) = 0
170       If GetOptionSetting("DisablePrinting", 0) = True Then
180           g.ColWidth(14) = 1200
190       End If
          
200       Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

210       intEL = Erl
220       strES = Err.Description
230       LogError "frmGps", "Form_Load", intEL, strES, sql


End Sub

Private Sub FillPractices()


End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

10        On Error GoTo Form_QueryUnload_Error

20        If cmdSave.Visible Then
30            If iMsg("Cancel without saving?", vbQuestion + vbYesNo) = vbNo Then
40                Cancel = True
50            End If
60        End If

70        Exit Sub

Form_QueryUnload_Error:

          Dim strES As String
          Dim intEL As Integer



80        intEL = Erl
90        strES = Err.Description
100       LogError "frmGps", "Form_QueryUnload", intEL, strES


End Sub


Private Sub Form_Unload(Cancel As Integer)

10        Activated = False

End Sub


Private Sub g_Click()

          Static SortOrder As Boolean
          Dim X As Long
          Dim Y As Long
          Dim ySave As Long

10        On Error GoTo g_Click_Error

20        ySave = g.row

30        If g.MouseRow = 0 Then
40            If SortOrder Then
50                g.Sort = flexSortGenericAscending
60            Else
70                g.Sort = flexSortGenericDescending
80            End If
90            SortOrder = Not SortOrder
100           cmdMoveUp.Enabled = False
110           cmdMoveDown.Enabled = False
120           cmdSave.Visible = True
130           Exit Sub
140       End If

150       If g.Col = 11 Then
160           g = IIf(g = "Full", "Compiled", "Full")
170           cmdSave.Visible = True
180           Exit Sub
190       End If

200       If g.Col = 1 Or g.Col = 12 Then
210           g = IIf(g = "No", "Yes", "No")
220           cmdSave.Visible = True
230           Exit Sub
240       End If

250       If g.Col = 13 Then
260           g = iBOX("Enter MC Number ", , g, False)
270           cmdSave.Visible = True
280           Exit Sub
290       End If

300       If g.Col = 0 Then
310           g.Enabled = False
320           If iMsg("Edit/Remove this line?", vbQuestion + vbYesNo) = vbYes Then

330               txtCode = g.TextMatrix(g.row, 0)
340               txtCode.Locked = True
350               txtTitle = g.TextMatrix(g.row, 5)
360               txtForeName = g.TextMatrix(g.row, 6)
370               txtSurname = g.TextMatrix(g.row, 7)
380               txtPhone = g.TextMatrix(g.row, 8)
390               txtFAX = g.TextMatrix(g.row, 9)
400               cmbPractice = g.TextMatrix(g.row, 10)
410               txtAddr0 = g.TextMatrix(g.row, 3)
420               txtAddr1 = g.TextMatrix(g.row, 4)
430               lblCompiled = IIf(g.TextMatrix(g.row, 11) = "Compiled", "Yes", "No")
440               lblHL = g.TextMatrix(g.row, 12)
450               g.RemoveItem g.row
460               cmdSave.Visible = True
470           End If
480           g.Enabled = True
490           Exit Sub
500       End If

510       If g.Col = 9 Then
520           g = iBOX("Enter Fax Number ", , g, False)
530           cmdSave.Visible = True
540           Exit Sub
550       End If

560       g.Visible = False
570       g.Col = 0
580       For Y = 1 To g.Rows - 1
590           g.row = Y
600           If g.CellBackColor = vbYellow Then
610               For X = 0 To g.Cols - 1
620                   g.Col = X
630                   g.CellBackColor = 0
640               Next
650               Exit For
660           End If
670       Next
680       g.row = ySave
690       g.Visible = True

700       For X = 0 To g.Cols - 1
710           g.Col = X
720           g.CellBackColor = vbYellow
730       Next
          
740       If g.Col = 14 Then
750               fraGPPrinting.Visible = True
760               fraGPPrinting.Caption = g.TextMatrix(g.row, 2)
770               FillLstGpPrinting
780       End If
          
790       cmdMoveUp.Enabled = True
800       cmdMoveDown.Enabled = True

810       Exit Sub

g_Click_Error:

          Dim strES As String
          Dim intEL As Integer

820       intEL = Erl
830       strES = Err.Description
840       LogError "frmGps", "g_Click", intEL, strES

End Sub
Private Sub FillLstGpPrinting()

    Dim tb As Recordset
    Dim sql As String
    'Dim dept As String
    Dim Y As Integer
       
    sql = "SELECT * FROM DisablePrinting WHERE " & _
            " GPCode = '" & g.TextMatrix(g.row, 0) & "'"
    Set tb = New Recordset
    RecOpenServer 0, tb, sql
    
    For Y = 0 To LstGpPrinting.ListCount - 1
        'dept = LstGpPrinting.List(Y)
        LstGpPrinting.Selected(Y) = False
    Next Y
    
    Do While Not tb.EOF
        For Y = 0 To LstGpPrinting.ListCount - 1
            If LstGpPrinting.List(Y) = tb!Department Then
                LstGpPrinting.Selected(Y) = True
            End If
        Next
        tb.MoveNext
    Loop

    
End Sub

Private Sub lblCompiled_Click()

10        On Error GoTo lblCompiled_Click_Error

20        lblCompiled = IIf(lblCompiled = "Yes", "No", "Yes")

30        Exit Sub

lblCompiled_Click_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmGps", "lblCompiled_Click", intEL, strES


End Sub

Private Sub lblHL_Click()

10        On Error GoTo lblHL_Click_Error

20        lblHL = IIf(lblHL = "Yes", "No", "Yes")

30        Exit Sub

lblHL_Click_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmGps", "lblHL_Click", intEL, strES


End Sub



Private Sub tmrDown_Timer()

10        On Error GoTo tmrDown_Timer_Error

20        FireDown

30        Exit Sub

tmrDown_Timer_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmGps", "tmrDown_Timer", intEL, strES


End Sub

Private Sub tmrUp_Timer()

10        On Error GoTo tmrUp_Timer_Error

20        FireUp

30        Exit Sub

tmrUp_Timer_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmGps", "tmrUp_Timer", intEL, strES


End Sub


Private Sub txtCode_LostFocus()

10        On Error GoTo txtCode_LostFocus_Error

20        txtCode = UCase$(Trim$(txtCode))

30        Exit Sub

txtCode_LostFocus_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmGps", "txtCode_LostFocus", intEL, strES


End Sub

Private Sub txtForeName_Change()

10        On Error GoTo txtForeName_Change_Error

20        txtForeName = Trim$(txtForeName)

30        strFullName = txtTitle & " " & txtForeName & " " & txtSurname

40        Exit Sub

txtForeName_Change_Error:

          Dim strES As String
          Dim intEL As Integer



50        intEL = Erl
60        strES = Err.Description
70        LogError "frmGps", "txtForeName_Change", intEL, strES


End Sub

Private Sub txtSurname_Change()

10        On Error GoTo txtSurname_Change_Error

20        txtSurname = Trim$(txtSurname)

30        strFullName = txtTitle & " " & txtForeName & " " & txtSurname

40        Exit Sub

txtSurname_Change_Error:

          Dim strES As String
          Dim intEL As Integer



50        intEL = Erl
60        strES = Err.Description
70        LogError "frmGps", "txtSurname_Change", intEL, strES


End Sub


Private Sub txtTitle_Change()

10        On Error GoTo txtTitle_Change_Error

20        txtTitle = Trim$(txtTitle)

30        strFullName = txtTitle & " " & txtForeName & " " & txtSurname

40        Exit Sub

txtTitle_Change_Error:

          Dim strES As String
          Dim intEL As Integer



50        intEL = Erl
60        strES = Err.Description
70        LogError "frmGps", "txtTitle_Change", intEL, strES


End Sub

Private Function CheckDisablePrinting(ByVal GPCode As String) As Boolean

    Dim strCode As String
    Dim strSurName As String
    Dim s As String
    Dim sql As String
    Dim tb As Recordset
    
    CheckDisablePrinting = False
    
    sql = "SELECT * from DisablePrinting WHERE " & _
    "GPCode = '" & GPCode & "'"
        Set tb = New Recordset
           RecOpenClient 0, tb, sql
           If Not tb.EOF Then
               CheckDisablePrinting = True
           End If
    
End Function

