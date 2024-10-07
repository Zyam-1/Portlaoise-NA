VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmExtOut 
   Caption         =   "NetAcquire - Outstanding External Tests "
   ClientHeight    =   8085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12510
   Icon            =   "frmExtOut.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8085
   ScaleWidth      =   12510
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Between Dates"
      Height          =   1875
      Left            =   135
      TabIndex        =   4
      Top             =   90
      Width           =   5175
      Begin VB.CommandButton cmdStart 
         Caption         =   "Start"
         Height          =   780
         Left            =   3825
         Picture         =   "frmExtOut.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   180
         Width           =   1185
      End
      Begin Threed.SSOption obetween 
         Height          =   225
         Index           =   6
         Left            =   720
         TabIndex        =   5
         Top             =   870
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Today"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Value           =   -1  'True
      End
      Begin Threed.SSOption obetween 
         Height          =   225
         Index           =   3
         Left            =   300
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1470
         Width           =   1155
         _Version        =   65536
         _ExtentX        =   2037
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Last Quarter"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin Threed.SSOption obetween 
         Height          =   225
         Index           =   4
         Left            =   1530
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1470
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Last Full Quarter"
         ForeColor       =   0
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
      Begin Threed.SSOption obetween 
         Height          =   225
         Index           =   2
         Left            =   1530
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1170
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Last Full Month"
         ForeColor       =   0
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
      Begin Threed.SSOption obetween 
         Height          =   225
         Index           =   5
         Left            =   3300
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1470
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Year to Date"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin Threed.SSOption obetween 
         Height          =   225
         Index           =   1
         Left            =   390
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1170
         Width           =   1065
         _Version        =   65536
         _ExtentX        =   1879
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Last Month"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin Threed.SSOption obetween 
         Height          =   225
         Index           =   0
         Left            =   3450
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1170
         Width           =   1065
         _Version        =   65536
         _ExtentX        =   1879
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Last Week"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   285
         Left            =   495
         TabIndex        =   12
         Top             =   315
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         _Version        =   393216
         Format          =   296747009
         CurrentDate     =   37019
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   315
         Left            =   2160
         TabIndex        =   13
         Top             =   300
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         Format          =   296747009
         CurrentDate     =   37019
      End
      Begin VB.Label Label2 
         Caption         =   "From"
         Height          =   240
         Left            =   45
         TabIndex        =   15
         Top             =   315
         Width           =   510
      End
      Begin VB.Label Label5 
         Caption         =   "To"
         Height          =   285
         Left            =   1890
         TabIndex        =   14
         Top             =   315
         Width           =   285
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   780
      Left            =   11025
      Picture         =   "frmExtOut.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   630
      Width           =   1275
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "Export to Excel"
      Height          =   780
      Left            =   9675
      Picture         =   "frmExtOut.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   630
      Width           =   1230
   End
   Begin MSFlexGridLib.MSFlexGrid grdExt 
      Height          =   5985
      Left            =   90
      TabIndex        =   0
      Top             =   2025
      Width           =   12285
      _ExtentX        =   21669
      _ExtentY        =   10557
      _Version        =   393216
      Cols            =   5
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      FormatString    =   $"frmExtOut.frx":0C28
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
      Height          =   315
      Left            =   9660
      TabIndex        =   17
      Top             =   1440
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Label Label3 
      Caption         =   "The Report is being Generated.              Please Wait."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   2430
      TabIndex        =   1
      Top             =   2475
      Width           =   6765
   End
End
Attribute VB_Name = "frmExtOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExcel_Click()

10        On Error GoTo cmdExcel_Click_Error

20        ExportFlexGrid grdExt, Me

30        Exit Sub

cmdExcel_Click_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmExtOut", "cmdExcel_Click", intEL, strES


End Sub

Private Sub cmdExit_Click()

10        Unload Me

End Sub

Private Sub cmdStart_Click()

10        On Error GoTo cmdStart_Click_Error

20        Me.Refresh

30        LoadExt

40        Exit Sub

cmdStart_Click_Error:

          Dim strES As String
          Dim intEL As Integer



50        intEL = Erl
60        strES = Err.Description
70        LogError "frmExtOut", "cmdStart_Click", intEL, strES


End Sub

Private Sub LoadExt()


          Dim tb As Recordset
          Dim sql As String
          Dim sn As Recordset
          Dim Test As String
          Dim s As String

10        On Error GoTo LoadExt_Error

20        With grdExt
30            .Rows = 2
40            .AddItem ""
50            .RemoveItem 1
60        End With

70        grdExt.Visible = False

80        Me.Refresh
          'sql = "Select * from ExtResults where retdate = '' " & _
           '      "or retdate is null and (sampleid > 30000 and sampleid < 70000) " & _
           '      " and sentdate between '" & Format(dtFrom, "dd/MMM/yyyy") & "' and '" & _
           '      Format(dtTo, "dd/MMM/yyyy") & "'"
          'QMS Ref #817830
90        sql = "Select * from ExtResults where COALESCE(retdate,'') = '' " & _
                " and CONVERT(DATE,sentdate,111)between '" & Format(dtFrom, "dd/MMM/yyyy") & "' and '" & _
                Format(dtTo, "dd/MMM/yyyy") & "'"

100       Set tb = New Recordset
110       RecOpenServer 0, tb, sql
120       Do While Not tb.EOF
130           Test = ""
140           s = tb!SampleID & vbTab
150           sql = "Select Patname from demographics where sampleid = " & tb!SampleID & ""
160           Set sn = New Recordset
170           RecOpenServer 0, sn, sql
180           If Not sn.EOF Then
190               s = s & sn!PatName & vbTab
200           Else
210               s = s & vbTab
220           End If
230           Test = tb!Analyte & ""
240           If IsNumeric(Test) Then
250               Test = eNumber2Name(tb!Analyte, "General")
260           End If
270           s = s & Test & vbTab
280           s = s & eName2SendTo(Test) & vbTab
290           s = s & Format(tb!SentDate, "dd/MMM/yyyy")
300           grdExt.AddItem s
310           tb.MoveNext
320       Loop

330       If grdExt.Rows > 2 And grdExt.TextMatrix(1, 0) = "" Then grdExt.RemoveItem 1

340       grdExt.Visible = True

350       Exit Sub

LoadExt_Error:

          Dim strES As String
          Dim intEL As Integer



360       intEL = Erl
370       strES = Err.Description
380       LogError "frmExtOut", "LoadExt", intEL, strES, sql


End Sub

Private Sub Form_Load()
10        dtFrom.Value = Now - 1
20        dtTo.Value = Now
End Sub

Private Sub oBetween_Click(Index As Integer, Value As Integer)

          Dim upto As String

10        On Error GoTo oBetween_Click_Error

20        dtFrom = BetweenDates(Index, upto)
30        dtTo = upto

40        LoadExt

50        Exit Sub

oBetween_Click_Error:

          Dim strES As String
          Dim intEL As Integer



60        intEL = Erl
70        strES = Err.Description
80        LogError "frmExtOut", "oBetween_Click", intEL, strES

End Sub
