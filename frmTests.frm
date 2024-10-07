VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmTests 
   Appearance      =   0  'Flat
   BackColor       =   &H80000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Biochemistry Test Count"
   ClientHeight    =   6810
   ClientLeft      =   750
   ClientTop       =   1020
   ClientWidth     =   8955
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
   Icon            =   "frmTests.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6810
   ScaleWidth      =   8955
   Begin VB.PictureBox SSPanel1 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      ScaleHeight     =   2115
      ScaleWidth      =   5295
      TabIndex        =   5
      Top             =   150
      Width           =   5355
      Begin MSComCtl2.DTPicker calFrom 
         Height          =   315
         Left            =   270
         TabIndex        =   17
         Top             =   450
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         Format          =   20316161
         CurrentDate     =   36942
      End
      Begin MSComCtl2.DTPicker calTo 
         Height          =   315
         Left            =   1950
         TabIndex        =   14
         Top             =   450
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         Format          =   20316161
         CurrentDate     =   36942
      End
      Begin VB.CommandButton bReCalc 
         Caption         =   "Calculate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   3870
         Picture         =   "frmTests.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   480
         Width           =   1155
      End
      Begin VB.OptionButton oBetween 
         Alignment       =   1  'Right Justify
         Caption         =   "Today"
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
         Index           =   6
         Left            =   1050
         TabIndex        =   12
         Top             =   900
         Width           =   795
      End
      Begin VB.OptionButton oBetween 
         Caption         =   "Year To Date"
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
         Index           =   5
         Left            =   2010
         TabIndex        =   11
         Top             =   1410
         Width           =   1305
      End
      Begin VB.OptionButton oBetween 
         Caption         =   "Last Full Quarter"
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
         Index           =   4
         Left            =   2010
         TabIndex        =   10
         Top             =   1140
         Width           =   1515
      End
      Begin VB.OptionButton oBetween 
         Caption         =   "Last Quarter"
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
         Index           =   3
         Left            =   2010
         TabIndex        =   9
         Top             =   900
         Width           =   1185
      End
      Begin VB.OptionButton oBetween 
         Alignment       =   1  'Right Justify
         Caption         =   "Last Full Month"
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
         Index           =   2
         Left            =   450
         TabIndex        =   8
         Top             =   1680
         Width           =   1395
      End
      Begin VB.OptionButton oBetween 
         Alignment       =   1  'Right Justify
         Caption         =   "Last Month"
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
         Index           =   1
         Left            =   690
         TabIndex        =   7
         Top             =   1410
         Width           =   1155
      End
      Begin VB.OptionButton oBetween 
         Alignment       =   1  'Right Justify
         Caption         =   "Last Week"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   750
         TabIndex        =   6
         Top             =   1140
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "To"
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
         Left            =   2010
         TabIndex        =   16
         Top             =   240
         Width           =   195
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "From"
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
         Left            =   300
         TabIndex        =   15
         Top             =   240
         Width           =   345
      End
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   165
      Left            =   90
      TabIndex        =   4
      Top             =   2340
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   6525
      Left            =   5550
      TabIndex        =   3
      Top             =   180
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   11509
      _Version        =   393216
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      FormatString    =   "<Analyte              |<Count              "
   End
   Begin VB.CommandButton bprint 
      Appearance      =   0  'Flat
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
      Height          =   825
      Left            =   1215
      Picture         =   "frmTests.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5445
      Width           =   1245
   End
   Begin VB.ListBox List1 
      Height          =   4815
      IntegralHeight  =   0   'False
      Left            =   9045
      TabIndex        =   1
      Top             =   135
      Width           =   975
   End
   Begin VB.CommandButton bcancel 
      Appearance      =   0  'Flat
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
      Height          =   825
      Left            =   3060
      Picture         =   "frmTests.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5445
      Width           =   1245
   End
End
Attribute VB_Name = "frmTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub obetween_Click(Index As Integer)

Dim upto As String

calFrom = BetweenDates(Index, upto)
calTo = upto

End Sub

Private Sub bCancel_Click()

Unload Me

End Sub

Private Sub bPrint_Click()

Printer.Print
Printer.Print
Printer.Print
Printer.Print
Printer.Print "Total Tests - "; Format$(calFrom, "dd/mmm/yyyy"); " to "; Format$(calTo, "dd/mmm/yyyy")
Printer.Print
Printer.Print
G.Col = 0
G.Row = 0
G.ColSel = 1
G.RowSel = G.Rows - 1
Printer.Print G.Clip
Printer.EndDoc

End Sub

Private Sub FillList()
Dim tb As New Recordset
Dim SQL As String

List1.Clear

SQL = "SELECT distinct ShortName, PrintPriority from BioTestDefinitions " & _
      " WHERE inuse = 1 order by PrintPriority asc"
Set tb = New Recordset
RecOpenServer 0, tb, SQL

Do While Not tb.EOF
  List1.AddItem tb!ShortName & ""
  tb.MoveNext
Loop

End Sub

Private Sub Form_Load()

calFrom = Format$(Now, "dd/MMM/yyyy")
calTo = Format$(Now, "dd/MMM/yyyy")

FillList

End Sub

Private Sub brecalc_Click()
Dim tb As New Recordset
Dim SQL As String
Dim n As Long
Dim fromdate As String
Dim ToDate As String
Dim s As String
Dim Tot As Long

fromdate = Format$(calFrom, "dd/MMM/yyyy") & " 00:00:00"
ToDate = Format$(calTo, "dd/MMM/yyyy") & " 23:59:59"

G.Rows = 2
G.AddItem ""
G.RemoveItem 1

pb.max = List1.ListCount
pb.Visible = True

For n = 0 To List1.ListCount - 1
  List1.Selected(n) = True
  List1.Refresh
  pb = List1.ListIndex
  SQL = "SELECT count(distinct sampleid) as tot " & _
        "from bioresults as r " & _
        "WHERE r.runtime between '" & fromdate & "' and '" & ToDate & "' " & _
        "and R.code = '" & CodeForShortName(List1.List(n)) & "' "
  Set tb = New Recordset
  RecOpenServer 0, tb, SQL
  If tb!Tot <> 0 Then
    s = List1.List(n) & vbTab & _
        Format$(tb!Tot)
    G.AddItem s
    G.Refresh
  End If
Next
pb.Visible = False

If G.Rows > 2 Then
  G.RemoveItem 1
End If

For n = 1 To G.Rows - 1
  Tot = Tot + Val(G.TextMatrix(n, 1))
Next

G.AddItem ""
G.AddItem "Total" & vbTab & Tot

End Sub
