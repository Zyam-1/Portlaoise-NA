VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmRFTBlood 
   Caption         =   "Netacquire - Report Viewer"
   ClientHeight    =   10185
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14955
   LinkTopic       =   "Form1"
   ScaleHeight     =   17.965
   ScaleMode       =   7  'Centimeter
   ScaleWidth      =   26.379
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   825
      Left            =   13635
      Picture         =   "frmRFT.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9225
      Width           =   1275
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8880
      Left            =   90
      TabIndex        =   0
      Top             =   1215
      Width           =   13470
      _ExtentX        =   23760
      _ExtentY        =   15663
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Page 1"
      TabPicture(0)   =   "frmRFT.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "rtfRep"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Page 2"
      TabPicture(1)   =   "frmRFT.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "rtfRep1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin RichTextLib.RichTextBox rtfRep 
         Height          =   8295
         Left            =   90
         TabIndex        =   1
         Top             =   405
         Width           =   13080
         _ExtentX        =   23072
         _ExtentY        =   14631
         _Version        =   393217
         ScrollBars      =   2
         TextRTF         =   $"frmRFT.frx":0342
      End
      Begin RichTextLib.RichTextBox rtfRep1 
         Height          =   8295
         Left            =   -74865
         TabIndex        =   2
         Top             =   450
         Width           =   13080
         _ExtentX        =   23072
         _ExtentY        =   14631
         _Version        =   393217
         ScrollBars      =   2
         TextRTF         =   $"frmRFT.frx":03C4
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdRep 
      Height          =   1215
      Left            =   90
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   45
      Width           =   14520
      _ExtentX        =   25612
      _ExtentY        =   2143
      _Version        =   393216
      Cols            =   8
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
      FormatString    =   $"frmRFT.frx":0446
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
   Begin VB.Label lblNoRep 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   13815
      TabIndex        =   6
      Top             =   2295
      Width           =   690
   End
   Begin VB.Label Label1 
      Caption         =   "No. of Reports"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   13680
      TabIndex        =   5
      Top             =   1620
      Width           =   915
   End
End
Attribute VB_Name = "frmRFTBlood"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mSampleID As String
Private mDept As String


Private Sub TabStrip1_Click()


    rtfRep.SelStart = 0
    rtfRep.SelPrint Printer.hDC

End Sub

Private Sub cmdExit_Click()

Unload Me

End Sub

Private Sub Form_Load()
Dim tb As Recordset
Dim sql As String
Dim s As String


If mDept = "B" Then
  sql = "select * from reports where sampleid = '" & mSampleID & "' and dept = '" & mDept & "' or sampleid = '" & mSampleID & "' and dept = 'M' or sampleid = '" & mSampleID & "' and dept = 'R' order by printtime desc"
Else
  sql = "select * from reports where sampleid = '" & mSampleID & "' and dept = '" & mDept & "' or sampleid = '" & mSampleID & "' and dept = 'M' order by printtime desc"
End If
Set tb = New Recordset
RecOpenServer 0, tb, sql
If tb.EOF Then
  grdRep.AddItem mSampleID & vbTab & vbTab & mDept & vbTab & vbTab & "No Reports"
  lblNoRep = "0"
Else
  Do While Not tb.EOF
      s = tb!SampleID & vbTab & Trim(tb!Name) & vbTab & Trim(tb!Dept)
      s = s & vbTab & Format(tb!printtime, "dd/MMM/yyyy hh:mm:ss")
      s = s & vbTab & Trim(tb!repno) & vbTab & Trim(tb!Initiator) & vbTab
      If Trim(tb!pagetwo) & "" <> "" Then s = s & "2" Else s = s & "1"
      s = s & vbTab & tb!Printer
      grdRep.AddItem s
      tb.MoveNext
  Loop
  lblNoRep = grdRep.Rows - 2
End If


grdRep.RemoveItem 1


End Sub

Private Sub grdRep_Click()
Dim tb As Recordset
Dim sql As String


rtfRep = ""
rtfRep1 = ""
rtfRep.SelText = ""
rtfRep1.SelText = ""
SSTab1.TabVisible(1) = False

sql = "select * from reports where repno = '" & grdRep.TextMatrix(grdRep.RowSel, 4) & "'"
Set tb = New Recordset
RecOpenServer 0, tb, sql
Do While Not tb.EOF
  If Trim(tb!pageone & "") <> "" Then
    rtfRep.SelText = Trim(tb!pageone)
  End If
  If Trim(tb!pagetwo & "") <> "" Then
    SSTab1.TabVisible(1) = True
    rtfRep1.SelText = Trim(tb!pagetwo)
  End If
  tb.MoveNext
Loop


End Sub

Public Property Let SampleID(ByVal SID As String)

mSampleID = SID

End Property

Public Property Let Dept(ByVal Dep As String)

mDept = Dep

End Property

