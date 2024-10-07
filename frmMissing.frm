VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.ocx"
Begin VB.Form frmMissing 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Search for Missing Records"
   ClientHeight    =   6300
   ClientLeft      =   1605
   ClientTop       =   1320
   ClientWidth     =   7455
   Icon            =   "frmMissing.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.ProgressBar pb 
      Height          =   195
      Left            =   330
      TabIndex        =   13
      Top             =   1560
      Visible         =   0   'False
      Width           =   4605
      _ExtentX        =   8123
      _ExtentY        =   344
      _Version        =   327682
      Appearance      =   1
   End
   Begin MSFlexGridLib.MSFlexGrid gIncomp 
      Height          =   3645
      Left            =   5220
      TabIndex        =   11
      Top             =   2160
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   6429
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
      FormatString    =   "                        "
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   3645
      Left            =   3600
      TabIndex        =   9
      Top             =   2190
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   6429
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
   End
   Begin VB.TextBox tInfo 
      BackColor       =   &H80000014&
      Height          =   4005
      Left            =   330
      MaxLength       =   30000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   1860
      Width           =   3195
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   330
      TabIndex        =   1
      Top             =   330
      Width           =   4605
      Begin VB.CommandButton bStart 
         Caption         =   "&Start"
         Height          =   750
         Left            =   3180
         Picture         =   "frmMissing.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   255
         Width           =   1245
      End
      Begin VB.TextBox tYear 
         Height          =   285
         Left            =   600
         MaxLength       =   4
         TabIndex        =   3
         Text            =   "2000"
         Top             =   330
         Width           =   525
      End
      Begin ComCtl2.UpDown udYear 
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   630
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   344
         _Version        =   327681
         Value           =   2000
         BuddyControl    =   "tYear"
         BuddyDispid     =   196615
         OrigLeft        =   4920
         OrigTop         =   3630
         OrigRight       =   5160
         OrigBottom      =   4395
         Max             =   2010
         Min             =   1997
         Orientation     =   1
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin Threed.SSOption oDept 
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   5
         Top             =   600
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Histology"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   -1  'True
      End
      Begin Threed.SSOption oDept 
         Height          =   345
         Index           =   0
         Left            =   1680
         TabIndex        =   6
         Top             =   270
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   609
         _StockProps     =   78
         Caption         =   "Cytology"
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Year"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   390
         Width           =   330
      End
   End
   Begin VB.CommandButton bCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   705
      Left            =   5130
      Picture         =   "frmMissing.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   630
      Width           =   1245
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Incomplete"
      Height          =   285
      Left            =   5265
      TabIndex        =   12
      Top             =   1890
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Not Valid"
      Height          =   285
      Left            =   3600
      TabIndex        =   10
      Top             =   1890
      Width           =   1530
   End
End
Attribute VB_Name = "frmMissing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub bcancel_Click()

10        Unload Me

End Sub


Private Sub bStart_Click()

          Dim sn As Recordset
          Dim tb As Recordset
          Dim sql As String
          Dim Dept As String
          Dim n As Long
          Dim First As String
          Dim Second As String
          Dim DeptNum As Long
          Dim rs As Recordset
          Dim Yadd As Long

10        On Error GoTo bStart_Click_Error

20        tInfo = "Searching...." & vbCrLf
30        tInfo.Refresh

40        With g
50            .Rows = 1
60            .AddItem ""
70            .RemoveItem 0
80            .ColWidth(0) = 1250
90            .Visible = False
100       End With

110       With gIncomp
120           .Rows = 1
130           .AddItem ""
140           .RemoveItem 0
150           .ColWidth(0) = 1250
160           .Visible = False
170       End With

180       Yadd = Val(Swap_Year(Trim(tYear))) * 1000

190       If oDept(0) Then
200           Dept = "C"
210           DeptNum = 40000000
220       Else
230           Dept = "H"
240           DeptNum = 30000000
250       End If

260       sql = "SELECT sampleid, hyear from demographics WHERE " & _
                "hyear like '" & tYear & "' and sampleid > " & DeptNum & " and sampleid < " & (DeptNum + 9999999) & ""
270       Set sn = New Recordset
280       RecOpenClient 0, sn, sql

290       If Not sn.EOF Then
300           sn.MoveLast
310           tInfo = tInfo & "Total Records found = " & sn.RecordCount & vbCrLf
320           pb.Max = sn.RecordCount
330           tInfo.Refresh
340           sn.MoveFirst
350           Do While Not sn.EOF
360               g.AddItem sn!SampleID
370               sn.MoveNext
380           Loop
390           g.AddItem ""

400           pb.Visible = True

410           g.Sort = flexSortNumericAscending
420           For n = 2 To g.Rows - 2
430               pb = n
440               g.Row = n
450               First = g
460               g.Row = n + 1
470               Second = g
480               If Trim(Second) <> "" Then
490                   If Val(Left(First, 8)) = Val(Left(Second, 8)) Then
                          'duplicate
500                       tInfo = tInfo & "Erroneous entries: " & (Left(First, 8) - (DeptNum + Yadd)) & " and " & (Left(Second, 8) - (DeptNum + Yadd)) & vbCrLf
510                   ElseIf Val(Left(First, 8)) + 1 <> Val(Left(Second, 8)) Then
520                       tInfo = tInfo & "Missing between " & (Left(First, 8) - (DeptNum + Yadd)) & " and " & (Left(Second, 8) - (DeptNum + Yadd)) & vbCrLf
530                   End If
540               End If
550           Next
560       End If

570       pb.Max = g.Rows

580       For n = g.Rows - 1 To 2 Step -1
590           pb = pb.Max - n
600           g.Row = n
610           If oDept(0) Then
620               sql = "SELECT * FROM CytoResults WHERE " & _
                        "sampleid = '" & Left(g, 8) & "' " & _
                        "AND hYear = '" & tYear & "'"
630           Else
640               sql = "SELECT * FROM Historesults WHERE " & _
                        "SampleID = '" & Left(g, 8) & "' " & _
                        "AND hyear = '" & tYear & "'"
650           End If
660           Set tb = New Recordset
670           RecOpenServer 0, tb, sql
680           If Not tb.EOF Then
690               If oDept(0) Then
700                   If Trim(tb!cytoreport & "") = "" Then
710                       gIncomp.AddItem tYear & "/" & (Left(g, 8) - (DeptNum + Yadd)) & "C"
720                   End If
730                   sql = "SELECT cytovalid from demographics WHERE sampleid = '" & Left(g, 8) & "' and hyear = '" & tYear & "'"
740                   Set rs = New Recordset
750                   RecOpenServer 0, rs, sql
760                   If Not (Trim(tb!cytoreport & "") <> "" And Not rs!cytovalid) Then
770                       g.RemoveItem n
780                   Else
790                       g = tYear & "/" & (Left(g, 8) - (DeptNum + Yadd)) & "C"
800                   End If
810               Else
820                   If Trim(tb!historeport & "") = "" Then
830                       gIncomp.AddItem tYear & "/" & (Left(g, 8) - (DeptNum + Yadd)) & "H"
840                   End If
850                   sql = "SELECT coalesce(histovalid,0) as Valid from demographics WHERE sampleid = '" & Left(g, 8) & "' and hyear = '" & tYear & "'"
860                   Set rs = New Recordset
870                   RecOpenServer 0, rs, sql
880                   If rs.EOF Then
890                       g = tYear & "/" & (Left(g, 8) - (DeptNum + Yadd)) & "H"
900                   Else
910                       If Trim(tb!historeport & "") <> "" And rs!Valid Then
920                           g.RemoveItem n
930                       Else
940                           g = tYear & "/" & (Left(g, 8) - (DeptNum + Yadd)) & "H"
950                       End If
960                   End If
970               End If
980           End If
990       Next
1000      gIncomp.Sort = flexSortNumericAscending

1010      g.Visible = True
1020      gIncomp.Visible = True
1030      pb.Visible = False


1040      If g.Rows > 2 And g.TextMatrix(0, 0) = "" Then g.RemoveItem 0

1050      If gIncomp.Rows > 2 And gIncomp.TextMatrix(0, 0) = "" Then gIncomp.RemoveItem 0




1060      Exit Sub

bStart_Click_Error:

          Dim strES As String
          Dim intEL As Integer

1070      intEL = Erl
1080      strES = Err.Description
1090      LogError "frmMissing", "bStart_Click", intEL, strES, sql


End Sub


Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        tYear = Format(Now, "yyyy")

30        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmMissing", "Form_Load", intEL, strES


End Sub
