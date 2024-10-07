VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmAddress 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Add External Address"
   ClientHeight    =   7530
   ClientLeft      =   1950
   ClientTop       =   2265
   ClientWidth     =   12540
   ClipControls    =   0   'False
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
   Icon            =   "frmAddress.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7530
   ScaleWidth      =   12540
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
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
      Left            =   11340
      Picture         =   "frmAddress.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3870
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Frame Frame1 
      Caption         =   "Send To"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1590
      Left            =   90
      TabIndex        =   3
      Top             =   90
      Width           =   11070
      Begin VB.CommandButton cmdAdd 
         Appearance      =   0  'Flat
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
         Height          =   780
         Left            =   9180
         Picture         =   "frmAddress.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   660
         Width           =   840
      End
      Begin VB.TextBox txtPhone 
         Height          =   285
         Left            =   6165
         MaxLength       =   15
         TabIndex        =   12
         Top             =   675
         Width           =   2520
      End
      Begin VB.TextBox txtFax 
         Height          =   285
         Left            =   6165
         MaxLength       =   15
         TabIndex        =   11
         Top             =   1170
         Width           =   2520
      End
      Begin VB.TextBox txtAddr 
         Height          =   285
         Index           =   2
         Left            =   1560
         MaxLength       =   40
         TabIndex        =   8
         Top             =   870
         Width           =   3915
      End
      Begin VB.TextBox txtAddr 
         Height          =   285
         Index           =   0
         Left            =   900
         MaxLength       =   40
         TabIndex        =   7
         Top             =   210
         Width           =   4575
      End
      Begin VB.TextBox txtAddr 
         Height          =   285
         Index           =   1
         Left            =   1560
         MaxLength       =   40
         TabIndex        =   6
         Top             =   570
         Width           =   3915
      End
      Begin VB.TextBox txtAddr 
         Height          =   285
         Index           =   3
         Left            =   1560
         MaxLength       =   40
         TabIndex        =   5
         Top             =   1170
         Width           =   3915
      End
      Begin VB.TextBox txtCode 
         Height          =   285
         Left            =   6165
         MaxLength       =   5
         TabIndex        =   4
         Top             =   210
         Width           =   735
      End
      Begin VB.Label Label1 
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
         Left            =   5775
         TabIndex        =   14
         Top             =   1215
         Width           =   300
      End
      Begin VB.Label Label5 
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
         Left            =   5610
         TabIndex        =   13
         Top             =   705
         Width           =   465
      End
      Begin VB.Label Label6 
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
         Left            =   900
         TabIndex        =   10
         Top             =   600
         Width           =   570
      End
      Begin VB.Label Label2 
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
         Left            =   5700
         TabIndex        =   9
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdPrint 
      Appearance      =   0  'Flat
      Caption         =   "Print"
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
      Left            =   11340
      Picture         =   "frmAddress.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5550
      Width           =   1080
   End
   Begin MSFlexGridLib.MSFlexGrid grdAddr 
      Height          =   5505
      Left            =   60
      TabIndex        =   1
      Top             =   1770
      Width           =   11160
      _ExtentX        =   19685
      _ExtentY        =   9710
      _Version        =   393216
      Cols            =   7
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
      FormatString    =   $"frmAddress.frx":0C28
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
   Begin VB.CommandButton cmdCancel 
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
      Height          =   750
      Left            =   11340
      Picture         =   "frmAddress.frx":0CF5
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6540
      Width           =   1080
   End
End
Attribute VB_Name = "frmAddress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdadd_Click()

          Dim tb As New Recordset
          Dim sql As String
          Dim Ans As Long

10        On Error GoTo cmdadd_Click_Error

20        cmdDelete.Visible = False

30        If Trim(txtAddr(0)) = "" Then
40            iMsg "First line of Address must be filled.", vbCritical, "Save Error"
50            Exit Sub
60        End If

70        If Trim(txtCode) = "" Then
80            iMsg "Code must be entered.", vbCritical, "Save Error"
90            Exit Sub
100       End If

110       sql = "SELECT * FROM ExtAddress WHERE " & _
                "Code = '" & txtCode & "'"
120       Set tb = New Recordset
130       RecOpenServer 0, tb, sql
140       If tb.EOF Then
150           tb.AddNew
160       Else
170           Ans = iMsg("This code already used. Edit this entry?", vbQuestion + vbYesNo, "NetAcquire")
180           If Ans <> vbYes Then
190               Exit Sub
200           End If
210       End If

220       tb!Code = UCase(txtCode)
230       tb!Addr0 = txtAddr(0)
240       tb!Addr1 = txtAddr(1)
250       tb!addr2 = txtAddr(2)
260       tb!addr3 = txtAddr(3)
270       tb!Phone = txtPhone
280       tb!FAX = txtFAX

290       tb.Update

300       txtCode = ""
310       txtAddr(0) = ""
320       txtAddr(1) = ""
330       txtAddr(2) = ""
340       txtAddr(3) = ""
350       txtPhone = ""
360       txtFAX = ""

370       FillG

380       Exit Sub

cmdadd_Click_Error:

          Dim strES As String
          Dim intEL As Integer

390       intEL = Erl
400       strES = Err.Description
410       LogError "frmAddress", "cmdAdd_Click", intEL, strES, sql

End Sub

Private Sub cmdCancel_Click()

10        Unload Me

End Sub

Private Sub cmdDelete_Click()

          Dim sql As String

10        On Error GoTo cmdDelete_Click_Error

20        sql = "DELETE FROM ExtAddress WHERE " & _
                "Code = '" & txtCode & "'"
30        Cnxn(0).Execute sql

40        txtCode = ""
50        txtAddr(0) = ""
60        txtAddr(1) = ""
70        txtAddr(2) = ""
80        txtAddr(3) = ""
90        txtPhone = ""
100       txtFAX = ""

110       FillG

120       cmdDelete.Visible = False

130       Exit Sub

cmdDelete_Click_Error:

          Dim strES As String
          Dim intEL As Integer

140       intEL = Erl
150       strES = Err.Description
160       LogError "frmAddress", "cmdDelete_Click", intEL, strES, sql


End Sub

Private Sub cmdPrint_Click()

          Dim Num As Long

10        On Error GoTo cmdPrint_Click_Error

20        cmdDelete.Visible = False

30        Printer.Print
40        Printer.Font.Name = "Courier New"
50        Printer.Font.Size = 12

60        For Num = 0 To grdAddr.Rows - 1
70            grdAddr.Row = Num
80            grdAddr.Col = 0
90            Printer.Print grdAddr; Tab(10);
100           grdAddr.Col = 1
110           Printer.Print grdAddr; Tab(40);
120           grdAddr.Col = 2
130           Printer.Print grdAddr; Tab(70);
140           grdAddr.Col = 3
150           Printer.Print grdAddr; Tab(100);
160           grdAddr.Col = 4
170           Printer.Print grdAddr; Tab(110);
180           grdAddr.Col = 5
190           Printer.Print grdAddr; Tab(130);
200           grdAddr.Col = 6
210           Printer.Print grdAddr
220       Next

230       Printer.EndDoc

240       Exit Sub

cmdPrint_Click_Error:

          Dim strES As String
          Dim intEL As Integer

250       intEL = Erl
260       strES = Err.Description
270       LogError "frmAddress", "cmdPrint_Click", intEL, strES

End Sub

Private Sub FillG()

          Dim tb As New Recordset
          Dim sql As String
          Dim Str As String

10        On Error GoTo FillG_Error

20        sql = "SELECT * from ExtAddress order by code"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql

50        ClearFGrid grdAddr

60        Do While Not tb.EOF
70            Str = tb!Code & vbTab & _
                    tb!Addr0 & vbTab & _
                    tb!Addr1 & vbTab & _
                    tb!addr2 & vbTab & _
                    tb!addr3 & vbTab & _
                    tb!Phone & vbTab & _
                    tb!FAX & ""
80            grdAddr.AddItem Str
90            tb.MoveNext
100       Loop

110       FixG grdAddr

120       Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

130       intEL = Erl
140       strES = Err.Description
150       LogError "frmAddress", "FillG", intEL, strES, sql

End Sub

Private Sub Form_Load()

10        FillG

End Sub

Private Sub grdAddr_Click()

          Static SortOrder As Boolean

10        On Error GoTo grdAddr_Click_Error

20        With grdAddr

30            If .MouseRow = 0 Then

40                If SortOrder Then
50                    .Sort = flexSortGenericAscending
60                Else
70                    .Sort = flexSortGenericDescending
80                End If
90                SortOrder = Not SortOrder
100               Exit Sub
110               cmdDelete.Visible = False
120           Else

130               txtCode = .TextMatrix(.Row, 0)
140               txtAddr(0) = .TextMatrix(.Row, 1)
150               txtAddr(1) = .TextMatrix(.Row, 2)
160               txtAddr(2) = .TextMatrix(.Row, 3)
170               txtAddr(3) = .TextMatrix(.Row, 4)
180               txtPhone = .TextMatrix(.Row, 5)
190               txtFAX = .TextMatrix(.Row, 6)
200               cmdDelete.Visible = True
210           End If

220       End With

230       Exit Sub

grdAddr_Click_Error:

          Dim strES As String
          Dim intEL As Integer

240       intEL = Erl
250       strES = Err.Description
260       LogError "frmAddress", "grdAddr_Click", intEL, strES

End Sub

Private Sub txtAddr_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

10        cmdDelete.Visible = False

End Sub


Private Sub txtCode_KeyUp(KeyCode As Integer, Shift As Integer)

10        cmdDelete.Visible = False

End Sub


Private Sub txtFax_KeyUp(KeyCode As Integer, Shift As Integer)

10        cmdDelete.Visible = False

End Sub


Private Sub txtPhone_KeyUp(KeyCode As Integer, Shift As Integer)

10        cmdDelete.Visible = False

End Sub


