VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmBatLactose 
   Caption         =   "NetAcquire - Batch Entry - Lactose/Urea & Purity"
   ClientHeight    =   6915
   ClientLeft      =   375
   ClientTop       =   735
   ClientWidth     =   11055
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
   Icon            =   "frmBatchLactose.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6915
   ScaleWidth      =   11055
   Begin VB.CommandButton cmdSave 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Save"
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
      Left            =   1530
      Picture         =   "frmBatchLactose.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   5940
      Width           =   1245
   End
   Begin VB.CommandButton bu 
      Appearance      =   0  'Flat
      Caption         =   "Up"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   8865
      TabIndex        =   25
      Top             =   1050
      Width           =   495
   End
   Begin VB.CommandButton bl 
      Appearance      =   0  'Flat
      Caption         =   "Lac"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   9345
      TabIndex        =   24
      Top             =   1050
      Width           =   495
   End
   Begin VB.CommandButton br 
      Appearance      =   0  'Flat
      Caption         =   "Urea"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   9825
      TabIndex        =   23
      Top             =   1050
      Width           =   495
   End
   Begin VB.CommandButton bp 
      Appearance      =   0  'Flat
      Caption         =   "Pur"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   10305
      TabIndex        =   22
      Top             =   1050
      Width           =   495
   End
   Begin MSFlexGridLib.MSFlexGrid grdLact 
      Height          =   4485
      Left            =   90
      TabIndex        =   21
      Top             =   1350
      Width           =   10890
      _ExtentX        =   19209
      _ExtentY        =   7911
      _Version        =   393216
      Rows            =   1
      Cols            =   21
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   -2147483638
      GridLines       =   2
   End
   Begin VB.CommandButton br 
      Appearance      =   0  'Flat
      Caption         =   "Urea"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   4050
      TabIndex        =   1
      Top             =   1050
      Width           =   495
   End
   Begin VB.CommandButton br 
      Appearance      =   0  'Flat
      Caption         =   "Urea"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2130
      TabIndex        =   20
      Top             =   1050
      Width           =   495
   End
   Begin VB.CommandButton bl 
      Appearance      =   0  'Flat
      Caption         =   "Lac"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1650
      TabIndex        =   19
      Top             =   1050
      Width           =   465
   End
   Begin VB.CommandButton bp 
      Appearance      =   0  'Flat
      Caption         =   "Pur"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   8370
      TabIndex        =   18
      Top             =   1050
      Width           =   495
   End
   Begin VB.CommandButton bp 
      Appearance      =   0  'Flat
      Caption         =   "Pur"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   4530
      TabIndex        =   17
      Top             =   1050
      Width           =   495
   End
   Begin VB.CommandButton bp 
      Appearance      =   0  'Flat
      Caption         =   "Pur"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2610
      TabIndex        =   16
      Top             =   1050
      Width           =   495
   End
   Begin VB.CommandButton br 
      Appearance      =   0  'Flat
      Caption         =   "Urea"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   5970
      TabIndex        =   15
      Top             =   1050
      Width           =   495
   End
   Begin VB.CommandButton bp 
      Appearance      =   0  'Flat
      Caption         =   "Pur"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   6450
      TabIndex        =   14
      Top             =   1050
      Width           =   495
   End
   Begin VB.CommandButton br 
      Appearance      =   0  'Flat
      Caption         =   "Urea"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   7890
      TabIndex        =   13
      Top             =   1050
      Width           =   495
   End
   Begin VB.CommandButton bl 
      Appearance      =   0  'Flat
      Caption         =   "Lac"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   3570
      TabIndex        =   12
      Top             =   1050
      Width           =   465
   End
   Begin VB.CommandButton bl 
      Appearance      =   0  'Flat
      Caption         =   "Lac"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   7410
      TabIndex        =   11
      Top             =   1050
      Width           =   495
   End
   Begin VB.CommandButton bu 
      Appearance      =   0  'Flat
      Caption         =   "Up"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   6930
      TabIndex        =   10
      Top             =   1050
      Width           =   495
   End
   Begin VB.CommandButton bu 
      Appearance      =   0  'Flat
      Caption         =   "Up"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   5010
      TabIndex        =   9
      Top             =   1050
      Width           =   495
   End
   Begin VB.CommandButton bl 
      Appearance      =   0  'Flat
      Caption         =   "Lac"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   5490
      TabIndex        =   8
      Top             =   1050
      Width           =   465
   End
   Begin VB.CommandButton bu 
      Appearance      =   0  'Flat
      Caption         =   "Up"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   3090
      TabIndex        =   7
      Top             =   1050
      Width           =   465
   End
   Begin VB.CommandButton bu 
      Appearance      =   0  'Flat
      Caption         =   "Up"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1170
      TabIndex        =   6
      Top             =   1050
      Width           =   465
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
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
      Left            =   9180
      Picture         =   "frmBatchLactose.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5940
      Width           =   1245
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Selenite"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8865
      TabIndex        =   26
      Top             =   780
      Width           =   1905
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Organism 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1170
      TabIndex        =   5
      Top             =   780
      Width           =   1935
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Organism 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3090
      TabIndex        =   4
      Top             =   780
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Organism 4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6930
      TabIndex        =   3
      Top             =   780
      Width           =   1905
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Organism 3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5010
      TabIndex        =   2
      Top             =   780
      Width           =   1935
   End
End
Attribute VB_Name = "frmBatLactose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub bl_Click(Index As Integer)

          Dim Y As Long

10        On Error GoTo bl_Click_Error

20        grdLact.Col = (Index - 1) * 4 + 2

30        For Y = 1 To grdLact.Rows - 1
40            grdLact.Row = Y
50            grdLact = "Neg"
60            saveculture
70        Next

80        Exit Sub

bl_Click_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmBatLactose", "bl_Click", intEL, strES

End Sub

Private Sub bp_Click(Index As Integer)

          Dim Y As Long

10        On Error GoTo bp_Click_Error

20        grdLact.Col = (Index - 1) * 4 + 4

30        For Y = 1 To grdLact.Rows - 1
40            grdLact.Row = Y
50            grdLact = "OK"
60            saveculture
70        Next

80        Exit Sub

bp_Click_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmBatLactose", "bp_Click", intEL, strES

End Sub

Private Sub br_Click(Index As Integer)

          Dim Y As Long

10        On Error GoTo br_Click_Error

20        grdLact.Col = (Index - 1) * 4 + 3

30        For Y = 1 To grdLact.Rows - 1
40            grdLact.Row = Y
50            grdLact = "Neg"
60            saveculture
70        Next

80        Exit Sub

br_Click_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmBatLactose", "br_Click", intEL, strES

End Sub

Private Sub bu_Click(Index As Integer)

          Dim Y As Long

10        On Error GoTo bu_Click_Error

20        grdLact.Col = (Index - 1) * 4 + 1

30        For Y = 1 To grdLact.Rows - 1
40            grdLact.Row = Y
50            grdLact = "Up"
60            saveculture
70        Next

80        Exit Sub

bu_Click_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmBatLactose", "bu_Click", intEL, strES

End Sub

Private Sub cmdCancel_Click()

10        Unload Me

End Sub

Private Sub cmdSave_Click()
          Dim X As Long
          Dim sql As String
          Dim tb As Recordset
          Dim n As Long
          Dim strb As String
          Dim IntB As String

10        On Error GoTo cmdSave_Click_Error

20        For n = 1 To grdLact.Rows - 1
30            sql = "SELECT * from faeces WHERE " & _
                    "Sampleid = " & Val(grdLact.TextMatrix(n, 0)) + SysOptMicroOffset(0)
40            Set tb = New Recordset
50            RecOpenServer 0, tb, sql
60            If tb.EOF Then
70                tb.AddNew
80                tb!SampleID = grdLact.TextMatrix(n, 0)
90            End If
100           strb = ""
110           For X = 2 To 21 Step 4
120               If Left(grdLact.TextMatrix(n, X), 1) = "" Then
130                   strb = strb & " "
140               Else
150                   strb = strb & Left(grdLact.TextMatrix(n, X), 1)
160               End If
170           Next
180           tb!lact = strb
190           strb = ""
200           For X = 3 To 21 Step 4
210               If Left(grdLact.TextMatrix(n, X), 1) = "" Then
220                   strb = strb & " "
230               Else
240                   strb = strb & Left(grdLact.TextMatrix(n, X), 1)
250               End If
260           Next
270           tb!urea = strb
280           IntB = 0
290           For X = 1 To 19 Step 4
300               If grdLact.TextMatrix(n, X) = "Up" Then
310                   If X = 1 Then IntB = IntB + 1
320                   If X = 5 Then IntB = IntB + 2
330                   If X = 9 Then IntB = IntB + 4
340                   If X = 13 Then IntB = IntB + 8
350                   If X = 17 Then IntB = IntB + 16
360               End If
370           Next
380           tb!Screen = IntB
390           IntB = 0
400           For X = 4 To 21 Step 4
410               If grdLact.TextMatrix(n, X) = "OK" Then
420                   If X = 4 Then IntB = IntB + 1
430                   If X = 8 Then IntB = IntB + 2
440                   If X = 12 Then IntB = IntB + 4
450                   If X = 16 Then IntB = IntB + 8
460                   If X = 20 Then IntB = IntB + 16
470               End If
480           Next
490           tb!purity = IntB

500           tb.Update

510       Next

520       Unload Me

530       Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

540       intEL = Erl
550       strES = Err.Description
560       LogError "frmBatLactose", "cmdsave_Click", intEL, strES, sql

End Sub

Private Sub FillG()

          Dim n As Long
          Dim sn As Recordset
          Dim sql As String
          Dim s As String

10        On Error GoTo FillG_Error

20        grdLact.ColWidth(0) = 1000
30        For n = 1 To 20
40            grdLact.ColWidth(n) = 487
50        Next

60        sql = "SELECT * From Faeces " & _
                "WHERE((pc = 'P') AND (Valid IS NULL) OR " & _
                "(pc = 'P') AND (Valid = 0) OR " & _
                "(pc = 'P') AND (Valid = '')) or ((selenite = 'P') AND (Valid IS NULL) OR " & _
                "(selenite = 'P') AND (Valid = 0) OR " & _
                "(selenite = 'P') AND (Valid = '')) ORDER BY SampleID"

70        Set sn = New Recordset
80        RecOpenServer 0, sn, sql

90        Do While Not sn.EOF
100           If sn!Pc = "P" Or sn!selenite = "P" Then
110               s = ""
120               s = s & sn!SampleID - SysOptMicroOffset(0) & vbTab
130               For n = 0 To 5

140                   If (IIf(sn!Screen And 2 ^ n, 1, 0)) = 1 Then
150                       s = s & "Up" & vbTab
160                   Else
170                       s = s & vbTab
180                   End If
190                   If Mid(sn!lact, n + 1, 1) = "P" Then
200                       s = s & "Pos" & vbTab
210                   ElseIf Mid(sn!lact, n + 1, 1) = "N" Then
220                       s = s & "Neg" & vbTab
230                   Else
240                       s = s & vbTab
250                   End If
260                   If Mid(sn!urea, n + 1, 1) = "P" Then
270                       s = s & "Pos" & vbTab
280                   ElseIf Mid(sn!urea, n + 1, 1) = "N" Then
290                       s = s & "Neg" & vbTab
300                   Else
310                       s = s & vbTab
320                   End If
330                   If (IIf(sn!purity And 2 ^ n, 1, 0)) = 1 Then
340                       s = s & "OK" & vbTab
350                   Else
360                       s = s & "" & vbTab
370                   End If
380               Next
390               grdLact.AddItem s
400           End If
410           sn.MoveNext
420       Loop

430       Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

440       intEL = Erl
450       strES = Err.Description
460       LogError "frmBatLactose", "FillG", intEL, strES, sql


End Sub

Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        FillG

30        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmBatLactose", "Form_Load", intEL, strES


End Sub

Private Sub grdLact_Click()

10        On Error GoTo grdLact_Click_Error

20        If grdLact.Col * grdLact.Row = 0 Then Exit Sub

30        Select Case grdLact.Col
          Case 1, 5, 9, 13, 17:    'done
40            grdLact = IIf(grdLact = "", "Up", "")
50        Case 2, 3, 6, 7, 10, 11, 14, 15, 18, 19:    'lactose/urea
60            If grdLact = "" Then
70                grdLact = "Neg"
80            ElseIf grdLact = "Neg" Then
90                grdLact = "Pos"
100           Else
110               grdLact = ""
120           End If
130       Case 4, 8, 12, 16, 20:    'purity
140           Select Case grdLact
              Case "OK": grdLact = "Con"
150           Case "Con": grdLact = ""
160           Case "": grdLact = "OK"
170           End Select
180       End Select

          'saveculture

190       Exit Sub

grdLact_Click_Error:

          Dim strES As String
          Dim intEL As Integer

200       intEL = Erl
210       strES = Err.Description
220       LogError "frmBatLactose", "grdLact_Click", intEL, strES


End Sub

Private Sub saveculture()

'Dim N as long
'Dim tb As Recordset
'Dim xsave as long
'Dim ysave as long
'
'xsave = grdLact.Col
'ysave = grdLact.Row
'
'Set tb = db.OpenRecordset("culture")
'tb.Index = "samplenumber"
'grdLact.Col = 0
'tb.Seek "=", g
'If tb.NoMatch Then
'  tb.AddNew
'  tb!SampleNumber = g
'Else
'  tb.Edit
'End If
'
'For N = 1 To 4
'  grdLact.Col = (N - 1) * 4 + 1
'  tb("s" & Format(N)) = IIf(g = "", False, True)
'  grdLact.Col = (N - 1) * 4 + 2
'  If g = "" Then
'    tb("ln" & Format(N)) = False
'    tb("lp" & Format(N)) = False
'  ElseIf g = "Neg" Then
'    tb("ln" & Format(N)) = True
'    tb("lp" & Format(N)) = False
'  Else
'    tb("ln" & Format(N)) = False
'    tb("lp" & Format(N)) = True
'  End If
'
'  grdLact.Col = (N - 1) * 4 + 3
'  If g = "" Then
'    tb("un" & Format(N)) = False
'    tb("up" & Format(N)) = False
'  ElseIf g = "Neg" Then
'    tb("un" & Format(N)) = True
'    tb("up" & Format(N)) = False
'  Else
'    tb("un" & Format(N)) = False
'    tb("up" & Format(N)) = True
'  End If
'
'  grdLact.Col = (N - 1) * 4 + 4
'  tb("p" & Format(N)) = IIf(g = "OK", True, False)
'Next
'
'tb.UPDATE
'
'grdLact.Col = xsave
'grdLact.Row = ysave

End Sub
