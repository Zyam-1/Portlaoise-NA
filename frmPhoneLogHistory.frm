VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPhoneLogHistory 
   Caption         =   "NetAcquire - Phone Log History"
   ClientHeight    =   6120
   ClientLeft      =   540
   ClientTop       =   915
   ClientWidth     =   14250
   Icon            =   "frmPhoneLogHistory.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   14250
   Begin VB.Frame Frame1 
      Caption         =   "Filter By"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   4995
      Begin VB.OptionButton optDate 
         Caption         =   "Date to Date"
         Height          =   255
         Left            =   180
         TabIndex        =   13
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton optSampleID 
         Caption         =   "Sample ID"
         Height          =   255
         Left            =   180
         TabIndex        =   11
         Top             =   1080
         Width           =   1035
      End
      Begin VB.TextBox txtSampleID 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1860
         TabIndex        =   6
         Top             =   1425
         Width           =   2925
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   330
         Left            =   1065
         TabIndex        =   7
         Top             =   660
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   582
         _Version        =   393216
         Format          =   59179009
         CurrentDate     =   38631
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   330
         Left            =   3090
         TabIndex        =   8
         Top             =   660
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   582
         _Version        =   393216
         Format          =   59179009
         CurrentDate     =   38631
      End
      Begin VB.Label Label4 
         Caption         =   "Enter SampleID"
         Height          =   285
         Left            =   660
         TabIndex        =   12
         Top             =   1440
         Width           =   1140
      End
      Begin VB.Label Label3 
         Caption         =   "From"
         Height          =   285
         Left            =   660
         TabIndex        =   10
         Top             =   705
         Width           =   1140
      End
      Begin VB.Label Label2 
         Caption         =   "To"
         Height          =   240
         Left            =   2820
         TabIndex        =   9
         Top             =   705
         Width           =   825
      End
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "Export to Excel"
      Height          =   1100
      Left            =   9960
      Picture         =   "frmPhoneLogHistory.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1140
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   1100
      Left            =   11280
      Picture         =   "frmPhoneLogHistory.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1140
      Width           =   1200
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   1100
      Left            =   5400
      Picture         =   "frmPhoneLogHistory.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1075
      Width           =   1200
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   3585
      Left            =   240
      TabIndex        =   0
      Top             =   2340
      Width           =   13890
      _ExtentX        =   24500
      _ExtentY        =   6324
      _Version        =   393216
      Cols            =   16
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   0
      ScrollBars      =   2
      AllowUserResizing=   1
      FormatString    =   $"frmPhoneLogHistory.frx":0D60
   End
   Begin VB.Image imgRedCross 
      Height          =   225
      Left            =   0
      Picture         =   "frmPhoneLogHistory.frx":0E47
      Top             =   0
      Visible         =   0   'False
      Width           =   210
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
      Left            =   9960
      TabIndex        =   4
      Top             =   720
      Visible         =   0   'False
      Width           =   1275
   End
End
Attribute VB_Name = "frmPhoneLogHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pSampleID As String

Private Sub FillG()

          Dim tb As Recordset
          Dim sql As String
          Dim s As String


10        On Error GoTo FillG_Error

20        g.Rows = 2
30        g.AddItem ""
40        g.RemoveItem 1

50        sql = "Select P.*, D.PatName From PhoneLog P Left Join Demographics D " & _
                "On P.SampleId = D.SampleId "
60        If optDate.Value = True Then
70            sql = sql & "Where DateTime Between '%date1' And '%date2'"
80            sql = Replace(sql, "%date1", Format(dtFrom.Value, "yyyy-MM-dd 00:00:00"))
90            sql = Replace(sql, "%date2", Format(dtTo.Value, "yyyy-MM-dd 23:59:59"))
100       ElseIf optSampleID.Value = True Then
110           sql = sql & "Where P.SampleID = '%sampleid'"
120           sql = Replace(sql, "%sampleid", txtSampleID)
130       End If

          'sql = "Select * from PhoneLog where " & _
           '      "SampleID = " & Val(txtSampleID) & " " & _
           '      "Order by DateTime Desc"

140       Set tb = New Recordset
150       RecOpenServer 0, tb, sql
160       Do While Not tb.EOF
170           s = Format$(tb!Datetime, "dd/mm/yy hh:mm") & vbTab & _
                  tb!SampleID & "" & vbTab & _
                  tb!PatName & "" & vbTab & _
                  IIf(InStr(tb!Discipline, "H"), "H", "") & vbTab & _
                  IIf(InStr(tb!Discipline, "B"), "B", "") & vbTab & _
                  IIf(InStr(tb!Discipline, "C"), "C", "") & vbTab & _
                  IIf(InStr(tb!Discipline, "I"), "I", "") & vbTab & _
                  IIf(InStr(tb!Discipline, "G"), "G", "") & vbTab & _
                  IIf(InStr(tb!Discipline, "X"), "X", "") & vbTab & _
                  IIf(InStr(tb!Discipline, "M"), "M", "") & vbTab & _
                  IIf(InStr(tb!Discipline, "E"), "E", "") & vbTab & _
                  tb!PhonedTo & vbTab & _
                  tb!Comment & vbTab & _
                  tb!PhonedBy & vbTab & _
                  vbTab & _
                  tb!Title & " " & tb!PersonName
180           g.AddItem s
190           g.Row = g.Rows - 1
200           g.Col = 14
210           g.CellPictureAlignment = flexAlignCenterCenter
220           Set g.CellPicture = imgRedCross.Picture
230           tb.MoveNext
240       Loop

250       If g.Rows > 2 Then
260           g.RemoveItem 1
270       End If



280       Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer



290       intEL = Erl
300       strES = Err.Description
310       LogError "frmPhoneLogHistory", "FillG", intEL, strES, sql


End Sub

Private Sub cmdCancel_Click()

10        Unload Me

End Sub


Private Sub cmdExcel_Click()

          Dim strHeading As String

10        On Error GoTo cmdExcel_Click_Error

20        strHeading = "Phone Log History" & vbCr
30        If optDate.Value = True Then
40            strHeading = strHeading & "From " & Format(dtFrom.Value, "dd/MM/yyyy") & " To " & Format(dtTo.Value, "dd/MM/yyyy") & vbCr
50        ElseIf optSampleID.Value = True Then
60            strHeading = strHeading & "For SampleId " & txtSampleID & vbCr
70        End If
80        ExportFlexGrid g, Me, strHeading


90        Exit Sub

cmdExcel_Click_Error:

          Dim strES As String
          Dim intEL As Integer

100       intEL = Erl
110       strES = Err.Description
120       LogError "frmPhoneLogHistory", "cmdExcel_Click", intEL, strES

End Sub

Private Sub cmdSearch_Click()

10        On Error GoTo cmdSearch_Click_Error

20        If optDate.Value = True Then
30            If DateDiff("d", dtTo.Value, dtFrom.Value) > 0 Then
40                iMsg "From date cannot be bigger than To date"
50                Exit Sub
60            End If
70        ElseIf optSampleID.Value = True Then
80            If txtSampleID = "" Then
90                iMsg "Enter sample id first"
100               Exit Sub
110           End If
120       End If

130       FillG

140       Exit Sub

cmdSearch_Click_Error:

          Dim strES As String
          Dim intEL As Integer



150       intEL = Erl
160       strES = Err.Description
170       LogError "frmPhoneLogHistory", "cmdSearch_Click", intEL, strES


End Sub

Private Sub Form_Activate()

10        On Error GoTo Form_Activate_Error

20        dtFrom.Value = Now - 1
30        dtTo.Value = Now
40        FillG

50        Exit Sub

Form_Activate_Error:

          Dim strES As String
          Dim intEL As Integer



60        intEL = Erl
70        strES = Err.Description
80        LogError "frmPhoneLogHistory", "Form_Activate", intEL, strES


End Sub

Public Property Let SampleID(ByVal strNewValue As String)

10        On Error GoTo SampleID_Error

20        pSampleID = strNewValue

30        Exit Property

SampleID_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmPhoneLogHistory", "SampleID", intEL, strES


End Property


Private Sub g_Click()

          Dim sql As String
          Dim Disc As String

10        On Error GoTo g_Click_Error

20        If g.Col = 14 Then
30            If g.Row > 0 Then
40                If iMsg("Are you sure you want to remove this phone log entry?", vbQuestion + vbYesNo) = vbYes Then
50                    sql = "DELETE FROM PhoneLog WHERE SampleID = '%sampleid'"
60                    sql = Replace(sql, "%sampleid", g.TextMatrix(g.Row, 1))
70                    Cnxn(0).Execute sql
80                    FillG
90                End If
100           End If
110       End If
120       Exit Sub

g_Click_Error:

          Dim strES As String
          Dim intEL As Integer

130       intEL = Erl
140       strES = Err.Description
150       LogError "frmPhoneLogHistory", "g_Click", intEL, strES, sql

End Sub

Private Sub g_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

          Dim s As String

10        On Error GoTo g_MouseMove_Error

20        If g.MouseCol = 0 Or g.MouseCol > 7 Or g.MouseRow = 0 Then
30            g.ToolTipText = ""
40            Exit Sub

50        End If



60        Select Case g.TextMatrix(g.MouseRow, g.MouseCol)
          Case "H": s = "Haematology"
70        Case "B": s = "Biochemistry"
80        Case "C": s = "Coagulation"
90        Case "I": s = "Immunology"
100       Case "G": s = "Blood Gas"
110       Case "E": s = "External"
120       Case "M": s = "Microbiology"
130       Case Else: s = ""
140       End Select

150       g.ToolTipText = s

160       Exit Sub

g_MouseMove_Error:

          Dim strES As String
          Dim intEL As Integer



170       intEL = Erl
180       strES = Err.Description
190       LogError "frmPhoneLogHistory", "g_MouseMove", intEL, strES


End Sub


Private Sub optDate_Click()
10        txtSampleID = ""
End Sub
