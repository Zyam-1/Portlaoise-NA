VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmViewHaemRep 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - View Repeats"
   ClientHeight    =   6660
   ClientLeft      =   225
   ClientTop       =   1230
   ClientWidth     =   13335
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C0C0C0&
   Icon            =   "frmViewHaemRep.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6660
   ScaleWidth      =   13335
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkRetics 
      Caption         =   "Retics"
      Height          =   315
      Left            =   6900
      TabIndex        =   12
      Top             =   780
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CheckBox chkESR 
      Caption         =   "ESR"
      Height          =   315
      Left            =   6900
      TabIndex        =   11
      Top             =   420
      Width           =   1695
   End
   Begin VB.CheckBox chkFBCRec 
      Caption         =   "FBC/Retics"
      Height          =   315
      Left            =   6900
      TabIndex        =   10
      Top             =   60
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   5355
      Left            =   90
      TabIndex        =   8
      Top             =   1260
      Width           =   13125
      _ExtentX        =   23151
      _ExtentY        =   9446
      _Version        =   393216
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      GridLines       =   3
      GridLinesFixed  =   3
      FormatString    =   "<Flags "
   End
   Begin VB.CommandButton bswap 
      Appearance      =   0  'Flat
      Caption         =   "&Swap Haematology Results"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4080
      TabIndex        =   7
      Top             =   360
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton bdelete 
      Appearance      =   0  'Flat
      Caption         =   "&Delete Haematology Repeat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   9105
      Picture         =   "frmViewHaemRep.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   45
      Visible         =   0   'False
      Width           =   2310
   End
   Begin VB.CommandButton bmove 
      Appearance      =   0  'Flat
      Caption         =   "&Move Haematology Result"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4080
      TabIndex        =   5
      Top             =   30
      Visible         =   0   'False
      Width           =   2535
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
      Height          =   675
      Left            =   11835
      Picture         =   "frmViewHaemRep.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   45
      Width           =   1245
   End
   Begin VB.Label lRunDate 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2310
      TabIndex        =   9
      Top             =   60
      Width           =   1635
   End
   Begin VB.Label lname 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1020
      TabIndex        =   4
      Top             =   360
      Width           =   2925
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Name"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   480
      TabIndex        =   3
      Top             =   390
      Width           =   495
   End
   Begin VB.Label lSampleID 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1020
      TabIndex        =   1
      Top             =   60
      Width           =   1245
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Sample ID"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   885
   End
End
Attribute VB_Name = "frmViewHaemRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mEditForm As Form
Private Activated As Boolean

Private Sub bcancel_Click()

10    Unload Me

End Sub

Private Sub bDELETE_Click()

      Dim sql As String



10    On Error GoTo bDELETE_Click_Error

20    g.Col = 0
30    sql = "DELETE from haemrepeats WHERE " & _
            "RunDateTime = '" & _
            Format$(g.TextMatrix(g.Row, 1), "dd/MMM/yyyy hh:mm:ss") & "' " & _
            "and SampleID = '" & Trim$(lSampleID) & "'"
40    Cnxn(0).Execute sql

50    sql = "DELETE from haemflagsrep WHERE " & _
            "DateTime = '" & _
            Format$(g.TextMatrix(g.Row, 1), "dd/MMM/yyyy hh:mm:ss") & "' " & _
            "and SampleID = '" & Trim$(lSampleID) & "'"
60    Cnxn(0).Execute sql

70    g.HighLight = False
80    bmove.Visible = False
90    bdelete.Visible = False
100   bswap.Visible = False

110   chkFBCRec.Visible = False
120   chkESR.Visible = False
130   chkRetics.Visible = False

140   FillG




150   Exit Sub

bDELETE_Click_Error:

      Dim strES As String
      Dim intEL As Integer



160   intEL = Erl
170   strES = Err.Description
180   LogError "frmViewHaemRep", "bDELETE_Click", intEL, strES, sql


End Sub




Private Function HaemFalgsExists(ByVal HaemOrRepeat As String, tb As Recordset) As String

      Dim sn As Recordset
      Dim sql As String
      Dim T As String

10    On Error GoTo HaemFalgsExists_Error


20    T = ""
30    sql = "SELECT * from " & HaemOrRepeat & " WHERE Sampleid = '" & lSampleID & "' AND " & _
          "DateTime between '" & Format$(tb!RunDateTime & "", "dd/mmm/yyyy hh:mm") & ":00'  and '" & Format$(tb!RunDateTime & "", "dd/mmm/yyyy hh:mm") & ":59' "
40    Set sn = New Recordset
50    RecOpenServer 0, sn, sql
60    If Not sn.EOF Then T = "Y"
70    If SysOptHaemAn1(0) = "ADVIA" Then
80        If GetOptionSetting("EnableAdviaOldFlags", 1) = 1 Then
90            If Trim(tb!LS & "") <> "" Or Trim(tb!va & "") <> "" _
                 Or Trim(tb!At & "") <> "" Or Trim(tb!bl & "") <> "" _
                 Or Trim(tb!An & "") <> "" Or Trim(tb!mi & "") <> "" _
                 Or Trim(tb!ca & "") <> "" Or Trim(tb!ho & "") <> "" _
                 Or Trim(tb!he & "") <> "" Or Trim(tb!Ig & "") <> "" _
                 Or Trim(tb!mpo & "") <> "" Or Trim(tb!lplt & "") <> "" _
                 Or Trim(tb!pclm & "") <> "" Or Trim(tb!rbcf & "") <> "" _
                 Or Trim(tb!rbcg & "") <> "" Then
100               T = "Y"
110           End If
120       End If

130   HaemFalgsExists = T

140   End If



150   Exit Function
HaemFalgsExists_Error:

160   LogError "frmViewHaemRep", "HaemFalgsExists", Erl, Err.Description, sql


End Function


Private Sub FillG()

          Dim tb As New Recordset
          Dim tbList As New Recordset
          Dim sql As String
          Dim s As String
          Dim sn As New Recordset
          Dim T As String
          Dim FS As String    'feildsName String
          Dim ListTest As String
          Dim NumberOfTests As Integer
          Dim i As Integer
          Dim J As Integer
          Dim listtypeA() As String


10        On Error GoTo FillG_Error

20        ClearFGrid g
30        g.Cols = 1
          '---------------Farhan --------------Dynamic grid cols
40        sql = "SELECT * from lists WHERE " & _
              "listType = 'HaemRepeatTests' " & _
              "order by Listorder"
50        Set tb = New Recordset
60        RecOpenServer 0, tbList, sql
70        If Not tbList.EOF Then
              'g.Cols = g.Cols - 1
80            g.AddItem ""
90            Do Until tbList.EOF
100               g.Cols = g.Cols + 1
110               If tbList!InUse = False Then
120                   g.ColWidth(g.Cols - 1) = 0
130               End If
140               g.TextMatrix(0, g.Cols - 1) = tbList!Text & ""
150               g.TextMatrix(1, g.Cols - 1) = tbList!Code & ""
160               NumberOfTests = NumberOfTests + 1
170               tbList.MoveNext
180           Loop
190       End If

          '-----------loading Haemresults in grid---------------------
200       sql = "SELECT * from haemresults WHERE " & _
              "sampleid = '" & lSampleID & "'"
210       Set tb = New Recordset
220       RecOpenServer 0, tb, sql
230       If Not tb.EOF Then

240           s = HaemFalgsExists("HaemFlags", tb)
250           g.AddItem s, g.Rows - 1
260           For J = 1 To g.Cols - 1
270               g.TextMatrix(g.Rows - 2, J) = LTrim(RTrim(tb(g.TextMatrix(1, J)) & ""))
280           Next J
290       End If
          '======================================================


300       sql = "SELECT * from haemrepeats WHERE " & _
              "sampleid = '" & lSampleID & "' " & _
              "order by rundatetime"
310       Set tb = New Recordset
320       RecOpenServer 0, tb, sql

330       Do While Not tb.EOF

340           s = HaemFalgsExists("HaemFlagsRep", tb)
350           g.AddItem s, g.Rows
360           i = g.Rows - 1
370           For J = 1 To g.Cols - 1
380               g.TextMatrix(i, J) = LTrim(RTrim(tb(g.TextMatrix(1, J)) & ""))
390           Next J

400           tb.MoveNext
410       Loop


420       g.Visible = True
430       FixGridColWidth g, Me
440       g.RowHeight(1) = 0

450       Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer



460       intEL = Erl
470       strES = Err.Description
480       LogError "frmViewHaemRep", "FillG", intEL, strES, sql


End Sub


Private Sub UpdateHaem(ByVal HaemOrRepeat As String, GridRow As Integer, ByVal RunDateTime As String)

      Dim GridCol As Integer
      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo UpdateHaem_Error

20    sql = "SELECT * FROM " & HaemOrRepeat & " WHERE SampleID = " & lSampleID & " AND RunDateTime = '" & Format(RunDateTime, "dd/MMM/yyyy HH:mm:ss") & "'"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    If Not tb.EOF Then
60        For GridCol = 1 To g.Cols - 1
70            If chkFBCRec Then
80                If UCase(g.TextMatrix(1, GridCol)) <> "ESR" And UCase(g.TextMatrix(1, GridCol)) <> UCase("RetA") And UCase(g.TextMatrix(1, GridCol)) <> UCase("RetP") Then
90                    tb(g.TextMatrix(1, GridCol)) = IIf(g.TextMatrix(GridRow, GridCol) = "", Null, g.TextMatrix(GridRow, GridCol))
100               End If
180               If UCase(g.TextMatrix(1, GridCol)) = UCase("RetA") Or UCase(g.TextMatrix(1, GridCol)) = UCase("RetP") Then
190                   tb(g.TextMatrix(1, GridCol)) = IIf(g.TextMatrix(GridRow, GridCol) = "", Null, g.TextMatrix(GridRow, GridCol))
200               End If
110           End If
120           If chkESR Then
130               If UCase(g.TextMatrix(1, GridCol)) = "ESR" Then
140                   tb(g.TextMatrix(1, GridCol)) = IIf(g.TextMatrix(GridRow, GridCol) = "", Null, g.TextMatrix(GridRow, GridCol))
150               End If
160           End If
220       Next GridCol
230       tb.Update
240   End If

250   Exit Sub
UpdateHaem_Error:

260   LogError "frmViewHaemRep", "UpdateHaem", Erl, Err.Description, sql


End Sub

Function ConvertDateFormat(inputDate As String) As String
    ' Split the input date string
    Dim dateParts() As String
    dateParts = Split(inputDate, " ")
    
    ' Split the date part further into day, month, and year
    Dim datePart() As String
    datePart = Split(dateParts(0), "/")
    
    ' Map month abbreviations to month numbers
    Dim month As String
    Select Case datePart(1)
        Case "Jan": month = "01"
        Case "Feb": month = "02"
        Case "Mar": month = "03"
        Case "Apr": month = "04"
        Case "May": month = "05"
        Case "Jun": month = "06"
        Case "Jul": month = "07"
        Case "Aug": month = "08"
        Case "Sep": month = "09"
        Case "Oct": month = "10"
        Case "Nov": month = "11"
        Case "Dec": month = "12"
    End Select
    
    ' Construct the new date format
    Dim formattedDate As String
    formattedDate = datePart(2) & "-" & month & "-" & datePart(0) & " " & dateParts(1) & ".813"
    
    ' Return the formatted date
    ConvertDateFormat = formattedDate
End Function

Private Sub bmove_Click()

      Dim sql As String

10    On Error GoTo bmove_Click_Error

20    UpdateHaem "HaemResults", g.Row, g.TextMatrix(2, 1)
30    sql = "DELETE FROM HaemRepeats WHERE SampleID = " & lSampleID & " AND RunDateTime = '" & Format(g.TextMatrix(g.Row, 1), "dd/MMM/yyyy HH:mm:ss") & "'"
40    Cnxn(0).Execute sql
50    If g.TextMatrix(g.Row, 0) = "Y" Then

60        sql = "DELETE FROM HaemFlags WHERE SampleID = '" & lSampleID & "'; " & _
                "INSERT INTO HaemFlags SELECT * FROM HaemFlagsRep WHERE SampleID = '" & lSampleID & "' and " & _
                "DateTime between '" & Format$(g, "dd/mmm/yyyy hh:mm") & ":00'  and '" & Format$(g, "dd/mmm/yyyy hh:mm") & ":59'; " & _
                "DELETE FROM HaemFlagsRep WHERE SampleID = '" & lSampleID & "' and " & _
                "DateTime between '" & Format$(g, "dd/mmm/yyyy hh:mm") & ":00'  and '" & Format$(g, "dd/mmm/yyyy hh:mm") & ":59'; " & _
                "DROP TABLE #temp"

70        Cnxn(0).Execute sql
80    End If

90    Unload Me

100   Exit Sub
bmove_Click_Error:

110   LogError "frmViewHaemRep", "bmove_Click", Erl, Err.Description, sql


End Sub

Private Sub bswap_Click()

      Dim ResultDate As String
      Dim RepeatDate As String
      Dim sql        As String

10    On Error GoTo bswap_Click_Error


20    ResultDate = g.TextMatrix(2, 1)
30    RepeatDate = g.TextMatrix(g.Row, 1)
40    UpdateHaem "HaemResults", g.Row, ResultDate
50    UpdateHaem "HaemRepeats", 2, RepeatDate


60    sql = "SELECT * INTO #temp FROM HaemFlags WHERE SampleID = '" & lSampleID & "'; " & _
            "DELETE FROM HaemFlags WHERE SampleID = '" & lSampleID & "'; " & _
            "INSERT INTO HaemFlags SELECT * FROM HaemFlagsRep WHERE SampleID = '" & lSampleID & "' and " & _
            "DateTime between '" & Format$(RepeatDate, "dd/mmm/yyyy hh:mm") & ":00'  and '" & Format$(RepeatDate, "dd/mmm/yyyy hh:mm") & ":59'; " & _
            "DELETE FROM HaemFlagsRep WHERE SampleID = '" & lSampleID & "' and " & _
            "DateTime between '" & Format$(RepeatDate, "dd/mmm/yyyy hh:mm") & ":00'  and '" & Format$(RepeatDate, "dd/mmm/yyyy hh:mm") & ":59'; " & _
            "INSERT INTO HaemFlagsRep SELECT * FROM #temp; " & _
            "DROP TABLE #temp"

70    Cnxn(0).Execute sql


80    Unload Me

90    Exit Sub
bswap_Click_Error:

100   LogError "frmViewHaemRep", "bswap_Click", Erl, Err.Description


End Sub

Private Sub Form_Activate()

10    On Error GoTo Form_Activate_Error

20    If Not Activated Then
30        FillG
40        Activated = True
50    End If

60    Exit Sub

Form_Activate_Error:

      Dim strES As String
      Dim intEL As Integer



70    intEL = Erl
80    strES = Err.Description
90    LogError "frmViewHaemRep", "Form_Activate", intEL, strES


End Sub



Private Sub Form_Unload(Cancel As Integer)
10    On Error GoTo Form_Unload_Error

20    Activated = False

30    Exit Sub
Form_Unload_Error:
         
40    LogError "frmViewHaemRep", "Form_Unload", Erl, Err.Description


End Sub

Private Sub g_Click()

10    On Error GoTo g_Click_Error

20    If g.Row < 4 Then
30        g.HighLight = flexHighlightNever
40        bmove.Visible = False
50        bdelete.Visible = False
60        bswap.Visible = False

70        chkFBCRec.Visible = False
80        chkESR.Visible = False
90        'chkRetics.Visible = False
100   Else
110       g.HighLight = flexHighlightAlways
120       If Trim(UCase(HospName(0))) <> "TULLAMORE" Then bmove.Visible = True
130       bdelete.Visible = True
140       bswap.Visible = True
150       chkFBCRec.Visible = True
160       chkESR.Visible = True
170       'chkRetics.Visible = True

180       g.Col = 1
190       g.ColSel = g.Cols - 1
200       g.RowSel = g.Row

210       bswap.SetFocus
220   End If

230   If g.MouseCol = 0 And g.TextMatrix(g.RowSel, 0) = "Y" And g.Row > 3 Then
              
240       Unload frmHaemErrorsRep
250       With frmHaemErrorsRep
260           .Top = 100
270           .Left = 6000
280           .SampleID = lSampleID
290           .Datetime = Format(g.TextMatrix(g.RowSel, 1), "dd/MMM/yyyy HH:mm:ss")
300           .Show 1
310       End With
          
320   End If

330   Exit Sub

g_Click_Error:

      Dim strES      As String
      Dim intEL      As Integer



340   intEL = Erl
350   strES = Err.Description
360   LogError "frmViewHaemRep", "g_Click", intEL, strES


End Sub


Public Property Let EditForm(ByVal EditForm As Form)

10    On Error GoTo EditForm_Error

20    Set mEditForm = EditForm

30    Exit Property

EditForm_Error:

      Dim strES As String
      Dim intEL As Integer



40    intEL = Erl
50    strES = Err.Description
60    LogError "frmViewHaemRep", "EditForm", intEL, strES


End Property


