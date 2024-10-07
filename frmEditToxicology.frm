VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEditToxicology 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetAcquire --- Toxicology Results"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   3465
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   900
      Left            =   1860
      Picture         =   "frmEditToxicology.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Tag             =   "101"
      ToolTipText     =   "Exit Screen"
      Top             =   4440
      Width           =   1005
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   900
      Left            =   585
      Picture         =   "frmEditToxicology.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4440
      Width           =   1005
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   255
      Left            =   75
      TabIndex        =   1
      Top             =   5580
      Width           =   3330
      _ExtentX        =   5874
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   4215
      Left            =   585
      TabIndex        =   0
      Top             =   60
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   7435
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      RowHeightMin    =   315
      BackColorFixed  =   16744448
      ForeColorFixed  =   16777215
      FormatString    =   "Code|Short|<Test                                                             |<Reslts                     "
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
End
Attribute VB_Name = "frmEditToxicology"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

10    On Error GoTo cmdCancel_Click_Error

20    Unload Me

30    Exit Sub

cmdCancel_Click_Error:

       Dim strES As String
       Dim intEL As Integer

40     intEL = Erl
50     strES = Err.Description
60     LogError "frmEditToxicology", "cmdCancel_Click", intEL, strES
          
End Sub

Private Sub cmdSave_Click()

      Dim tb As New Recordset
      Dim sql As String
      Dim n As Long
      Dim s As String
      Dim R As Integer

10    On Error GoTo cmdSave_Click_Error

20    If Val(frmEditAll.txtSampleID) = 0 Then Exit Sub

30    For R = 1 To g.Rows - 1

40        pBar = 0

50        With frmEditAll
60            .txtSampleID = Format(Val(.txtSampleID))

70            For n = 1 To .gBio.Rows - 1
80                If g.TextMatrix(n, 1) = .gBio.TextMatrix(n, 0) Then
90                    iMsg "Test already Exists. Please delete before adding!"
100                   Exit Sub
110               End If
120           Next

              '    s = Check_Bio(cAdd.Text, cUnits, cISampleType(3))
              '    If s <> "" Then
              '        iMsg s & " is incorrect!"
              '        Exit Sub
              '    End If

              '    If cAdd.Text = "" Then Exit Sub

              '    If Len(cUnits) = 0 Then
              '        If iMsg("SELECT Units?", vbYesNo) = vbYes Then
              '            Exit Sub
              '        End If
              '    End If

130           sql = "INSERT into BioResults " & _
                    "(RunDate, SampleID, Code, Result, RunTime, Units, SampleType, Valid, Printed) VALUES " & _
                    "('" & Format$(.dtRunDate, "dd/mmm/yyyy") & "', " & _
                    "'" & .txtSampleID & "', " & _
                    "'" & g.TextMatrix(R, 0) & "', " & _
                    "'" & g.TextMatrix(R, 3) & "', " & _
                    "'" & Format$(Now, "dd/mmm/yyyy hh:mm") & "', " & _
                    "'', " & _
                    "'T', 0, 0);"
140           Cnxn(0).Execute sql

150           sql = "DELETE FROM BioRequests " & _
                    "WHERE SampleID = '" & .txtSampleID & "' " & _
                    "AND Code = '" & g.TextMatrix(R, 0) & "'"
160           Cnxn(0).Execute sql

              'Code added 22/08/05
              'This allows the user delete
              'oustanding requests where sample is bad
              'it also marks bad samples printed and valid
              '    If SysOptBioCodeForBad(0) = CodeForShortName(cAdd.Text) Then
              '        sql = "update bioresults set valid = 1, printed = 1, Operator = '" & AddTicks(UserCode) & "' " & _
                       '              "where code = '" & SysOptBioCodeForBad(0) & "' " & _
                       '              "and sampleID = '" & txtSampleID & "'"
              '        Cnxn(0).Execute sql
              '        If iMsg("Do you wish all outstanding requests Deleted!", vbYesNo) = vbYes Then
              '            sql = "DELETE from biorequests WHERE sampleID = '" & txtSampleID & "'"
              '            Cnxn(0).Execute sql
              '        End If
              '        txtBioComment = Trim(txtBioComment & " " & iBOX("Enter Bad Comment"))
              '        SaveComments
              '    End If
170       End With
180   Next R
190   Unload Me

200   Exit Sub

cmdSave_Click_Error:

      Dim strES As String
      Dim intEL As Integer

210   intEL = Erl
220   strES = Err.Description
230   LogError "frmEditToxicology", "cmdSave_Click", intEL, strES, sql

End Sub

Private Sub Form_Load()

      Dim sql As String
      Dim tb As Recordset


10    On Error GoTo Form_Load_Error

20    With g
30        .Rows = 1
40        .Cols = 4
50        .ColWidth(0) = 0
60        .ColWidth(2) = 0
          
70    End With

80    sql = "SELECT  Code, ShortName, LongName FROM BioTestDefinitions WHERE SampleType = 'T' Order by PrintPriority"
90    Set tb = New Recordset
100   RecOpenServer 0, tb, sql
110   If tb.EOF Then
120       Unload Me
130       Exit Sub
140   End If

150   While Not tb.EOF
160       g.AddItem tb!Code & "" & vbTab & tb!ShortName & "" & vbTab & tb!LongName & "" & vbTab & "Negative"
170       tb.MoveNext
180   Wend


190   Exit Sub

Form_Load_Error:

       Dim strES As String
       Dim intEL As Integer

200    intEL = Erl
210    strES = Err.Description
220    LogError "frmEditToxicology", "Form_Load", intEL, strES
          
End Sub

Private Sub g_Click()

10    On Error GoTo g_Click_Error

20    If g.MouseCol = 3 Then
30        Select Case g.TextMatrix(g.MouseRow, g.MouseCol)
              Case "Negative"
40                g.TextMatrix(g.MouseRow, g.MouseCol) = "Positive"
50            Case "Positive"
60                g.TextMatrix(g.MouseRow, g.MouseCol) = "Negative"
70        End Select
80    End If

90    Exit Sub

g_Click_Error:

       Dim strES As String
       Dim intEL As Integer

100    intEL = Erl
110    strES = Err.Description
120    LogError "frmEditToxicology", "g_Click", intEL, strES
          
End Sub
