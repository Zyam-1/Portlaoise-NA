VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form frmBio2End 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetAcquire - Bio Related Endocrinology Results"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   5895
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy to Endocrinology"
      Height          =   1100
      Left            =   4500
      Picture         =   "frmBio2End.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   480
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Exit"
      Height          =   1100
      Left            =   4500
      Picture         =   "frmBio2End.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5355
      Width           =   1200
   End
   Begin MSFlexGridLib.MSFlexGrid grdB2I 
      Height          =   6030
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   4050
      _ExtentX        =   7144
      _ExtentY        =   10636
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      FormatString    =   "Test Name                       | Result         "
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
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "SampleID: 1234567"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2295
      TabIndex        =   4
      Top             =   240
      Width           =   1995
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Biochemistry results"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   1380
   End
End
Attribute VB_Name = "frmBio2End"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

10        Unload Me

End Sub

Private Sub cmdCopy_Click()

      Dim sql As String
      Dim tb As New Recordset
      Dim sn As New Recordset
      Dim n As Long
      Dim Code As String
      Dim DefIndex As Integer
      Dim BioDef As Recordset
      Dim EndDef As Recordset
      Dim tbNewIDX As Recordset

10    On Error GoTo cmdCopy_Click_Error


20    For n = 1 To grdB2I.Rows - 1
30        DefIndex = 0
40        grdB2I.Row = n
50        If grdB2I.CellBackColor = vbRed Then
60            Code = CodeForShortName(grdB2I.TextMatrix(n, 0))

70            sql = "SELECT * FROM BioResults WHERE " & _
                    "SampleID = '" & frmEditAll.txtSampleID & "' " & _
                    "AND Code = '" & Code & "'"
80            Set tb = New Recordset
90            RecOpenServer 0, tb, sql
100           If Not tb.EOF Then
110               If Not IsNull(tb!DefIndex) Then
120                   If tb!DefIndex > 0 Then
130                       sql = "SELECT * FROM BioDefIndex WHERE DefIndex = " & tb!DefIndex
140                       Set BioDef = New Recordset
150                       RecOpenServer 0, BioDef, sql
160                       If Not BioDef.EOF Then
170                           sql = "SELECT * FROM EndDefIndex WHERE " & _
                                    "NormalLow = " & BioDef!NormalLow & " AND " & _
                                    "NormalHigh= " & BioDef!NormalHigh & " AND " & _
                                    "FlagLow = " & BioDef!FlagLow & " AND " & _
                                    "FlagHigh = " & BioDef!FlagHigh & " AND " & _
                                    "PlausibleLow = " & BioDef!PlausibleLow & " AND " & _
                                    "PlausibleHigh = " & BioDef!PlausibleHigh
180                           Set EndDef = New Recordset
190                           RecOpenServer 0, EndDef, sql
200                           If Not EndDef.EOF Then
210                               DefIndex = EndDef!DefIndex
220                           Else
230                               sql = "INSERT INTO EndDefIndex " & _
                                        "( NormalLow, NormalHigh, FlagLow, FlagHigh, " & _
                                        "  PlausibleLow, PlausibleHigh, AutoValLow, AutoValHigh) " & _
                                        "VALUES ( " & _
                                        BioDef!NormalLow & ", " & BioDef!NormalHigh & ", " & BioDef!FlagLow & ", " & BioDef!FlagHigh & ", " & _
                                        BioDef!PlausibleLow & ", " & BioDef!PlausibleHigh & ", 0,9999) "
240                               Cnxn(0).Execute sql

250                               sql = "SELECT MAX(DefIndex) NewIndex FROM EndDefIndex"
260                               Set tbNewIDX = New Recordset
270                               RecOpenServer 0, tbNewIDX, sql
280                               DefIndex = tbNewIDX!NewIndex
290                           End If

300                       End If
310                   End If
320               End If
              
              
330               sql = "If Exists(Select 1 From EndResults " & _
                        "Where sampleid = @sampleid0 " & _
                        "And Code = '@Code1' ) " & _
                        "Begin " & _
                        "Insert Into EndRepeats (sampleid, Code, result, valid, printed, RunTime, RunDate, Operator, Units, SampleType, DefIndex) Values " & _
                        "(@sampleid0, '@Code1', '@result2', @valid3, @printed4, '@RunTime5', '@RunDate6', '@Operator7', '@Units9', '@SampleType10', @DefIndex11) " & _
                        "End " & _
                        "Else " & _
                        "Begin  " & _
                        "Insert Into EndResults (sampleid, Code, result, valid, printed, RunTime, RunDate, Operator, Units, SampleType, DefIndex) Values " & _
                        "(@sampleid0, '@Code1', '@result2', @valid3, @printed4, '@RunTime5', '@RunDate6', '@Operator7', '@Units9', '@SampleType10', @DefIndex11) " & _
                        "End"

340               sql = Replace(sql, "@sampleid0", tb!SampleID)
350               sql = Replace(sql, "@Code1", tb!Code & "")
360               sql = Replace(sql, "@result2", tb!Result & "")
370               sql = Replace(sql, "@valid3", tb!Valid)
380               sql = Replace(sql, "@printed4", 0)
390               sql = Replace(sql, "@RunTime5", Format(tb!RunTime, "dd/MMM/yyyy hh:mm:ss"))
400               sql = Replace(sql, "@RunDate6", Format(tb!RunTime, "dd/MMM/yyyy"))
410               sql = Replace(sql, "@Operator7", tb!Operator & "")
420               sql = Replace(sql, "@Units9", tb!Units & "")
430               sql = Replace(sql, "@SampleType10", tb!SampleType & "")
440               sql = Replace(sql, "@DefIndex11", DefIndex)

450               Cnxn(0).Execute sql

                  


                  '            sql = "SELECT * FROM EndResults WHERE " & _
                               '                  "SampleID = '" & frmEditAll.txtSampleID & "' " & _
                               '                  "AND Code = '" & Code & "'"
                  '            Set sn = New Recordset
                  '            RecOpenServer 0, sn, sql
                  '            If sn.EOF Then
                  '                sn.AddNew
                  '            Else
                  '                sql = "SELECT * FROM EndRepeats WHERE " & _
                                   '                      "SampleID = '" & frmEditAll.txtSampleID & "' " & _
                                   '                      "AND Code = '" & Code & "'"
                  '                Set sn = New Recordset
                  '                RecOpenServer 0, sn, sql
                  '                If sn.EOF Then
                  '                    sn.AddNew
                  '                End If
                  '            End If
                  '            sn!SampleID = tb!SampleID
                  '            sn!Code = tb!Code
                  '            sn!Result = tb!Result
                  '            sn!Valid = tb!Valid
                  '            sn!Printed = 0
                  '            sn!Rundate = Format(tb!RunTime, "dd/MMM/yyyy")
                  '            sn!RunTime = Format(tb!RunTime, "dd/MMM/yyyy hh:mm:ss")
                  '            sn!Operator = tb!Operator
                  '            sn!Units = tb!Units
                  '            sn!SampleType = tb!SampleType
                  '            sn.Update


460           End If
470       End If

480   Next

490   Unload Me

500   Exit Sub

cmdCopy_Click_Error:

      Dim strES As String
      Dim intEL As Integer

510   intEL = Erl
520   strES = Err.Description
530   LogError "frmBio2Imm", "cmdCopy_Click", intEL, strES, sql

End Sub



Private Sub Form_Load()

          Dim n As Long
          Dim s As String
          Dim Code As String

10        On Error GoTo Form_Load_Error

20        With grdB2I
30            .Rows = 2
40            .AddItem ""
50            .RemoveItem 1
60        End With

70        For n = 1 To frmEditAll.gBio.Rows - 1
              'Pickup only validated results
80            Code = CodeForShortName(frmEditAll.gBio.TextMatrix(n, 0))
90            If InStr(1, frmEditAll.gBio.TextMatrix(n, 6), "V") > 0 Then
100               s = frmEditAll.gBio.TextMatrix(n, 0) & vbTab & frmEditAll.gBio.TextMatrix(n, 1)
110               grdB2I.AddItem s
120           End If
130       Next

140       If grdB2I.Rows > 2 And grdB2I.TextMatrix(1, 0) = "" Then
150           grdB2I.RemoveItem 1

160       End If

170       Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

180       intEL = Erl
190       strES = Err.Description
200       LogError "frmBio2Imm", "Form_Load", intEL, strES

End Sub

Private Sub grdB2I_Click()

          Dim n As Long

10        On Error GoTo grdB2I_Click_Error

20        If grdB2I.CellBackColor = vbRed Then
30            For n = 0 To 1
40                grdB2I.Col = n
50                grdB2I.CellBackColor = vbWhite
60            Next
70        Else
80            For n = 0 To 1
90                grdB2I.Col = n
100               grdB2I.CellBackColor = vbRed
110           Next
120       End If

130       Exit Sub

grdB2I_Click_Error:

          Dim strES As String
          Dim intEL As Integer

140       intEL = Erl
150       strES = Err.Description
160       LogError "frmBio2Imm", "grdB2I_Click", intEL, strES

End Sub


