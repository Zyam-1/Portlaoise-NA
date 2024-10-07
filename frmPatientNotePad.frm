VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Begin VB.Form frmPatientNotePad 
   Caption         =   "NetAcquire - Patient NotePad"
   ClientHeight    =   7320
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   11940
   LinkTopic       =   "Form1"
   ScaleHeight     =   7320
   ScaleWidth      =   11940
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame framDemographics 
      Caption         =   "Demographics"
      Height          =   1332
      Left            =   120
      TabIndex        =   5
      Top             =   60
      Width           =   11712
      Begin VB.Label lblForeNameD 
         AutoSize        =   -1  'True
         Caption         =   "ForeName:"
         ForeColor       =   &H00000000&
         Height          =   192
         Left            =   4020
         TabIndex        =   26
         Top             =   600
         Width           =   816
      End
      Begin VB.Label lblDOBD 
         AutoSize        =   -1  'True
         Caption         =   "DOB:"
         ForeColor       =   &H00000000&
         Height          =   192
         Left            =   6600
         TabIndex        =   25
         Top             =   600
         Width           =   384
      End
      Begin VB.Label lblAgeD 
         AutoSize        =   -1  'True
         Caption         =   "Age :"
         ForeColor       =   &H00000000&
         Height          =   192
         Left            =   6600
         TabIndex        =   24
         Top             =   900
         Width           =   372
      End
      Begin VB.Label lblChartNoD 
         AutoSize        =   -1  'True
         Caption         =   "Chart Number:"
         ForeColor       =   &H00000000&
         Height          =   192
         Left            =   1500
         TabIndex        =   23
         Top             =   600
         Width           =   1020
      End
      Begin VB.Label lblAddressD 
         AutoSize        =   -1  'True
         Caption         =   "Address :"
         ForeColor       =   &H00000000&
         Height          =   192
         Left            =   4020
         TabIndex        =   22
         Top             =   900
         Width           =   684
      End
      Begin VB.Label lblDemoDateD 
         AutoSize        =   -1  'True
         Caption         =   "Demographics Date:"
         ForeColor       =   &H00000000&
         Height          =   192
         Left            =   9720
         TabIndex        =   21
         Top             =   600
         Width           =   1488
      End
      Begin VB.Label lblSexD 
         AutoSize        =   -1  'True
         Caption         =   "Sex:"
         ForeColor       =   &H00000000&
         Height          =   192
         Left            =   6600
         TabIndex        =   20
         Top             =   300
         Width           =   312
      End
      Begin VB.Label lblSurNameD 
         AutoSize        =   -1  'True
         Caption         =   "SurName:"
         ForeColor       =   &H00000000&
         Height          =   192
         Left            =   4020
         TabIndex        =   19
         Top             =   300
         Width           =   720
      End
      Begin VB.Label lblSampleIDD 
         AutoSize        =   -1  'True
         Caption         =   "SampleID:"
         ForeColor       =   &H00000000&
         Height          =   192
         Left            =   1500
         TabIndex        =   18
         Top             =   300
         Width           =   756
      End
      Begin VB.Label lblSampleDateD 
         AutoSize        =   -1  'True
         Caption         =   "Sample Date:"
         ForeColor       =   &H00000000&
         Height          =   192
         Left            =   9720
         TabIndex        =   17
         Top             =   300
         Width           =   984
      End
      Begin VB.Label lblSampleDate 
         AutoSize        =   -1  'True
         Caption         =   "Sample Date:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   192
         Left            =   8556
         TabIndex        =   16
         Top             =   300
         Width           =   1140
      End
      Begin VB.Label lblSample 
         AutoSize        =   -1  'True
         Caption         =   "SampleID:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   192
         Left            =   552
         TabIndex        =   15
         Top             =   300
         Width           =   876
      End
      Begin VB.Label lblSurName 
         AutoSize        =   -1  'True
         Caption         =   "SurName:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   192
         Left            =   3096
         TabIndex        =   14
         Top             =   300
         Width           =   828
      End
      Begin VB.Label lblSex 
         AutoSize        =   -1  'True
         Caption         =   "Sex:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   192
         Left            =   6192
         TabIndex        =   13
         Top             =   300
         Width           =   372
      End
      Begin VB.Label lblDemoDate 
         AutoSize        =   -1  'True
         Caption         =   "Demographics Date:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   192
         Left            =   7980
         TabIndex        =   12
         Top             =   600
         Width           =   1716
      End
      Begin VB.Label lblAddress 
         AutoSize        =   -1  'True
         Caption         =   "Address :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   192
         Left            =   3120
         TabIndex        =   11
         Top             =   900
         Width           =   804
      End
      Begin VB.Label lblChartNo 
         AutoSize        =   -1  'True
         Caption         =   "Chart Number:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   192
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Width           =   1188
      End
      Begin VB.Label lblAge 
         AutoSize        =   -1  'True
         Caption         =   "Age :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   192
         Left            =   6120
         TabIndex        =   9
         Top             =   900
         Width           =   444
      End
      Begin VB.Label lblDOB 
         AutoSize        =   -1  'True
         Caption         =   "DOB:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   192
         Left            =   6120
         TabIndex        =   8
         Top             =   600
         Width           =   444
      End
      Begin VB.Label lblForeName 
         AutoSize        =   -1  'True
         Caption         =   "ForeName:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   192
         Left            =   2988
         TabIndex        =   7
         Top             =   600
         Width           =   936
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   1000
      Left            =   10560
      Picture         =   "frmPatientNotePad.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5880
      Width           =   1100
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   5475
      Left            =   60
      TabIndex        =   0
      Top             =   1740
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   9657
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   420
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmPatientNotePad.frx":0CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "grid(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdAddComments(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtComments(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdDelete(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdEdit(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Biochemistry"
      TabPicture(1)   =   "frmPatientNotePad.frx":0CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdEdit(1)"
      Tab(1).Control(1)=   "cmdDelete(1)"
      Tab(1).Control(2)=   "cmdAddComments(1)"
      Tab(1).Control(3)=   "txtComments(1)"
      Tab(1).Control(4)=   "grid(1)"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Haematology"
      TabPicture(2)   =   "frmPatientNotePad.frx":0D02
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdEdit(2)"
      Tab(2).Control(1)=   "cmdDelete(2)"
      Tab(2).Control(2)=   "cmdAddComments(2)"
      Tab(2).Control(3)=   "txtComments(2)"
      Tab(2).Control(4)=   "grid(2)"
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Microbiology"
      TabPicture(3)   =   "frmPatientNotePad.frx":0D1E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdDelete(3)"
      Tab(3).Control(1)=   "cmdEdit(3)"
      Tab(3).Control(2)=   "cmdAddComments(3)"
      Tab(3).Control(3)=   "txtComments(3)"
      Tab(3).Control(4)=   "grid(3)"
      Tab(3).ControlCount=   5
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   612
         Index           =   3
         Left            =   -64500
         TabIndex        =   43
         Top             =   1560
         Width           =   1152
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   612
         Index           =   3
         Left            =   -64500
         TabIndex        =   42
         Top             =   900
         Width           =   1152
      End
      Begin VB.CommandButton cmdAddComments 
         Caption         =   "Add Comment"
         Height          =   1000
         Index           =   3
         Left            =   -65900
         Picture         =   "frmPatientNotePad.frx":0D3A
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   4126
         Width           =   1100
      End
      Begin VB.TextBox txtComments 
         Height          =   1332
         Index           =   3
         Left            =   -74880
         MultiLine       =   -1  'True
         TabIndex        =   40
         Top             =   3960
         Width           =   8892
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   612
         Index           =   2
         Left            =   -64500
         TabIndex        =   38
         Top             =   900
         Width           =   1152
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   612
         Index           =   2
         Left            =   -64500
         TabIndex        =   37
         Top             =   1560
         Width           =   1152
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   612
         Index           =   1
         Left            =   -64500
         TabIndex        =   36
         Top             =   900
         Width           =   1152
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   612
         Index           =   1
         Left            =   -64500
         TabIndex        =   35
         Top             =   1560
         Width           =   1152
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Enabled         =   0   'False
         Height          =   612
         Index           =   0
         Left            =   10500
         TabIndex        =   34
         Top             =   900
         Width           =   1152
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   612
         Index           =   0
         Left            =   10500
         TabIndex        =   33
         Top             =   1560
         Width           =   1152
      End
      Begin VB.CommandButton cmdAddComments 
         Caption         =   "Add Comment"
         Height          =   1000
         Index           =   2
         Left            =   -65900
         Picture         =   "frmPatientNotePad.frx":1604
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   4126
         Width           =   1100
      End
      Begin VB.CommandButton cmdAddComments 
         Caption         =   "Add Comment"
         Height          =   1000
         Index           =   1
         Left            =   -65900
         Picture         =   "frmPatientNotePad.frx":1ECE
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   4126
         Width           =   1100
      End
      Begin VB.TextBox txtComments 
         Height          =   1332
         Index           =   2
         Left            =   -74880
         MultiLine       =   -1  'True
         TabIndex        =   30
         Top             =   3960
         Width           =   8892
      End
      Begin VB.TextBox txtComments 
         Height          =   1332
         Index           =   1
         Left            =   -74880
         MultiLine       =   -1  'True
         TabIndex        =   29
         Top             =   3960
         Width           =   8892
      End
      Begin VB.TextBox txtComments 
         Height          =   1332
         Index           =   0
         Left            =   120
         MaxLength       =   999
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   3960
         Width           =   8892
      End
      Begin VB.CommandButton cmdAddComments 
         Caption         =   "Add Comment"
         Height          =   1000
         Index           =   0
         Left            =   9100
         Picture         =   "frmPatientNotePad.frx":2798
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   4126
         Width           =   1100
      End
      Begin MSFlexGridLib.MSFlexGrid grid 
         Height          =   3312
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   10332
         _ExtentX        =   18230
         _ExtentY        =   5847
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         RowHeightMin    =   10
         WordWrap        =   -1  'True
         FormatString    =   $"frmPatientNotePad.frx":3062
      End
      Begin MSFlexGridLib.MSFlexGrid grid 
         Height          =   3312
         Index           =   1
         Left            =   -74880
         TabIndex        =   27
         Top             =   600
         Width           =   10332
         _ExtentX        =   18230
         _ExtentY        =   5847
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         RowHeightMin    =   10
         WordWrap        =   -1  'True
         FormatString    =   $"frmPatientNotePad.frx":312B
      End
      Begin MSFlexGridLib.MSFlexGrid grid 
         Height          =   3312
         Index           =   2
         Left            =   -74880
         TabIndex        =   28
         Top             =   600
         Width           =   10332
         _ExtentX        =   18230
         _ExtentY        =   5847
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         RowHeightMin    =   10
         WordWrap        =   -1  'True
         FormatString    =   $"frmPatientNotePad.frx":31F4
      End
      Begin MSFlexGridLib.MSFlexGrid grid 
         Height          =   3312
         Index           =   3
         Left            =   -74880
         TabIndex        =   39
         Top             =   600
         Width           =   10332
         _ExtentX        =   18230
         _ExtentY        =   5847
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         RowHeightMin    =   10
         WordWrap        =   -1  'True
         FormatString    =   $"frmPatientNotePad.frx":32BD
      End
   End
   Begin VB.Label lbl1 
      Caption         =   "Previous Comments:"
      Height          =   252
      Left            =   180
      TabIndex        =   6
      Top             =   1500
      Width           =   1692
   End
End
Attribute VB_Name = "frmPatientNotePad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public SampleID As String
Public Caller As String
Public GLabNo As String

Private Sub cmdAddComments_Click(Index As Integer)
      Dim sql As String

10    On Error GoTo cmdAddComments_Click_Error

20    If Trim(txtComments(Index)) = "" Then Exit Sub


30    sql = "INSERT INTO PatientNotePad " & _
            "(SampleID, DateTimeofRecord, Comment, UserName, Descipline, LabNo)" & _
            "VALUES ( '" & lblSampleIDD & "', GETDATE(),'" & WordWrap(txtComments(Index), 185) & "','" & UserName & "','" & SSTab.TabCaption(SSTab.Tab) & "','" & Val(GLabNo) & "')"
40    Cnxn(0).Execute sql
50    grid(Index).AddItem Format$(Now, "dd/mmm/yyyy hh:mm:ss") & vbTab & WordWrap(txtComments(Index), 185) & vbTab & UserName
60    grid(Index).RowHeight(grid(Index).Rows - 1) = TextHeight(WordWrap(txtComments(Index), 185))
70    txtComments(Index) = ""
      'cmdAddComments(Index).Enabled = False

80    Exit Sub

cmdAddComments_Click_Error:

       Dim strES As String
       Dim intEL As Integer

90     intEL = Erl
100    strES = Err.Description
110    LogError "frmPatientNotePad", "cmdAddComments_Click", intEL, strES, sql
          
End Sub

Private Sub cmdClose_Click()
10    On Error GoTo cmdClose_Click_Error

20    Unload Me

30    Exit Sub

cmdClose_Click_Error:

       Dim strES As String
       Dim intEL As Integer

40     intEL = Erl
50     strES = Err.Description
60     LogError "frmPatientNotePad", "cmdClose_Click", intEL, strES
          
End Sub

Private Function WordWrap(strFullText As String, intLength As Integer) As String

      Dim intLen As Integer, intCr As Integer, intSpace As Integer
      Dim strText As String, strNextLine As String
      Dim blnDoneOnce As Boolean

10    On Error GoTo WordWrap_Error

20    intLength = intLength + 1
30    strFullText = Trim$(strFullText)

40    Do
50        intLen = Len(strNextLine)
60        intSpace = InStr(strFullText, " ")
70        intCr = InStr(strFullText, vbCr)

80        If intCr Then
90            If intLen + intCr <= intLength Then
100               strText = strText & strNextLine & Left$(strFullText, intCr)
110               strNextLine = ""
120               strFullText = Mid$(strFullText, intCr + 1)
130               GoTo LoopHere
140           End If
150       End If

160       If intSpace Then
170           If intLen + intSpace <= intLength Then
180               blnDoneOnce = True
190               strNextLine = strNextLine & Left$(strFullText, intSpace)
200               strFullText = Mid$(strFullText, intSpace + 1)
210           ElseIf intSpace > intLength Then
220               strText = strText & vbCrLf & Left$(strFullText, intLength)
230               strFullText = Mid$(strFullText, intLength + 1)
240           Else
250               strText = strText & strNextLine & vbCrLf
260               strNextLine = ""
270           End If
280       Else
290           If intLen Then
300               If intLen + Len(strFullText) > intLength Then
310                   strText = strText & strNextLine & vbCrLf & strFullText & vbCrLf
320               Else
330                   strText = strText & strNextLine & strFullText & vbCrLf
340               End If
350           Else
360               strText = strText & strFullText & vbCrLf
370           End If
380           Exit Do
390       End If

LoopHere:
400   Loop

410   WordWrap = strText

420   Exit Function

WordWrap_Error:

       Dim strES As String
       Dim intEL As Integer

430    intEL = Erl
440    strES = Err.Description
450    LogError "frmPatientNotePad", "WordWrap", intEL, strES
          

End Function

Private Sub LoadComments()



10    On Error GoTo LoadComments_Error

20    If Val(GLabNo) = 0 Then
30        If Caller = "Microbiology" Then
40            LoadDescipline Val(SampleID), SSTab.TabCaption(SSTab.Tab), 3
50        Else
60            LoadDescipline Val(SampleID), SSTab.TabCaption(0), 0
70            LoadDescipline Val(SampleID), SSTab.TabCaption(1), 1
80            LoadDescipline Val(SampleID), SSTab.TabCaption(2), 2
90        End If
100   Else
110       LoadDescipline Val(SampleID), SSTab.TabCaption(0), 0, GLabNo
120       LoadDescipline Val(SampleID), SSTab.TabCaption(1), 1, GLabNo
130       LoadDescipline Val(SampleID), SSTab.TabCaption(2), 2, GLabNo
140       LoadDescipline Val(SampleID), SSTab.TabCaption(3), 3, GLabNo
150   End If

160   Exit Sub

LoadComments_Error:

       Dim strES As String
       Dim intEL As Integer

170    intEL = Erl
180    strES = Err.Description
190    LogError "frmPatientNotePad", "LoadComments", intEL, strES
          
End Sub

Private Sub LoadDemo(SampleID As String)
    Dim sql As String
    Dim tb As Recordset

    On Error GoTo LoadDemo_Error

    sql = "Select * from Demographics where " & _
          "SampleID = '" & SampleID & "'"

    Set tb = New Recordset
    RecOpenClient 0, tb, sql
    With tb
        If Not tb.EOF Then
            lblSampleIDD = !SampleID
            lblChartNoD = !Chart & ""
            lblSurNameD = SurName(!PatName & "")
            lblForeNameD = ForeName(!PatName & "")
            lblAddressD = !Addr0 & " " & !Addr1 & ""
            lblSexD = !sex & ""
            lblDOBD = !Dob
            lblAgeD = !Age
            lblSampleDateD = !SampleDate
            lblDemoDateD = !DateTimeDemographics
            GLabNo = !Chart & ""
        Else: lblSampleIDD = ""
        End If
    End With

    Exit Sub

LoadDemo_Error:

    Dim strES As String
    Dim intEL As Integer

    intEL = Erl
    strES = Err.Description
    LogError "frmPatientNotePad", "LoadDemo", intEL, strES, sql

End Sub
Private Sub LoadDescipline(SampleID As String, Descipline As String, Index As Integer, Optional LabNo As String = "")
          Dim sql As String
          Dim tb As Recordset
          Dim LNo As Long

          'cmdEdit(Index).Enabled = False
          'cmdDelete(Index).Enabled = False
          'cmdAddComments(Index).Enabled = False
10        On Error GoTo LoadDescipline_Error




60        grid(Index).Clear
70        grid(Index).Rows = 1

80        grid(Index).SelectionMode = flexSelectionByRow


90        If LabNo = "" Then
100           sql = "Select * from PatientNotePad where " & _
                    "SampleID = '" & SampleID & "' and " & _
                    "Descipline = '" & Descipline & "'"
110       Else
120           sql = "Select * from PatientNotePad where " & _
                    "labNo = '" & LabNo & "' and " & _
                    "Descipline = '" & Descipline & "'"
130       End If
140       Set tb = New Recordset
150       RecOpenClient 0, tb, sql
160       With tb
170           If tb.EOF Then
180           Else
190               Do Until .EOF
200                   grid(Index).AddItem Format$(!DateTimeOfRecord, "dd/mmm/yyyy hh:mm:ss") & vbTab & WordWrap(!Comment, 185) & vbTab & !UserName
210                   grid(Index).RowHeight(grid(Index).Rows - 1) = TextHeight(WordWrap(!Comment, 185))
220                   .MoveNext
230               Loop
240           End If
250       End With

260       Exit Sub

LoadDescipline_Error:

          Dim strES As String
          Dim intEL As Integer

270       intEL = Erl
280       strES = Err.Description
290       LogError "frmPatientNotePad", "LoadDescipline", intEL, strES, sql


End Sub

Private Sub cmdDelete_Click(Index As Integer)
      Dim sql As String

10    On Error GoTo cmdDelete_Click_Error

20    If grid(Index).row = 0 Then Exit Sub

30    If iMsg("Do you want to delete selected comment", vbYesNo) = vbNo Then Exit Sub
          

40    sql = "delete from PatientNotePad where Sampleid = '" & SampleID & "' and DateTimeofRecord = '" & grid(Index).TextMatrix(grid(Index).row, 0) & "'"
50    Cnxn(0).Execute sql

60    If grid(Index).Rows = 2 Then
70        grid(Index).Clear
80        grid(Index).Rows = 1
90    Else
100       grid(Index).RemoveItem (grid(Index).row)
110   End If
120   grid(Index).row = 0
      'cmdDelete(Index).Enabled = False
      'cmdEdit(Index).Enabled = False

130   Exit Sub

cmdDelete_Click_Error:

       Dim strES As String
       Dim intEL As Integer

140    intEL = Erl
150    strES = Err.Description
160    LogError "frmPatientNotePad", "cmdDelete_Click", intEL, strES
          

End Sub

Private Sub cmdEdit_Click(Index As Integer)
10    On Error GoTo cmdEdit_Click_Error

20    txtComments(Index) = grid(Index).TextMatrix(grid(Index).row, 1)
30    If grid(Index).Rows = 2 Then
40        grid(Index).Clear
50        grid(Index).Rows = 1
60    Else
70        grid(Index).RemoveItem (grid(Index).row)
80    End If
90    grid(Index).row = 0
      'cmdEdit(Index).Enabled = False
      'cmdDelete(Index).Enabled = False
100   txtComments(Index).SetFocus

110   Exit Sub

cmdEdit_Click_Error:

       Dim strES As String
       Dim intEL As Integer

120    intEL = Erl
130    strES = Err.Description
140    LogError "frmPatientNotePad", "cmdEdit_Click", intEL, strES
          
End Sub

Private Sub Form_Activate()

    If lblSampleIDD = "" Then
        Unload Me
     End If
     
    End Sub

Private Sub Form_Load()
          Dim sql As String
10        On Error GoTo Form_Load_Error

20        SSTab.TabVisible(1) = False
30        SSTab.TabVisible(2) = False
40        For i = 0 To 3
50            grid(i).Clear
60            grid(i).Rows = 1
70        Next i

80        CheckPatientNotepadInDb

90        If Caller = "Microbiology" Then
100           SampleID = Val(SampleID) + SysOptMicroOffset(0)
110           SSTab.Tab = 3
              'Disable other discipline changes
120           LockForm True
130       Else
140           SSTab.Tab = 0
150           LockForm False
160       End If

170       LoadDemo SampleID
180       LoadComments



190       Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

200       intEL = Erl
210       strES = Err.Description
220       LogError "frmPatientNotePad", "Form_Load", intEL, strES, sql

End Sub



Private Sub LockForm(ByVal EnableMicro As Boolean)

      Dim i As Integer

10    On Error GoTo LockControls_Error


20    For i = 0 To 2
30        cmdAddComments(i).Enabled = Not EnableMicro
40        cmdEdit(i).Enabled = Not EnableMicro
50        cmdDelete(i).Enabled = Not EnableMicro
60    Next i
70    cmdAddComments(3).Enabled = EnableMicro
80    cmdEdit(3).Enabled = EnableMicro
90    cmdDelete(3).Enabled = EnableMicro

100   Exit Sub

LockControls_Error:

       Dim strES As String
       Dim intEL As Integer

110    intEL = Erl
120    strES = Err.Description
130    LogError "frmPatientNotePad", "LockForm", intEL, strES

End Sub

Private Sub grid_Click(Index As Integer)
      Dim CurRow As Integer
10    On Error GoTo grid_Click_Error

20    CurRow = grid(Index).MouseRow
      'If CurRow > 0 Then
      '    cmdEdit(Index).Enabled = True
      '    cmdDelete(Index).Enabled = True
      'End If

30    Exit Sub

grid_Click_Error:

       Dim strES As String
       Dim intEL As Integer

40     intEL = Erl
50     strES = Err.Description
60     LogError "frmPatientNotePad", "grid_Click", intEL, strES
          
End Sub

Private Sub SSTab_Click(PreviousTab As Integer)
10    On Error GoTo SSTab_Click_Error

20    LoadComments

30    Exit Sub

SSTab_Click_Error:

       Dim strES As String
       Dim intEL As Integer

40     intEL = Erl
50     strES = Err.Description
60     LogError "frmPatientNotePad", "SSTab_Click", intEL, strES
          
End Sub

'Private Sub txtComments_Change(Index As Integer)
'cmdAddComments(Index).Enabled = Len(txtComments(Index))
'End Sub
