VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   3555
   ClientLeft      =   2310
   ClientTop       =   2115
   ClientWidth     =   6030
   ClipControls    =   0   'False
   DrawMode        =   1  'Blackness
   DrawStyle       =   5  'Transparent
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Palette         =   "frmAbout.frx":030A
   PaletteMode     =   2  'Custom
   ScaleHeight     =   2453.724
   ScaleMode       =   0  'User
   ScaleWidth      =   5662.482
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1800
      TabIndex        =   6
      Top             =   1740
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Default         =   -1  'True
      Height          =   570
      Left            =   4410
      Picture         =   "frmAbout.frx":1D50
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2295
      Width           =   1260
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&System Info..."
      Height          =   570
      Left            =   4410
      TabIndex        =   1
      Top             =   2925
      Width           =   1245
   End
   Begin VB.Image Image1 
      Height          =   1035
      Left            =   0
      Picture         =   "frmAbout.frx":2C1A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2055
   End
   Begin VB.Label lblDisclaimer 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmAbout.frx":3B45
      ForeColor       =   &H00000000&
      Height          =   825
      Left            =   270
      TabIndex        =   5
      Top             =   2610
      Width           =   3870
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   211.287
      X2              =   5436.17
      Y1              =   1552.99
      Y2              =   1552.99
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Laboratory Information System."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   585
      Left            =   45
      TabIndex        =   2
      Top             =   1035
      Width           =   5865
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Application Title"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   2520
      TabIndex        =   3
      Top             =   180
      Width           =   3390
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   126.772
      X2              =   5606.139
      Y1              =   1552.99
      Y2              =   1563.343
   End
   Begin VB.Label lblVersion 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Version"
      Height          =   225
      Left            =   2520
      TabIndex        =   4
      Top             =   720
      Width           =   3390
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
      KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
      KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL

' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number

Const gREGKEYSysINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSysINFOLOC = "MSINFO"
Const gREGKEYSysINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSysINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Private Sub cmdOK_Click()

10        Unload Me

End Sub

Private Sub cmdSysInfo_Click()
10        Call StartSysInfo
End Sub



Private Sub Command1_Click()

Dim tb As Recordset
Dim sql As String

On Error GoTo Command1_Click_Error

sql = "SELECT * FROM SampleIds"
Set tb = New Recordset
RecOpenServer 0, tb, sql
While Not tb.EOF
    FixDefIndex tb!SID & ""
    tb.MoveNext
Wend

Exit Sub

Command1_Click_Error:
Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "frmAbout", "Command1_Click", intEL, strES

End Sub

Private Sub FixDefIndex(ByVal SampleID As String)

Dim tb As Recordset
Dim sql As String
Dim Obs As New Observations
Dim Ward As String
Dim Clinician As String
Dim GP As String

On Error GoTo FixDefIndex_Error

sql = "SELECT * FROM Demographics WHERE SampleID = " & SampleID
Set tb = New Recordset
RecOpenServer 0, tb, sql
If Not tb.EOF Then

    sql = "UPDATE BioResults SET DefIndex = Null, Printed = 0 WHERE SampleID = " & SampleID & " AND Code = '599'"
    Cnxn(0).Execute sql
    
    Obs.Save SampleID, False, _
             "Biochemistry", "Amended report. Amended GGT ref range"

    sql = "INSERT INTO [dbo].[PrintPending] " & _
        "([SampleID], [Department], [Initiator], [ptime], [Ward], [Clinician], [GP] )" & _
        " VALUES " & _
        "('" & SampleID & "', 'B', 'AV', '" & Format(Now, "dd/MMM/yyyy HH:mm:ss") & "', '" & AddTicks(tb!Ward & "") & "', '" & AddTicks(tb!Clinician & "") & "', '" & AddTicks(tb!GP & "") & "') "

    Cnxn(0).Execute sql


End If
Exit Sub

FixDefIndex_Error:
Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "frmAbout", "FixDefIndex", intEL, strES

End Sub


Private Sub Form_Load()
10        Me.Caption = "About " & App.Title
20        lblVersion.Caption = "Version " & App.Major & "." & App.Minor
30        lblTitle.Caption = App.Title
40        SetFormStyle Me
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
          Dim i As Long                                           ' Loop Counter
          Dim rc As Long                                          ' Return Code
          Dim hKey As Long                                        ' Handle To An Open Registry Key
          Dim KeyValType As Long                                  ' Data Type Of A Registry Key
          Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
          Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
      '------------------------------------------------------------
      ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
      '------------------------------------------------------------
10        rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey)    ' Open Registry Key

20        If (rc <> ERROR_SUCCESS) Then
30            KeyVal = ""                                             ' Set Return Val To Empty String
40            GetKeyValue = False                                     ' Return Failure
50            rc = RegCloseKey(hKey)                                  ' Close Registry Key
60            Exit Function
70        End If

80        tmpVal = String$(1024, 0)                             ' Allocate Variable Space
90        KeyValSize = 1024                                       ' Mark Variable Size

          '------------------------------------------------------------
          ' Retrieve Registry Key Value...
          '------------------------------------------------------------
100       rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                               KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value

110       If (rc <> ERROR_SUCCESS) Then          ' Handle Errors
120           KeyVal = ""                                             ' Set Return Val To Empty String
130           GetKeyValue = False                                     ' Return Failure
140           rc = RegCloseKey(hKey)                                  ' Close Registry Key
150           Exit Function
160       End If

170       If (Asc(Mid$(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
180           tmpVal = Left$(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
190       Else                                                    ' WinNT Does NOT Null Terminate String...
200           tmpVal = Left$(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
210       End If
          '------------------------------------------------------------
          ' Determine Key Value Type For Conversion...
          '------------------------------------------------------------
220       Select Case KeyValType                                  ' Search Data Types...
          Case REG_SZ                                             ' String Registry Key Data Type
230           KeyVal = tmpVal                                     ' Copy String Value
240       Case REG_DWORD                                          ' Double Word Registry Key Data Type
250           For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
260               KeyVal = KeyVal + Hex(Asc(Mid$(tmpVal, i, 1)))   ' Build Value Char. By Char.
270           Next
280           KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
290       End Select

300       GetKeyValue = True                                      ' Return Success
310       rc = RegCloseKey(hKey)                                  ' Close Registry Key

End Function

Public Sub StartSysInfo()
10        On Error GoTo SysInfoErr

          Dim rc As Long
          Dim SysInfoPath As String

          ' Try To Get System Info Program Path\Name From Registry...
20        If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSysINFO, gREGVALSysINFO, SysInfoPath) Then
              ' Try To Get System Info Program Path Only From Registry...
30        ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSysINFOLOC, gREGVALSysINFOLOC, SysInfoPath) Then
              ' Validate Existance Of Known 32 Bit File Version
40            If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
50                SysInfoPath = SysInfoPath & "\MSINFO32.EXE"

                  ' Error - File Can Not Be Found...
60            Else
70                GoTo SysInfoErr
80            End If
              ' Error - Registry Entry Can Not Be Found...
90        Else
100           GoTo SysInfoErr
110       End If

120       Call Shell(SysInfoPath, vbNormalFocus)

130       Exit Sub
SysInfoErr:
140       iMsg "System Information Is Unavailable At This Time", vbInformation
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
      Dim ShiftTest As Integer
10    ShiftTest = Shift And 7
20    Select Case ShiftTest
      Case 1    ' or vbShiftMask
          'Print "You pressed the SHIFT key."
30    Case 2    ' or vbCtrlMask
          Print "You pressed the CTRL key."
40        frmSystemFunctions.Show 1
50    Case 4    ' or vbAltMask
          'Print "You pressed the ALT key."
60    Case 3
          'Print "You pressed both SHIFT and CTRL."
70    Case 5
          'Print "You pressed both SHIFT and ALT."
80    Case 6
          'Print "You pressed both CTRL and ALT."
90    Case 7
          'Print "You pressed SHIFT, CTRL, and ALT."
100   End Select
End Sub
