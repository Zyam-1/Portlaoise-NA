VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H80000014&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4245
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000014&
      Height          =   4050
      Left            =   135
      TabIndex        =   0
      Top             =   60
      Width           =   7080
      Begin VB.Image Image1 
         Height          =   1515
         Left            =   60
         Picture         =   "frmSplash.frx":0ECA
         Stretch         =   -1  'True
         Top             =   3420
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblCopyright 
         BackColor       =   &H00FFFFFF&
         Caption         =   "© Copyright Custom Software"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4620
         TabIndex        =   2
         Top             =   3720
         Width           =   2415
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5910
         TabIndex        =   3
         Top             =   3300
         Width           =   885
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         Caption         =   "Windows NT/Citrix/Windows XP"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2070
         TabIndex        =   4
         Top             =   2940
         Width           =   4725
      End
      Begin VB.Label lblProductName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Product"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   330
         Left            =   3840
         TabIndex        =   5
         Top             =   2100
         Width           =   1110
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Licensed to "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   6855
      End
      Begin VB.Image Image2 
         Height          =   2955
         Left            =   2100
         Picture         =   "frmSplash.frx":11D4
         Stretch         =   -1  'True
         Top             =   180
         Width           =   3105
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)

10        Unload Me

End Sub

Private Sub Form_Load()

10        CheckIDE

20        lblVersion.Caption = "Version " & App.Major & "." & App.Minor
          'Zyam added app revision 12-22-23
30        lblProductName.Caption = App.Major & "." & App.Minor & "." & App.Revision
          'Zyam
40        lblLicenseTo.Caption = lblLicenseTo.Caption & " " & HospName(0)
50        Me.Show     ' Display startup form.
60        DoEvents    ' Ensure startup form is painted.


70        Load frmMain  ' Load main application fom.
80        Unload Me   ' Unload startup form.
90        frmMain.Show  ' Display main form.

100       Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer



110       intEL = Erl
120       strES = Err.Description
130       LogError "frmSplash", "Form_Load", intEL, strES


End Sub

Private Sub Frame1_Click()

10        Unload Me

End Sub


