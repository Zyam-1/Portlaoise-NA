VERSION 5.00
Begin VB.Form frmPanelsBarcodeSelection 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetAquire-Panels Barcode Selection"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   5265
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   975
      Left            =   4005
      Picture         =   "frmPanelsBarcodeSelection.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3015
      Width           =   1155
   End
   Begin VB.CommandButton cmdIMM 
      Caption         =   "Immunology"
      Height          =   1275
      Left            =   855
      TabIndex        =   2
      Top             =   2700
      Width           =   1815
   End
   Begin VB.CommandButton cmdEndo 
      Caption         =   "Endocrinology"
      Height          =   1275
      Left            =   855
      TabIndex        =   1
      Top             =   1395
      Width           =   1815
   End
   Begin VB.CommandButton cmdBio 
      Caption         =   "Biochemistry"
      Height          =   1275
      Left            =   855
      TabIndex        =   0
      Top             =   90
      Width           =   1815
   End
End
Attribute VB_Name = "frmPanelsBarcodeSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBio_Click()
On Error GoTo cmdBio_Click_Error

frmPanelBarCodes.Tabelname = "PANELS"
frmPanelBarCodes.frmHeading = "Biochemestry"
frmPanelBarCodes.Show 1

Exit Sub

cmdBio_Click_Error:
Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "frmPanelsBarcodeSelection", "cmdBio_Click", intEL, strES

End Sub

Private Sub cmdCancel_Click()
10    Unload Me
End Sub

Private Sub cmdEndo_Click()
10    On Error GoTo cmdEndo_Click_Error

20    frmPanelBarCodes.Tabelname = "ENDPANELS"
30    frmPanelBarCodes.frmHeading = "Endocrinology"
40    frmPanelBarCodes.Show 1

50    Exit Sub

cmdEndo_Click_Error:
      Dim strES As String
      Dim intEL As Integer

60    intEL = Erl
70    strES = Err.Description
80    LogError "frmPanelsBarcodeSelection", "cmdEndo_Click", intEL, strES

End Sub

Private Sub cmdIMM_Click()
On Error GoTo cmdIMM_Click_Error

frmPanelBarCodes.Tabelname = "IPANELS"
frmPanelBarCodes.frmHeading = "Immunology"
frmPanelBarCodes.Show 1

Exit Sub

cmdIMM_Click_Error:
Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "frmPanelsBarcodeSelection", "cmdIMM_Click", intEL, strES

End Sub
