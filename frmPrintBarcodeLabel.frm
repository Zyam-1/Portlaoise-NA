VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form frmPrintBarcodeLabel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3630
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   6750
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optCopies 
      Caption         =   "1"
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
      Index           =   1
      Left            =   2940
      TabIndex        =   16
      Top             =   330
      Value           =   -1  'True
      Width           =   525
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   1100
      Left            =   5130
      Picture         =   "frmPrintBarcodeLabel.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Tag             =   "101"
      ToolTipText     =   "Exit Screen"
      Top             =   630
      Width           =   1200
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   1100
      Left            =   3660
      Picture         =   "frmPrintBarcodeLabel.frx":049C
      Style           =   1  'Graphical
      TabIndex        =   7
      Tag             =   "bprint"
      ToolTipText     =   "Print Result"
      Top             =   630
      Width           =   1200
   End
   Begin VB.OptionButton optCopies 
      Caption         =   "6"
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
      Index           =   6
      Left            =   2940
      TabIndex        =   6
      Top             =   1530
      Width           =   525
   End
   Begin VB.OptionButton optCopies 
      Caption         =   "5"
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
      Index           =   5
      Left            =   2940
      TabIndex        =   5
      Top             =   1290
      Width           =   525
   End
   Begin VB.OptionButton optCopies 
      Caption         =   "4"
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
      Index           =   4
      Left            =   2940
      TabIndex        =   4
      Top             =   1050
      Width           =   525
   End
   Begin VB.OptionButton optCopies 
      Caption         =   "3"
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
      Index           =   3
      Left            =   2940
      TabIndex        =   3
      Top             =   810
      Width           =   525
   End
   Begin VB.OptionButton optCopies 
      Caption         =   "2"
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
      Index           =   2
      Left            =   2940
      TabIndex        =   2
      Top             =   570
      Width           =   525
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   6750
      _ExtentX        =   11906
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lblDepartment 
      Height          =   315
      Left            =   5400
      TabIndex        =   17
      Top             =   2220
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Label lblAge 
      Alignment       =   2  'Center
      Caption         =   "lblTestOrderTitle"
      Height          =   195
      Left            =   180
      TabIndex        =   15
      Top             =   2340
      Width           =   2535
   End
   Begin VB.Label lblAgeSexDoB 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   180
      TabIndex        =   14
      Top             =   2610
      Width           =   6390
   End
   Begin VB.Label lblSampleDateTitle 
      Alignment       =   2  'Center
      Caption         =   "lblSampleDate"
      Height          =   195
      Left            =   180
      TabIndex        =   13
      Top             =   1590
      Width           =   2535
   End
   Begin VB.Label lblSampleDate 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblSampleDate"
      Height          =   375
      Left            =   180
      TabIndex        =   12
      Top             =   1860
      Width           =   2565
   End
   Begin VB.Label lblSurnameTitle 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   195
      Left            =   180
      TabIndex        =   11
      Top             =   1020
      Width           =   2535
   End
   Begin VB.Label lblPatName 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   150
      TabIndex        =   10
      Top             =   1230
      Width           =   2565
   End
   Begin VB.Label lblSampleIDTitle 
      Alignment       =   2  'Center
      Caption         =   "Sample ID"
      Height          =   195
      Left            =   150
      TabIndex        =   1
      Top             =   285
      Width           =   2565
   End
   Begin VB.Label lblSampleID 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   150
      TabIndex        =   0
      Top             =   495
      Width           =   2565
   End
   Begin VB.Menu mnuOffset 
      Caption         =   "Offset"
   End
   Begin VB.Menu mnuSetPrinter 
      Caption         =   "Set Printer"
   End
End
Attribute VB_Name = "frmPrintBarcodeLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private OffSet As Integer

Private Sub cmdCancel_Click()

10        Unload Me

End Sub

Private Sub cmdPrint_Click()

      Dim Copies As Integer
      Dim n As Integer
      Dim TargetPrinter As String
      Dim Found As Boolean
      Dim Px As Printer
      Dim ComputerName As String
      Dim OffsetCounter As Integer
      Dim intBarcodeFontSize As Integer

10    On Error GoTo cmdPrint_Click_Error

20    ComputerName = vbGetComputerName()
30    TargetPrinter = GetOptionSetting("BarcodePrinter" & lblDepartment, "")

40    If TargetPrinter = "" Then
50        iMsg "Please select barcode printer first", vbOKOnly
60        Exit Sub
70    End If

80    For n = 1 To 6
90        If optCopies(n) = True Then
100           Copies = n
110           Exit For
120       End If
130   Next



140   intBarcodeFontSize = GetOptionSetting("BarCodeLabelFontSize", 24)

150   Found = False
160   For Each Px In Printers
170       If UCase$(Px.DeviceName) = UCase$(TargetPrinter) Then
180           Set Printer = Px
190           Found = True
200           Exit For
210       End If
220   Next
230   If Not Found Then
240       iMsg "No Printer installed on pc"    ' LS(csNoPrinterInstalledonPc)
          'If TimedOut Then Exit Sub
250       Exit Sub
260   End If

270   For n = 1 To Copies
280       Printer.FontName = "Courier New"
290       Printer.Font.Size = 8
300       Printer.Print " "

310       For OffsetCounter = 1 To OffSet
320           Printer.Print
330       Next


          ' print barcode in first row
          'Printer.CurrentX = 1
340       Printer.Font.Name = "Free 3 of 9 Extended"    'FRE3OF9X.TTF"
350       Printer.Font.Size = intBarcodeFontSize
360       Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth("*" & lblSampleID & "*")) / 2
370       Printer.Print "*" & lblSampleID & "*"

          'Printer.CurrentX = 1
380       Printer.FontName = "Courier New"
390       Printer.Font.Size = 8
400       Printer.Font.Bold = True
          'print sampleid in 2nd row
410       Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(lblSampleID)) / 2
420       Printer.Print lblSampleID
          'print PatientName  in 3rd row
430       Printer.Font.Size = 6
440       Printer.Font.Bold = False
450       Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(lblPatName)) / 2
460       Printer.CurrentX = 3
470       Printer.Print Trim(lblPatName)
          'print Patient DOB in 4th row
480       Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth("A/S/DOB: 21 M 20150215")) / 2
490       Printer.CurrentX = 3
500       Printer.Print lblAgeSexDoB
          'print Sample Date in 5th row
510       Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth("Date:" & lblSampledate & " " & Format(Time, "hh:mm:ss"))) / 2
520       Printer.CurrentX = 3
530       Printer.Print "Date:" & lblSampledate
          'print Hospital Name in 6th row
540       Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth("Hospital Name")) / 2
550       Printer.CurrentX = 3
560       Printer.Print HospName(0) & " Hospital"

          '        If Len(lblTestOrder) > 29 Then
          '            Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(Left(lblTestOrder, InStrRev(Left(lblTestOrder, 29), ",")))) / 2
          '            Printer.Print Left(lblTestOrder, InStrRev(Left(lblTestOrder, 29), ",") - 1)
          '        Else
          '            Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(lblTestOrder)) / 2
          '            Printer.Print lblTestOrder
          '        End If

570       Printer.EndDoc
580   Next

590   Exit Sub

cmdPrint_Click_Error:

      Dim strES As String
      Dim intEL As Integer

600   intEL = Erl
610   strES = Err.Description
620   LogError "frmPrintBarcodeLabel", "cmdPrint_Click", intEL, strES

End Sub

Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        Me.Caption = "NetAcquire " & "Barcode"
30        lblSampleIDTitle.Caption = "Sampleid"
40        lblSurnameTitle.Caption = "Patient Name "
50        lblSampleDateTitle.Caption = "Sample Time"
60        lblAgeSexDoB.Caption = "Age/Sex/DoB"
70        cmdPrint.Caption = "Print"
80        cmdCancel.Caption = "Exit"
90        cmdPrint.ToolTipText = "Print Barcode"

100       mnuOffset.Visible = False
110       If UserMemberOf = "Managers" Then
120           mnuOffset.Visible = True
130       End If

140       OffSet = Val(GetOptionSetting("BarCodeLabelPrintOffset", "1"))

150       Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

160       intEL = Erl
170       strES = Err.Description
180       LogError "frmPrintBarcodeLabel", "Form_Load", intEL, strES

End Sub


Private Sub mnuOffset_Click()

          Dim NewOffset As Integer

10        NewOffset = Val(iBOX("Offset 0/1/2/3 ?", , CStr(OffSet)))

20        If Val(NewOffset) > -1 And Val(NewOffset) < 4 Then
30            OffSet = Int(NewOffset)
40            SaveOptionSetting "BarCodeLabelPrintOffset", OffSet
50        End If

End Sub


Private Sub mnuSetPrinter_Click()

      Dim PrinterName As String
10    On Error GoTo mnuSetPrinter_Click_Error

20    PrinterName = iBOX("Please enter printer path for " & lblDepartment, , GetOptionSetting("BarcodePrinter" & lblDepartment, ""))
30    If PrinterName <> "" Then
40        SaveOptionSetting "BarcodePrinter" & lblDepartment, PrinterName
50    End If

60    Exit Sub
mnuSetPrinter_Click_Error:
         
70    LogError "frmPrintBarcodeLabel", "mnuSetPrinter_Click", Erl, Err.Description
End Sub
