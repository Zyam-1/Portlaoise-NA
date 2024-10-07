VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEnterDate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5265
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComCtl2.DTPicker dtTime 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "HH:mm"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   6153
         SubFormatType   =   4
      EndProperty
      Height          =   315
      Left            =   3810
      TabIndex        =   2
      Top             =   1290
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "HH:mm"
      Format          =   272236547
      UpDown          =   -1  'True
      CurrentDate     =   39682
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   705
      Left            =   3780
      Picture         =   "frmEnterDate.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1950
      Width           =   1275
   End
   Begin MSComCtl2.MonthView cal 
      Height          =   2370
      Left            =   330
      TabIndex        =   1
      Top             =   270
      Width           =   3120
      _ExtentX        =   5503
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      StartOfWeek     =   272236546
      CurrentDate     =   39682
   End
End
Attribute VB_Name = "frmEnterDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pDateVal As String

Private Sub cal_DateClick(ByVal DateClicked As Date)

10        pDateVal = DateClicked & " " & dtTime
20        cmdSave.SetFocus

End Sub

Private Sub cmdSave_Click()

10        Me.Hide

End Sub


Private Sub dtTime_Change()

10        pDateVal = cal.Value & " " & dtTime
20        cmdSave.SetFocus

End Sub


Private Sub Form_Activate()

10        If IsDate(pDateVal) Then
20            cal.Value = Format(pDateVal, "dd/MMM/yyyy")
30            dtTime.Value = Format(pDateVal, "HH:mm")
40        Else
50            cal.Value = Format(Now, "dd/MMM/yyyy")
60            dtTime.Value = Format(Now, "HH:mm")
70        End If

End Sub

Private Sub Form_Load()

10        cal.Value = Now

End Sub



Public Property Get DateVal() As String

10        DateVal = pDateVal

End Property

Public Property Let DateVal(ByVal sNewValue As String)

10        pDateVal = sNewValue

End Property
