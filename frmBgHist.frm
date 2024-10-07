VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmBGHist 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Blood Gas History"
   ClientHeight    =   4650
   ClientLeft      =   225
   ClientTop       =   960
   ClientWidth     =   7425
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "frmBgHist.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4650
   ScaleWidth      =   7425
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   2985
      Picture         =   "frmBgHist.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3660
      Width           =   1245
   End
   Begin MSFlexGridLib.MSFlexGrid grdBga 
      Height          =   2925
      Left            =   90
      TabIndex        =   9
      Top             =   675
      Width           =   7125
      _ExtentX        =   12568
      _ExtentY        =   5159
      _Version        =   393216
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "&View Result"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   1635
      Picture         =   "frmBgHist.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3660
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   4365
      MaskColor       =   &H80000000&
      Picture         =   "frmBgHist.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3690
      UseMaskColor    =   -1  'True
      Width           =   1245
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "NOPAS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6330
      TabIndex        =   7
      Top             =   120
      Width           =   555
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "MRN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4950
      TabIndex        =   6
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1770
      TabIndex        =   5
      Top             =   120
      Width           =   420
   End
   Begin VB.Label lblNopas 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   6300
      TabIndex        =   4
      Top             =   330
      Width           =   1035
   End
   Begin VB.Label lblMrn 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   4920
      TabIndex        =   3
      Top             =   330
      Width           =   1035
   End
   Begin VB.Label lblName 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   1710
      TabIndex        =   2
      Top             =   330
      Width           =   2775
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cumulative Report for:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "frmBGHist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit  '© Custom Software 2001

Private Sub cmdCancel_Click()

Unload Me

End Sub

Private Sub cmdPrint_Click()

Dim Y As Long
Dim X As Long

Printer.Font.Name = "Courier New"
Printer.Font.Size = 10

For X = 0 To grdBga.Cols - 1
  For Y = 0 To grdBga.Rows - 1
    If Y > 4 Then Exit For
    Printer.Print Tab(Choose(Y + 1, 1, 10, 30, 50, 70, 90, 110));
    Printer.Print grdBga.TextMatrix(Y, X);
  Next
  Printer.Print
Next
Printer.EndDoc

End Sub

Private Sub cmdView_Click()

If grdBga.Row < 1 Then Exit Sub

If Trim(grdBga.TextMatrix(grdBga.Row, 2)) = "" Then Exit Sub

If frmWardViewBloodGas.Visible = True Then Exit Sub
With frmWardViewBloodGas
  .cmbNo = grdBga.TextMatrix(grdBga.Row, 1)
  .lNOPAS = lblNOPAS
  .lMRN = lblMrn
  .lName = lblName
  .cmdCum.Enabled = False
  .Show 1
End With


End Sub

Private Sub FillG()

Dim sn As New Recordset
Dim SQL As String
Dim s As String



grdBga.Rows = 2
grdBga.AddItem ""
grdBga.RemoveItem 1
  

SQL = "SELECT * from BGAresults WHERE " & _
      "nopas = '" & lblNOPAS & "' " & _
      "ORDER BY rundatetime DESC"

Set sn = New Recordset
RecOpenServer 0, sn, SQL
Do While Not sn.EOF
  s = Format(sn!RunDateTime, "dd/MM/yyyy hh:mm:ss") & vbTab & _
      sn!SampleID & "" & vbTab & _
      Format(sn!pH, "0.00") & vbTab & _
      Format(sn!PCO2, "#0.0") & vbTab & _
      Format(sn!PO2, "#0.0") & vbTab & _
      Format(sn!HCO3, "####") & vbTab & _
      Format(sn!O2SAT, "#0.0") & vbTab & _
      Format(sn!BE, "#0.0") & vbTab & _
      Format(sn!TotCO2, "#0.0")
  grdBga.AddItem s
  sn.MoveNext
Loop

If grdBga.Rows > 2 Then
  grdBga.RemoveItem 1
End If


End Sub

Private Sub Form_Activate()

FillG

End Sub

Private Sub Form_Load()

grdBga.FormatString = "<Date/Time  |<Run   |<Ph  |<PCO2  |" & _
                 "<PO2  |<HCO3 |<O2SAT|<BE   |<TOT CO2"

End Sub

