VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "mschrt20.ocx"
Begin VB.Form frmHaemGraphs 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetAcquire - Haematology Graphs"
   ClientHeight    =   8940
   ClientLeft      =   45
   ClientTop       =   6285
   ClientWidth     =   6690
   Icon            =   "frmHaemGraphs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8940
   ScaleWidth      =   6690
   Begin MSChart20Lib.MSChart gRBC 
      Height          =   2955
      Left            =   1110
      OleObjectBlob   =   "frmHaemGraphs.frx":030A
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   -30
      Visible         =   0   'False
      Width           =   4275
   End
   Begin MSChart20Lib.MSChart gWBC 
      Height          =   2955
      Left            =   1080
      OleObjectBlob   =   "frmHaemGraphs.frx":1BCC
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2910
      Visible         =   0   'False
      Width           =   4305
   End
   Begin MSChart20Lib.MSChart gPla 
      Height          =   2955
      Left            =   1050
      OleObjectBlob   =   "frmHaemGraphs.frx":3B9C
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5850
      Visible         =   0   'False
      Width           =   4365
   End
   Begin VB.Image Image1 
      Height          =   6795
      Left            =   90
      Stretch         =   -1  'True
      Top             =   180
      Visible         =   0   'False
      Width           =   6495
   End
End
Attribute VB_Name = "frmHaemGraphs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mSampleID As String
Private Sub LoadResults()

          Dim tb As New Recordset
          Dim sql As String
          Dim n As Long
          Dim Img As String

          Dim gDataRBC(1 To 64, 1 To 1) As Variant
          Dim gDataWBC(1 To 64, 1 To 3) As Variant
          Dim gDataPLa(1 To 64, 1 To 1) As Variant
          Dim PltVal As Single


10        On Error GoTo LoadResults_Error

20        sql = "SELECT * from HaemResults WHERE " & _
                "SampleID = '" & mSampleID & "'"
30        Set tb = New Recordset
40        RecOpenClient 0, tb, sql
50        If tb.EOF Then
60            Exit Sub
70        End If

80        If Not IsNull(tb!Image) Or Trim(tb!Image) & "" <> "" Then
90            Me.Top = 200
100           Me.Height = 7458
110           Me.Left = 5000
120           Image1.Visible = True
130           Img = tb!Image.GetChunk(10000000#) & ""
140           UnCompressToFile Img
150           Image1.Picture = LoadPicture("C:\UncompressedImage.bmp")
160       Else
170           Me.Top = 200
180           Me.Height = 9315
190           gWBC.Visible = True
200           gRbc.Visible = True
210           gPla.Visible = True
220           For n = 1 To 64
230               gDataRBC(n, 1) = Asc(Mid$(tb!gRbc & String$(64, 1), n, 1))
240               gDataWBC(n, 1) = Asc(Mid$(tb!gwb1 & String$(64, 1), n, 1))
250               gDataWBC(n, 2) = Asc(Mid$(tb!gwb2 & String$(64, 1), n, 1))
260               gDataWBC(n, 3) = Asc(Mid$(tb!gwic & String$(64, 1), n, 1))
270               gDataPLa(n, 1) = Asc(Mid$(tb!gplt & String$(64, 1), n, 1))
280           Next

290           gRbc.ChartData = gDataRBC
300           gWBC.ChartData = gDataWBC
310           gPla.ChartData = gDataPLa
320           PltVal = Val(tb!Plt & "")
330           If PltVal < 100 Then
340               gPla.Plot.Axis(VtChAxisIdY).ValueScale.Auto = False
350               gPla.Plot.Axis(VtChAxisIdY).ValueScale.Maximum = 250
360           Else
370               gPla.Plot.Axis(VtChAxisIdY).ValueScale.Auto = True
380           End If
390       End If



400       Exit Sub

LoadResults_Error:

          Dim strES As String
          Dim intEL As Integer

410       intEL = Erl
420       strES = Err.Description
430       LogError "frmHaemGraphs", "LoadResults", intEL, strES, sql


End Sub
Public Property Let SampleID(ByVal sNewValue As String)

10        On Error GoTo SampleID_Error

20        mSampleID = sNewValue

30        Exit Property

SampleID_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmHaemGraphs", "SampleID", intEL, strES


End Property

Private Sub Form_Activate()

10        On Error GoTo Form_Activate_Error

20        LoadResults

30        Exit Sub

Form_Activate_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmHaemGraphs", "Form_Activate", intEL, strES


End Sub

Private Sub Form_Click()

10        Unload Me

End Sub


Private Sub gPla_Click()

10        Unload Me

End Sub

Private Sub gRBC_Click()

10        Unload Me

End Sub

Private Sub gWBC_Click()

10        Unload Me

End Sub

