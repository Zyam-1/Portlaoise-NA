VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "mschrt20.ocx"
Begin VB.Form frmGraph 
   Caption         =   "NetAcquire - Graph"
   ClientHeight    =   10020
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12525
   Icon            =   "frmGraph.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10020
   ScaleWidth      =   12525
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   795
      Left            =   5445
      Picture         =   "frmGraph.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9090
      Width           =   1515
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Exit"
      Height          =   795
      Left            =   7470
      Picture         =   "frmGraph.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9090
      Width           =   1515
   End
   Begin VB.CommandButton cmdToggle 
      Caption         =   "Toggle Graph"
      Height          =   810
      Left            =   3510
      Picture         =   "frmGraph.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9090
      Width           =   1545
   End
   Begin MSChart20Lib.MSChart graChart 
      Height          =   9045
      Left            =   30
      OleObjectBlob   =   "frmGraph.frx":0D60
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   12495
   End
   Begin VB.PictureBox Picture1 
      Height          =   3930
      Left            =   630
      ScaleHeight     =   3870
      ScaleWidth      =   4500
      TabIndex        =   4
      Top             =   1845
      Width           =   4560
   End
End
Attribute VB_Name = "frmGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

10        Unload Me

End Sub


Private Sub cmdPrint_Click()

10        On Error GoTo cmdPrint_Click_Error

20        graChart.EditCopy
30        Picture1.Picture = Clipboard.GetData()
40        Printer.Print " "
50        Printer.PaintPicture Picture1.Picture, 0, 0
60        Printer.EndDoc

70        Exit Sub

cmdPrint_Click_Error:

          Dim strES As String
          Dim intEL As Integer



80        intEL = Erl
90        strES = Err.Description
100       LogError "frmGraph", "cmdPrint_Click", intEL, strES


End Sub

Private Sub cmdToggle_Click()

10        On Error GoTo cmdToggle_Click_Error

20        If graChart.chartType = VtChChartType3dBar Then
30            graChart.chartType = VtChChartType2dBar
40        ElseIf graChart.chartType = VtChChartType3dLine Then
50            graChart.chartType = VtChChartType2dLine
60        ElseIf graChart.chartType = VtChChartType2dBar Then
70            graChart.chartType = VtChChartType3dLine
80        ElseIf graChart.chartType = VtChChartType2dLine Then
90            graChart.chartType = VtChChartType3dBar


100       End If

110       Exit Sub

cmdToggle_Click_Error:

          Dim strES As String
          Dim intEL As Integer



120       intEL = Erl
130       strES = Err.Description
140       LogError "frmGraph", "cmdToggle_Click", intEL, strES


End Sub

Public Sub DrawGraph(ByVal f As Form, ByVal g As MSFlexGrid)
          Dim n As Long
          Dim Row As Long
          Dim c As Integer


10        On Error GoTo DrawGraph_Error

20        c = 2

30        If f.Name = "frmTotals" Then
40            c = 5
50        ElseIf f.Name = "frmCoagSourceTotals" Then
60            c = 6
70        End If

80        With graChart
              ' Displays a 3d chart with 8 columns and 8 rows
              ' data.
90            .chartType = VtChChartType3dBar
100           For n = 0 To 2
110               If f.o(n).Value = True Then .Title = f.o(n).Caption
120           Next
130           .RowCount = g.Rows - c
140           For Row = 1 To g.Rows - c
150               .Row = Row
160               If g.Cols > 2 Then
170                   .data = g.TextMatrix(Row, 2)
180               Else
190                   .data = g.TextMatrix(Row, 1)
200               End If
210               .RowLabel = g.TextMatrix(Row, 0)
220           Next Row
230           graChart.ColumnLabel = ""
              ' Use the chart as the backdrop of the legend.
240           .ShowLegend = True
250       End With

260       graChart.Visible = True




270       Exit Sub

DrawGraph_Error:

          Dim strES As String
          Dim intEL As Integer



280       intEL = Erl
290       strES = Err.Description
300       LogError "frmGraph", "DrawGraph", intEL, strES


End Sub

Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        Set_Font Me

30        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmGraph", "Form_Load", intEL, strES

End Sub

Private Sub graChart_PointSELECTed(Series As Integer, DataPoint As Integer, MouseFlags As Integer, Cancel As Integer)

10        On Error GoTo graChart_PointSELECTed_Error

20        graChart.Row = DataPoint
30        graChart.ColumnLabel = graChart.data

40        Exit Sub

graChart_PointSELECTed_Error:

          Dim strES As String
          Dim intEL As Integer



50        intEL = Erl
60        strES = Err.Description
70        LogError "frmGraph", "graChart_PointSELECTed", intEL, strES


End Sub
