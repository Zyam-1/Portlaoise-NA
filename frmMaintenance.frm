VERSION 5.00
Begin VB.Form frmMaintenance 
   Caption         =   "NetAcquire"
   ClientHeight    =   6330
   ClientLeft      =   3690
   ClientTop       =   2400
   ClientWidth     =   6150
   Icon            =   "frmMaintenance.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6330
   ScaleWidth      =   6150
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   525
      Left            =   2490
      Picture         =   "frmMaintenance.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5340
      Width           =   1245
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Remove All References"
      Height          =   615
      Left            =   3060
      Picture         =   "frmMaintenance.frx":0974
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1020
      Width           =   1875
   End
   Begin VB.Frame Frame1 
      Caption         =   "Auto Select From"
      Height          =   1935
      Left            =   240
      TabIndex        =   2
      Top             =   150
      Width           =   1845
      Begin VB.OptionButton optAuto 
         Caption         =   "Demographics"
         Height          =   255
         Index           =   4
         Left            =   210
         TabIndex        =   7
         Top             =   1500
         Width           =   1365
      End
      Begin VB.OptionButton optAuto 
         Caption         =   "Externals"
         Height          =   255
         Index           =   3
         Left            =   210
         TabIndex        =   6
         Top             =   1200
         Width           =   1005
      End
      Begin VB.OptionButton optAuto 
         Caption         =   "Coagulation"
         Height          =   255
         Index           =   2
         Left            =   210
         TabIndex        =   5
         Top             =   900
         Width           =   1185
      End
      Begin VB.OptionButton optAuto 
         Caption         =   "Biochemistry"
         Height          =   255
         Index           =   1
         Left            =   210
         TabIndex        =   4
         Top             =   600
         Width           =   1305
      End
      Begin VB.OptionButton optAuto 
         Caption         =   "Haematology"
         Height          =   255
         Index           =   0
         Left            =   210
         TabIndex        =   3
         Top             =   300
         Width           =   1275
      End
   End
   Begin VB.TextBox txtSampleID 
      Height          =   285
      Left            =   3060
      TabIndex        =   0
      Top             =   660
      Width           =   1845
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmMaintenance.frx":0FDE
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   2505
      Left            =   240
      TabIndex        =   9
      Top             =   2280
      Width           =   5685
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "SampleID"
      Height          =   195
      Left            =   3600
      TabIndex        =   1
      Top             =   450
      Width           =   690
   End
End
Attribute VB_Name = "frmMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub AutoFill(ByVal Index As Integer)

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo AutoFill_Error

20        sql = "Select top 1 SampleID from "

30        Select Case Index
          Case 0: sql = sql & "HaemResults"
40        Case 1: sql = sql & "BioResults"
50        Case 2: sql = sql & "CoagResults"
60        Case 3: sql = sql & "ExtResults"
70        Case 4: sql = sql & "Demographics"
80        End Select

90        sql = sql & " order by SampleID Desc"
100       Set tb = New Recordset
110       RecOpenServer 0, tb, sql
120       If Not tb.EOF Then
130           txtSampleID = tb!SampleID
140       Else
150           txtSampleID = ""
160       End If

170       Exit Sub

AutoFill_Error:

          Dim strES As String
          Dim intEL As Integer



180       intEL = Erl
190       strES = Err.Description
200       LogError "frmMaintenance", "AutoFill", intEL, strES, sql


End Sub

Private Sub cmdCancel_Click()

10        Unload Me

End Sub

Private Sub cmdDelete_Click()

          Dim n As Integer
          Dim sql As String


10        On Error GoTo cmdDelete_Click_Error

20        If iMsg("Are you sure?", vbQuestion + vbYesNo, , vbRed, 18) = vbYes Then

30            sql = "Delete from EndResults where " & _
                    "SampleID = '" & txtSampleID & "'"
40            Cnxn(0).Execute sql

50            sql = "Delete from EndRequests where " & _
                    "SampleID = '" & txtSampleID & "'"
60            Cnxn(0).Execute sql

70            sql = "Delete from immResults where " & _
                    "SampleID = '" & txtSampleID & "'"
80            Cnxn(0).Execute sql

90            sql = "Delete from immRequests where " & _
                    "SampleID = '" & txtSampleID & "'"
100           Cnxn(0).Execute sql

110           sql = "Delete from HaemResults where " & _
                    "SampleID = '" & txtSampleID & "'"
120           Cnxn(0).Execute sql

130           sql = "Delete from BioResults where " & _
                    "SampleID = '" & txtSampleID & "'"
140           Cnxn(0).Execute sql

150           sql = "Delete from BioRequests where " & _
                    "SampleID = '" & txtSampleID & "'"
160           Cnxn(0).Execute sql

170           sql = "Delete from CoagResults where " & _
                    "SampleID = '" & txtSampleID & "'"
180           Cnxn(0).Execute sql

190           sql = "Delete from coagRequests where " & _
                    "SampleID = '" & txtSampleID & "'"
200           Cnxn(0).Execute sql

210           sql = "Delete from ExtResults where " & _
                    "SampleID = '" & txtSampleID & "'"
220           Cnxn(0).Execute sql

230           sql = "Delete from Demographics where " & _
                    "SampleID = '" & txtSampleID & "'"
240           Cnxn(0).Execute sql

250       End If

260       For n = 0 To 4
270           If optAuto(n) Then
280               AutoFill n
290               Exit For
300           End If
310       Next

320       Exit Sub

cmdDelete_Click_Error:

          Dim strES As String
          Dim intEL As Integer



330       intEL = Erl
340       strES = Err.Description
350       LogError "frmMaintenance", "cmdDelete_Click", intEL, strES, sql


End Sub


Private Sub optAuto_Click(Index As Integer)

10        On Error GoTo optAuto_Click_Error

20        AutoFill Index

30        Exit Sub

optAuto_Click_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmMaintenance", "optAuto_Click", intEL, strES


End Sub


