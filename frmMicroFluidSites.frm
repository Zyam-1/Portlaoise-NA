VERSION 5.00
Begin VB.Form frmMicroFluidSites 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6750
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   555
      Left            =   3060
      Picture         =   "frmMicroFluidSites.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Exit"
      Top             =   4470
      Width           =   855
   End
   Begin VB.CommandButton cmdRemove 
      Height          =   555
      Left            =   3150
      Picture         =   "frmMicroFluidSites.frx":1986
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Remove from Fluid List"
      Top             =   2280
      Width           =   675
   End
   Begin VB.CommandButton cmdAdd 
      Height          =   555
      Left            =   3150
      Picture         =   "frmMicroFluidSites.frx":3308
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Add to Fluid List"
      Top             =   1560
      Width           =   675
   End
   Begin VB.ListBox lstFluid 
      Height          =   4545
      Left            =   3990
      TabIndex        =   1
      Top             =   480
      Width           =   2625
   End
   Begin VB.ListBox lstKnown 
      Height          =   4545
      Left            =   300
      TabIndex        =   0
      Top             =   480
      Width           =   2625
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Known Sites"
      Height          =   285
      Left            =   300
      TabIndex        =   6
      Top             =   210
      Width           =   2625
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sites regarded as 'Fluid'"
      Height          =   285
      Left            =   3990
      TabIndex        =   5
      Top             =   180
      Width           =   2625
   End
End
Attribute VB_Name = "frmMicroFluidSites"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub FillLists()

          Dim tb As Recordset
          Dim sql As String

10        lstFluid.Clear
20        lstKnown.Clear

30        sql = "SELECT Text FROM Lists WHERE " & _
                "ListType = 'FF'"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        Do While Not tb.EOF
70            lstFluid.AddItem tb!Text & ""
80            tb.MoveNext
90        Loop

          'sql = "SELECT Text FROM Lists WHERE " & _
           '      "ListType = 'SI' " & _
           '      "AND Text NOT IN " & _
           '      "  (SELECT Text FROM Lists WHERE " & _
           '      "  ListType = 'FF')"
100       sql = "SELECT A.Text FROM Lists A " & _
                "LEFT OUTER JOIN Lists B " & _
                "ON A.Text = B.Text AND B.ListType = 'FF' " & _
                "WHERE A.ListType = 'SI' " & _
                "AND B.Text IS NULL"

110       Set tb = New Recordset
120       RecOpenServer 0, tb, sql
130       Do While Not tb.EOF
140           lstKnown.AddItem tb!Text & ""
150           tb.MoveNext
160       Loop

End Sub

Private Sub cmdadd_Click()

          Dim Y As Integer
          Dim tb As Recordset
          Dim sql As String
          Dim Code As String

10        For Y = 0 To lstKnown.ListCount - 1
20            If lstKnown.Selected(Y) Then
30                lstFluid.AddItem lstKnown.List(Y)
40                sql = "SELECT * FROM Lists WHERE " & _
                        "ListType = 'SI' " & _
                        "AND Text = '" & lstKnown.List(Y) & "'"
50                Set tb = New Recordset
60                RecOpenServer 0, tb, sql
70                If Not tb.EOF Then
80                    Code = tb!Code
90                    tb.AddNew
100                   tb!Code = Code
110                   tb!ListType = "FF"
120                   tb!Text = lstKnown.List(Y)
130                   tb!InUse = 1
140                   tb.Update
150               End If
160               Exit For
170           End If
180       Next

190       FillLists

End Sub

Private Sub cmdCancel_Click()

10        Unload Me

End Sub

Private Sub cmdRemove_Click()

          Dim Y As Integer
          Dim sql As String

10        For Y = 0 To lstFluid.ListCount - 1
20            If lstFluid.Selected(Y) Then
30                sql = "DELETE FROM Lists WHERE " & _
                        "ListType = 'FF' " & _
                        "AND Text = '" & lstFluid.List(Y) & "'"
40                Cnxn(0).Execute sql
50                Exit For
60            End If
70        Next

80        FillLists

End Sub

Private Sub Form_Load()

10        FillLists

End Sub

