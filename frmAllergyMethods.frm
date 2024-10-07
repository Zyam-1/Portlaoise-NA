VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmAllergyMethods 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Allergy Method Assignment"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   9315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   885
      Left            =   7530
      Picture         =   "frmAllergyMethods.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4680
      Width           =   1245
   End
   Begin VB.CommandButton cmdAssign 
      Caption         =   "Assign"
      Height          =   795
      Left            =   6540
      Picture         =   "frmAllergyMethods.frx":1982
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1110
      Width           =   825
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   885
      Left            =   7530
      Picture         =   "frmAllergyMethods.frx":3304
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6630
      Width           =   1245
   End
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   7305
      Left            =   120
      TabIndex        =   2
      Top             =   210
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   12885
      _Version        =   393216
      Cols            =   5
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      AllowUserResizing=   1
      FormatString    =   "<Code       |<Long Name                                             |<Short Name |<Method  |<Allergy "
   End
   Begin VB.CommandButton cmdAddMethod 
      Caption         =   "&Add New Method"
      Height          =   885
      Left            =   7530
      Picture         =   "frmAllergyMethods.frx":4C86
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3030
      Width           =   1245
   End
   Begin VB.ListBox lstMethods 
      Height          =   2790
      Left            =   7380
      TabIndex        =   0
      Top             =   180
      Width           =   1515
   End
End
Attribute VB_Name = "frmAllergyMethods"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub FillG()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo FillG_Error

20        g.Rows = 2
30        g.AddItem ""
40        g.RemoveItem 1

50        sql = "SELECT Distinct Code, LongName, ShortName, Method, COALESCE(IsAllergy, 0) IsAllergy " & _
                "FROM ImmTestDefinitions WHERE " & _
                "Analyser = 'ImmunoCAP'"
60        Set tb = New Recordset
70        RecOpenServer 0, tb, sql
80        Do While Not tb.EOF
90            g.AddItem tb!Code & vbTab & _
                        tb!LongName & vbTab & _
                        tb!ShortName & vbTab & _
                        tb!Method & vbTab & _
                        IIf(tb!IsAllergy = 0, "No", "Yes")
100           tb.MoveNext
110       Loop

120       If g.Rows > 2 Then
130           g.RemoveItem 1
140       End If

150       Exit Sub

FillG_Error:

          Dim strES As String
          Dim intEL As Integer

160       intEL = Erl
170       strES = Err.Description
180       LogError "frmAllergyMethods", "FillG", intEL, strES, sql

End Sub

Private Sub FillList()

          Dim tb As Recordset
          Dim sql As String

10        On Error GoTo FillList_Error

20        lstMethods.Clear

30        sql = "SELECT Text FROM Lists WHERE " & _
                "ListType = 'Conjugate' " & _
                "ORDER BY ListOrder"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        Do While Not tb.EOF
70            lstMethods.AddItem tb!Text & ""
80            tb.MoveNext
90        Loop

100       Exit Sub

FillList_Error:

          Dim strES As String
          Dim intEL As Integer

110       intEL = Erl
120       strES = Err.Description
130       LogError "frmAllergyMethods", "FillList", intEL, strES, sql

End Sub

Private Sub cmdAddMethod_Click()

10        With frmListsGeneric
20            .ListType = "Conjugate"
30            .ListTypeName = "Conjugate"
40            .ListTypeNames = "Conjugates"
50            .Show 1
60        End With

70        FillList

End Sub


Private Sub cmdAssign_Click()

          Dim Y As Integer

10        On Error GoTo cmdAssign_Click_Error

20        If lstMethods.SelCount = 1 Then

30            g.Col = 3
40            For Y = 1 To g.Rows - 1
50                g.Row = Y
60                If g.CellBackColor = vbYellow Then
70                    g.TextMatrix(Y, 3) = lstMethods.Text
80                    g.CellBackColor = 0
90                    cmdSave.Enabled = True
100                   Exit For
110               End If
120           Next

130       End If

140       Exit Sub

cmdAssign_Click_Error:

          Dim strES As String
          Dim intEL As Integer

150       intEL = Erl
160       strES = Err.Description
170       LogError "frmAllergyMethods", "cmdAssign_Click", intEL, strES

End Sub

Private Sub cmdCancel_Click()

10        Unload Me

End Sub

Private Sub cmdSave_Click()

          Dim Y As Integer
          Dim Method As String
          Dim IsAllergy As Integer
          Dim LongName As String
          Dim sql As String

10        On Error GoTo cmdSave_Click_Error

20        For Y = 1 To g.Rows - 1
30            LongName = g.TextMatrix(Y, 1)
40            Method = g.TextMatrix(Y, 3)
50            IsAllergy = IIf(g.TextMatrix(Y, 4) = "Yes", 1, 0)
60            sql = "UPDATE ImmTestDefinitions " & _
                    "SET Method = '" & Method & "', " & _
                    "IsAllergy = " & IsAllergy & " " & _
                    "WHERE LongName = '" & LongName & "'"
70            Cnxn(0).Execute sql
80        Next

90        cmdSave.Enabled = False

100       Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

110       intEL = Erl
120       strES = Err.Description
130       LogError "frmAllergyMethods", "cmdSave_Click", intEL, strES, sql

End Sub

Private Sub Form_Load()

10        FillList
20        FillG

End Sub

Private Sub Form_Unload(Cancel As Integer)

10        If cmdSave.Enabled Then
20            If iMsg("Cancel without Saving?", vbQuestion + vbYesNo) = vbNo Then
30                Cancel = True
40            End If
50        End If

End Sub


Private Sub g_Click()

          Static SortOrder As Boolean
          Dim ySave As Integer
          Dim Y As Integer

10        If g.MouseRow = 0 Then
20            If SortOrder Then
30                g.Sort = flexSortGenericAscending
40            Else
50                g.Sort = flexSortGenericDescending
60            End If
70            SortOrder = Not SortOrder
80            Exit Sub
90        End If

100       If g.Col = 4 Then
110           g.TextMatrix(g.Row, 4) = IIf(g.TextMatrix(g.Row, 4) = "No", "Yes", "No")
120           cmdSave.Enabled = True
130       Else
140           ySave = g.Row
150           g.Col = 3
160           For Y = 1 To g.Rows - 1
170               g.Row = Y
180               g.CellBackColor = 0
190           Next
200           g.Row = ySave
210           g.CellBackColor = vbYellow
220       End If

End Sub


