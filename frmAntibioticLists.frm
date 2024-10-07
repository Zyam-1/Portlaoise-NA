VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmAntibioticLists 
   Caption         =   "NetAcquire - Micro Information"
   ClientHeight    =   7695
   ClientLeft      =   345
   ClientTop       =   570
   ClientWidth     =   10650
   Icon            =   "frmAntibioticLists.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   10650
   Begin VB.ComboBox cmbPrioritySec 
      Height          =   315
      Left            =   9240
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Top             =   660
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ComboBox cmbPriorityPri 
      Height          =   315
      Left            =   9240
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   825
      Left            =   9570
      Picture         =   "frmAntibioticLists.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4710
      Width           =   825
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   825
      Left            =   9570
      Picture         =   "frmAntibioticLists.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3810
      Width           =   825
   End
   Begin VB.CommandButton cmdOrganisms 
      Caption         =   "Organisms"
      Height          =   735
      Left            =   5760
      Picture         =   "frmAntibioticLists.frx":091E
      TabIndex        =   14
      Top             =   180
      Width           =   1725
   End
   Begin VB.CommandButton cmdNewSite 
      Caption         =   "New Site"
      Height          =   735
      Left            =   3960
      Picture         =   "frmAntibioticLists.frx":0C28
      TabIndex        =   13
      ToolTipText     =   "New Site"
      Top             =   945
      Width           =   1725
   End
   Begin VB.CommandButton cmdNewOrganismGroup 
      Caption         =   "New Organism Group"
      Height          =   735
      Left            =   3960
      Picture         =   "frmAntibioticLists.frx":0F32
      TabIndex        =   12
      ToolTipText     =   "New Organism"
      Top             =   180
      Width           =   1725
   End
   Begin VB.CommandButton cmdNewAntibiotic 
      Caption         =   "New Antibiotic"
      Height          =   735
      Left            =   5760
      Picture         =   "frmAntibioticLists.frx":123C
      TabIndex        =   11
      Top             =   945
      Width           =   1725
   End
   Begin VB.CommandButton cmdRemoveFromSecondary 
      Height          =   525
      Left            =   3180
      Picture         =   "frmAntibioticLists.frx":1546
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6360
      Width           =   765
   End
   Begin VB.CommandButton cmdTransferToSecondary 
      Height          =   525
      Left            =   3180
      Picture         =   "frmAntibioticLists.frx":1988
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5760
      Width           =   765
   End
   Begin VB.CommandButton cmdRemoveFromPrimary 
      Height          =   525
      Left            =   3180
      Picture         =   "frmAntibioticLists.frx":1DCA
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2850
      Width           =   765
   End
   Begin VB.Frame Frame4 
      Caption         =   "Secondary List"
      Height          =   2655
      Left            =   3990
      TabIndex        =   7
      Top             =   4860
      Width           =   5385
      Begin VB.CommandButton cmdMoveUpSec 
         Caption         =   "Move Up"
         Height          =   825
         Left            =   4500
         Picture         =   "frmAntibioticLists.frx":220C
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   660
         Width           =   795
      End
      Begin VB.CommandButton cmdMoveDownSec 
         Caption         =   "Move Down"
         Height          =   825
         Left            =   4530
         Picture         =   "frmAntibioticLists.frx":264E
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1500
         Width           =   795
      End
      Begin MSFlexGridLib.MSFlexGrid gSecondary 
         Height          =   2265
         Left            =   120
         TabIndex        =   24
         Top             =   300
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   3995
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         RowHeightMin    =   315
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
         ScrollBars      =   2
         FormatString    =   "|<Antibiotic                                    |^ AR  |^ RC  |^  ARP   "
      End
   End
   Begin VB.CommandButton cmdTransferToPrimary 
      Height          =   525
      Left            =   3180
      Picture         =   "frmAntibioticLists.frx":2A90
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2250
      Width           =   765
   End
   Begin VB.Frame Frame3 
      Caption         =   "Available Antibiotics"
      Height          =   5715
      Left            =   210
      TabIndex        =   4
      Top             =   1770
      Width           =   2715
      Begin VB.ListBox lstAvailable 
         Height          =   5130
         Left            =   150
         TabIndex        =   5
         Top             =   330
         Width           =   2385
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Primary List"
      Height          =   2595
      Left            =   3990
      TabIndex        =   3
      Top             =   1770
      Width           =   5385
      Begin VB.CommandButton cmdMoveUpPri 
         Caption         =   "Move Up"
         Height          =   825
         Left            =   4500
         Picture         =   "frmAntibioticLists.frx":2ED2
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   600
         Width           =   795
      End
      Begin VB.CommandButton cmdMoveDownPri 
         Caption         =   "Move Down"
         Height          =   825
         Left            =   4500
         Picture         =   "frmAntibioticLists.frx":3314
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   1440
         Width           =   795
      End
      Begin MSFlexGridLib.MSFlexGrid gPrimary 
         Height          =   2265
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   3995
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         RowHeightMin    =   315
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
         ScrollBars      =   2
         FormatString    =   "|<Antibiotic                                    |^ AR  |^ RC  |^  ARP   "
      End
   End
   Begin VB.Frame fraOrg 
      Height          =   1425
      Left            =   210
      TabIndex        =   0
      Top             =   150
      Width           =   2715
      Begin VB.ComboBox cmbSite 
         Height          =   315
         Left            =   90
         TabIndex        =   16
         Top             =   990
         Width           =   2415
      End
      Begin VB.ComboBox cmbOrganismGroup 
         Height          =   315
         Left            =   90
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   390
         Width           =   2415
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Site"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   810
         Width           =   270
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Organism Group"
         Height          =   195
         Left            =   90
         TabIndex        =   1
         Top             =   180
         Width           =   1140
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "ARP = Auto Report Priority"
      Height          =   195
      Left            =   7800
      TabIndex        =   28
      Top             =   1470
      Width           =   1875
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "RC = Report Criteria"
      Height          =   195
      Left            =   7800
      TabIndex        =   27
      Top             =   1275
      Width           =   1410
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "AR = Auto Report"
      Height          =   195
      Left            =   7800
      TabIndex        =   26
      Top             =   1080
      Width           =   1260
   End
   Begin VB.Image imgSquareTick 
      Height          =   225
      Left            =   9900
      Picture         =   "frmAntibioticLists.frx":3756
      Top             =   2280
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgSquareCross 
      Height          =   225
      Left            =   9900
      Picture         =   "frmAntibioticLists.frx":3A2C
      Top             =   2520
      Visible         =   0   'False
      Width           =   210
   End
End
Attribute VB_Name = "frmAntibioticLists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private OldRowIndex As Integer
Private NewRowIndex As Integer
Private bDontEnterCell As Boolean

Private Sub ClearLists()

10        On Error GoTo ClearLists_Error

20        gPrimary.Clear
30        gPrimary.Rows = 1
40        gPrimary.FormatString = "|<Antibiotic                                    |^ AR  |^ RC  |^  ARP   "
50        gPrimary.ColWidth(0) = 0

60        gSecondary.Clear
70        gSecondary.Rows = 1
80        gSecondary.FormatString = "|<Antibiotic                                    |^ AR  |^ RC  |^  ARP   "
90        gSecondary.ColWidth(0) = 0

100       Exit Sub

ClearLists_Error:

          Dim strES As String
          Dim intEL As Integer

110       intEL = Erl
120       strES = Err.Description
130       LogError "frmAntibioticLists", "ClearLists", intEL, strES


End Sub

Private Sub cmbOrganismGroup_Click()

10        On Error GoTo cmbOrganismGroup_Click_Error

20        ClearLists

30        FillAvailable

40        cmbSite = "Generic"

50        FillLists

60        Exit Sub

cmbOrganismGroup_Click_Error:

          Dim strES As String
          Dim intEL As Integer

70        intEL = Erl
80        strES = Err.Description
90        LogError "frmAntibioticLists", "cmbOrganismGroup_Click", intEL, strES


End Sub



Private Sub cmbPriorityPri_Click()

          Dim i As Integer
          Dim OldValue As Integer
          Dim NewValue As Integer
          Dim TempRow As Integer
          Dim EntryFound As Boolean

10        On Error GoTo cmbPriorityPri_Click_Error

20        EntryFound = False
30        With gPrimary
40            If .TextMatrix(.Row, 4) = "" Then
50                TempRow = .Row
60                For i = 1 To .Rows - 1
70                    If Val(.TextMatrix(i, 4)) = cmbPriorityPri Then
80                        EntryFound = True
90                    End If
100               Next i
110               If EntryFound Then
120                   .TextMatrix(TempRow, 4) = cmbPriorityPri
130                   For i = 1 To .Rows - 1
140                       If Val(.TextMatrix(i, 4)) >= cmbPriorityPri Then
150                           .TextMatrix(i, 4) = Val(.TextMatrix(i, 4)) + 1
160                       End If
170                   Next i
180               End If
190               .TextMatrix(TempRow, 4) = cmbPriorityPri
200           Else
210               OldValue = .TextMatrix(.Row, 4)
220               NewValue = cmbPriorityPri
230               If NewValue > OldValue Then
240                   For i = 1 To .Rows - 1
250                       If Val(.TextMatrix(i, 4)) > OldValue And Val(.TextMatrix(i, 4)) < NewValue Then
260                           .TextMatrix(i, 4) = Val(.TextMatrix(i, 4)) - 1
270                       ElseIf Val(.TextMatrix(i, 4)) = NewValue Then
280                           .TextMatrix(i, 4) = NewValue - 1
290                       ElseIf Val(.TextMatrix(i, 4)) = OldValue Then
300                           .TextMatrix(i, 4) = NewValue
310                       End If
320                   Next i
330               ElseIf NewValue < OldValue Then
340                   For i = 1 To .Rows - 1
350                       If Val(.TextMatrix(i, 4)) > NewValue And Val(.TextMatrix(i, 4)) < OldValue Then
360                           .TextMatrix(i, 4) = Val(.TextMatrix(i, 4)) + 1
370                       ElseIf Val(.TextMatrix(i, 4)) = NewValue Then
380                           .TextMatrix(i, 4) = NewValue + 1
390                       ElseIf Val(.TextMatrix(i, 4)) = OldValue Then
400                           .TextMatrix(i, 4) = NewValue
410                       End If
420                   Next i
430               End If
440           End If
450       End With
460       cmbPriorityPri.Visible = False
470       Exit Sub

cmbPriorityPri_Click_Error:

          Dim strES As String
          Dim intEL As Integer

480       intEL = Erl
490       strES = Err.Description
500       LogError "frmAntibioticLists", "cmbPriorityPri_Click", intEL, strES

End Sub

Private Sub cmbPriorityPri_LostFocus()
10        cmbPriorityPri.Visible = False
End Sub



Private Sub cmbPrioritySec_Click()
          Dim i As Integer
          Dim OldValue As Integer
          Dim NewValue As Integer
          Dim TempRow As Integer
          Dim EntryFound As Boolean


10        On Error GoTo cmbPrioritySec_Click_Error

20        EntryFound = False
30        With gSecondary
40            If .TextMatrix(.Row, 4) = "" Then
50                TempRow = .Row
60                For i = 1 To .Rows - 1
70                    If Val(.TextMatrix(i, 4)) = cmbPrioritySec Then
80                        EntryFound = True
90                    End If
100               Next i
110               If EntryFound Then
120                   .TextMatrix(TempRow, 4) = cmbPrioritySec
130                   For i = 1 To .Rows - 1
140                       If Val(.TextMatrix(i, 4)) >= cmbPrioritySec Then
150                           .TextMatrix(i, 4) = Val(.TextMatrix(i, 4)) + 1
160                       End If
170                   Next i
180               End If
190               .TextMatrix(TempRow, 4) = cmbPrioritySec
200           Else
210               OldValue = Val(.TextMatrix(.Row, 4))
220               NewValue = cmbPrioritySec
230               If NewValue > OldValue Then
240                   For i = 1 To .Rows - 1
250                       If Val(.TextMatrix(i, 4)) > OldValue And Val(.TextMatrix(i, 4)) < NewValue Then
260                           .TextMatrix(i, 4) = Val(.TextMatrix(i, 4)) - 1
270                       ElseIf Val(.TextMatrix(i, 4)) = NewValue Then
280                           .TextMatrix(i, 4) = NewValue - 1
290                       ElseIf Val(.TextMatrix(i, 4)) = OldValue Then
300                           .TextMatrix(i, 4) = NewValue
310                       End If
320                   Next i
330               ElseIf NewValue < OldValue Then
340                   For i = 1 To .Rows - 1
350                       If Val(.TextMatrix(i, 4)) > NewValue And Val(.TextMatrix(i, 4)) < OldValue Then
360                           .TextMatrix(i, 4) = Val(.TextMatrix(i, 4)) + 1
370                       ElseIf Val(.TextMatrix(i, 4)) = NewValue Then
380                           .TextMatrix(i, 4) = NewValue + 1
390                       ElseIf Val(.TextMatrix(i, 4)) = OldValue Then
400                           .TextMatrix(i, 4) = NewValue
410                       End If
420                   Next i
430               End If
440           End If
450       End With
460       cmbPrioritySec.Visible = False

470       Exit Sub

cmbPrioritySec_Click_Error:

          Dim strES As String
          Dim intEL As Integer

480       intEL = Erl
490       strES = Err.Description
500       LogError "frmAntibioticLists", "cmbPrioritySec_Click", intEL, strES

End Sub

Private Sub cmbPrioritySec_LostFocus()
10        cmbPrioritySec.Visible = False
End Sub

Private Sub cmbSite_Click()

10        On Error GoTo cmbSite_Click_Error

20        ClearLists

30        If cmbOrganismGroup = "" Then Exit Sub

40        FillAvailable
50        FillLists

60        Exit Sub

cmbSite_Click_Error:

          Dim strES As String
          Dim intEL As Integer

70        intEL = Erl
80        strES = Err.Description
90        LogError "frmAntibioticLists", "cmbSite_Click", intEL, strES


End Sub

Private Sub cmdCancel_Click()

10        Unload Me

End Sub

Private Sub cmdMoveDownPri_Click()

10        On Error GoTo cmdMoveDownPri_Click_Error

20        If gPrimary.Row = gPrimary.Rows - 1 Then Exit Sub

30        With gPrimary
40            .RowPosition(.Row) = .Row + 1
50            .Row = .Row + 1
60        End With

70        EnableSave True

80        Exit Sub

cmdMoveDownPri_Click_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmAntibioticLists", "cmdMoveDownPri_Click", intEL, strES

End Sub

Private Sub cmdMoveDownSec_Click()

10        On Error GoTo cmdMoveDownSec_Click_Error

20        If gSecondary.Row = gSecondary.Rows - 1 Then Exit Sub

30        With gSecondary
40            .RowPosition(.Row) = .Row + 1
50            .Row = .Row + 1
60        End With

70        EnableSave True

80        Exit Sub

cmdMoveDownSec_Click_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmAntibioticLists", "cmdMoveDownSec_Click", intEL, strES


End Sub

Private Sub cmdMoveUpPri_Click()

10        On Error GoTo cmdMoveUpPri_Click_Error

20        If gPrimary.Row = 1 Then Exit Sub

30        With gPrimary
40            .RowPosition(.Row) = .Row - 1
50            .Row = .Row - 1
60        End With

70        EnableSave True

80        Exit Sub

cmdMoveUpPri_Click_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmAntibioticLists", "cmdMoveUpPri_Click", intEL, strES

End Sub

Private Sub cmdMoveUpSec_Click()

10        On Error GoTo cmdMoveUpSec_Click_Error

20        If gSecondary.Row = 1 Then Exit Sub

30        With gSecondary
40            .RowPosition(.Row) = .Row - 1
50            .Row = .Row - 1
60        End With

70        EnableSave True

80        Exit Sub

cmdMoveUpSec_Click_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmAntibioticLists", "cmdMoveUpSec_Click", intEL, strES


End Sub

Private Sub cmdNewAntibiotic_Click()

10        On Error GoTo cmdNewAntibiotic_Click_Error

20        frmNewAntibiotics.Show 1
30        FillAvailable

40        Exit Sub

cmdNewAntibiotic_Click_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmAntibioticLists", "cmdNewAntibiotic_Click", intEL, strES


End Sub

Private Sub cmdNewOrganismGroup_Click()

10        With frmListsGeneric
20            .ListType = "OR"
30            .ListTypeName = "Organism Group"
40            .ListTypeNames = "Organism Groups"
50            .Show 1
60        End With

70        FillOrganismGroups

End Sub

Private Sub cmdNewSite_Click()

10        On Error GoTo cmdNewSite_Click_Error

20        frmMicroSites.Show 1
30        FillOrganismGroups

40        Exit Sub

cmdNewSite_Click_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "frmAntibioticLists", "cmdNewSite_Click", intEL, strES


End Sub

Private Sub cmdOrganisms_Click()

10        On Error GoTo cmdOrganisms_Click_Error

20        frmOrganisms.Show 1

30        Exit Sub

cmdOrganisms_Click_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmAntibioticLists", "cmdOrganisms_Click", intEL, strES


End Sub

Private Sub cmdRemoveFromPrimary_Click()

10        On Error GoTo cmdRemoveFromPrimary_Click_Error

20        lstAvailable.AddItem gPrimary.TextMatrix(gPrimary.Row, 1)
30        If gPrimary.Rows = 2 Then
40            gPrimary.Rows = 1
50        Else
60            gPrimary.RemoveItem gPrimary.Row
70        End If
80        EnableSave True

90        Exit Sub

cmdRemoveFromPrimary_Click_Error:

          Dim strES As String
          Dim intEL As Integer

100       intEL = Erl
110       strES = Err.Description
120       LogError "frmAntibioticLists", "cmdRemoveFromPrimary_Click", intEL, strES

End Sub

Private Sub cmdRemoveFromSecondary_Click()


10        On Error GoTo cmdRemoveFromSecondary_Click_Error

20        lstAvailable.AddItem gSecondary.TextMatrix(gSecondary.Row, 1)
30        If gSecondary.Rows = 2 Then
40            gSecondary.Rows = 1
50        Else
60            gSecondary.RemoveItem gSecondary.Row
70        End If
80        EnableSave True

90        Exit Sub

cmdRemoveFromSecondary_Click_Error:

          Dim strES As String
          Dim intEL As Integer

100       intEL = Erl
110       strES = Err.Description
120       LogError "frmAntibioticLists", "cmdRemoveFromSecondary_Click", intEL, strES


End Sub

Private Sub cmdSave_Click()

          Dim sql As String
          Dim Num As Long



10        On Error GoTo cmdSave_Click_Error



20        If cmbOrganismGroup = "" Then
30            iMsg "SELECT Organism", vbCritical
40            Exit Sub
50        End If

60        With gPrimary
70            For Num = 1 To .Rows - 1
80                .Row = Num
90                .Col = 2
100               If .CellPicture = imgSquareTick And (.TextMatrix(Num, 3) = "" Or .TextMatrix(Num, 4) = "") Then
110                   iMsg "Please select report criteria and report priorty for all auto reportable antibiotics in primary list"
120                   Exit Sub
130               End If
140           Next Num
150       End With

160       With gSecondary
170           For Num = 1 To .Rows - 1
180               .Row = Num
190               .Col = 2
200               If .CellPicture = imgSquareTick And (.TextMatrix(Num, 3) = "" Or .TextMatrix(Num, 4) = "") Then
210                   iMsg "Please select report criteria and report priorty for all auto reportable antibiotics in secondary list"
220                   Exit Sub
230               End If
240           Next Num
250       End With


260       Cnxn(0).BeginTrans

270       sql = "DELETE from ABDefinitions WHERE " & _
                "OrganismGroup = '" & cmbOrganismGroup & "' " & _
                "and Site = '" & cmbSite.Text & "'"
280       Cnxn(0).Execute sql

290       If cmbSite.Text = "Generic" Then
300           sql = "DELETE from ABDefinitions WHERE " & _
                    "OrganismGroup = '" & cmbOrganismGroup & "' " & _
                    "and Site is null"
310           Cnxn(0).Execute sql
320       End If

330       For Num = 1 To gPrimary.Rows - 1
340           sql = "INSERT into ABDefinitions " & _
                    "(AntibioticName, OrganismGroup, Site, Listorder, PriSec, AutoReport, AutoReportIf, AutoPriority) VALUES " & _
                    "('" & gPrimary.TextMatrix(Num, 1) & "', " & _
                    "'" & cmbOrganismGroup & "', " & _
                    "'" & cmbSite.Text & "', " & _
                    "'" & Num & "', " & _
                    "'P', "
350           gPrimary.Row = Num
360           gPrimary.Col = 2
370           If gPrimary.CellPicture = imgSquareTick Then
380               sql = sql & "1, '" & _
                        gPrimary.TextMatrix(gPrimary.Row, 3) & "', " & _
                        gPrimary.TextMatrix(gPrimary.Row, 4) & ")"
390           Else
400               sql = sql & "0, Null, Null)"

410           End If

420           Cnxn(0).Execute sql
430       Next


440       For Num = 1 To gSecondary.Rows - 1
450           sql = "INSERT into ABDefinitions " & _
                    "(AntibioticName, OrganismGroup, Site, Listorder, PriSec, AutoReport, AutoReportIf, AutoPriority) VALUES " & _
                    "('" & gSecondary.TextMatrix(Num, 1) & "', " & _
                    "'" & cmbOrganismGroup & "', " & _
                    "'" & cmbSite.Text & "', " & _
                    "'" & Num & "', " & _
                    "'S', "

460           gSecondary.Row = Num
470           gSecondary.Col = 2
480           If gSecondary.CellPicture = imgSquareTick Then
490               sql = sql & "1, '" & _
                        gSecondary.TextMatrix(gSecondary.Row, 3) & "', " & _
                        gSecondary.TextMatrix(gSecondary.Row, 4) & ")"
500           Else
510               sql = sql & "0, Null, Null)"

520           End If
530           Cnxn(0).Execute sql
540       Next

550       Cnxn(0).CommitTrans

560       EnableSave False




570       Exit Sub

cmdSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

580       intEL = Erl
590       strES = Err.Description
600       LogError "frmAntibioticLists", "cmdsave_Click", intEL, strES


End Sub

Private Sub cmdTransferToPrimary_Click()

          Dim Num As Long
          Dim s As String

10        On Error GoTo cmdTransferToPrimary_Click_Error


20        For Num = 0 To lstAvailable.ListCount - 1

30            If lstAvailable.Selected(Num) Then
40                s = vbTab & _
                      lstAvailable.List(Num)

50                gPrimary.AddItem s, gPrimary.Rows
60                gPrimary.Row = gPrimary.Rows - 1
70                gPrimary.Col = 2
80                gPrimary.CellPictureAlignment = flexAlignCenterCenter
90                Set gPrimary.CellPicture = imgSquareCross
100               lstAvailable.RemoveItem Num
110               Exit For
120           End If
130       Next

140       EnableSave True

150       Exit Sub

cmdTransferToPrimary_Click_Error:

          Dim strES As String
          Dim intEL As Integer

160       intEL = Erl
170       strES = Err.Description
180       LogError "frmAntibioticLists", "cmdTransferToPrimary_Click", intEL, strES


End Sub

Private Sub cmdTransferToSecondary_Click()

          Dim Num As Long
          Dim s As String

10        On Error GoTo cmdTransferToSecondary_Click_Error

20        For Num = 0 To lstAvailable.ListCount - 1
30            If lstAvailable.Selected(Num) Then
40                s = vbTab & _
                      lstAvailable.List(Num)

50                gSecondary.AddItem s, gSecondary.Rows
60                gSecondary.Row = gSecondary.Rows - 1
70                gSecondary.Col = 2
80                gSecondary.CellPictureAlignment = flexAlignCenterCenter
90                Set gSecondary.CellPicture = imgSquareCross
100               lstAvailable.RemoveItem Num
110               Exit For
120           End If
130       Next

140       EnableSave True

150       Exit Sub

cmdTransferToSecondary_Click_Error:

          Dim strES As String
          Dim intEL As Integer

160       intEL = Erl
170       strES = Err.Description
180       LogError "frmAntibioticLists", "cmdTransferToSecondary_Click", intEL, strES


End Sub

Private Sub EnableSave(ByVal blnEnable As Boolean)

10        On Error GoTo EnableSave_Error

20        cmdSave.Enabled = blnEnable
30        fraOrg.Enabled = Not blnEnable
40        cmdNewOrganismGroup.Enabled = Not blnEnable
50        cmdOrganisms.Enabled = Not blnEnable
60        cmdNewSite.Enabled = Not blnEnable
70        cmdNewAntibiotic.Enabled = Not blnEnable

80        Exit Sub

EnableSave_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmAntibioticLists", "EnableSave", intEL, strES


End Sub

Private Sub FillAvailable()

          Dim sql As String
          Dim tb As New Recordset

10        On Error GoTo FillAvailable_Error

20        lstAvailable.Clear

30        sql = "SELECT * from Antibiotics " & _
                "order by ListOrder asc"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql

60        Do While Not tb.EOF
70            lstAvailable.AddItem Trim$(tb!AntibioticName & "")
80            tb.MoveNext
90        Loop

100       Exit Sub

FillAvailable_Error:

          Dim strES As String
          Dim intEL As Integer

110       intEL = Erl
120       strES = Err.Description
130       LogError "frmAntibioticLists", "FillAvailable", intEL, strES, sql


End Sub

Private Sub FillLists()

          Dim tb As New Recordset
          Dim sql As String
          Dim Num As Long
          Dim strAB As String

          'Find Group and Site
10        On Error GoTo FillLists_Error

20        sql = "SELECT AntibioticName, OrganismGroup, Site, ListOrder, PriSec, " & _
                "COALESCE(AutoReport,0) AutoReport, AutoReportIf, COALESCE(AutoPriority,0) AutoPriority " & _
                "From ABDefinitions Where " & _
                "OrganismGroup = '" & cmbOrganismGroup & "' " & _
                "AND Site = '" & cmbSite.Text & "' " & _
                "order by ListOrder"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If tb.EOF Then    'Possibly only Generic site known
60            sql = "SELECT AntibioticName, OrganismGroup, Site, ListOrder, PriSec, " & _
                    "COALESCE(AutoReport,0) AutoReport, AutoReportIf, COALESCE(AutoPriority,0) AutoPriority " & _
                    "From ABDefinitions Where " & _
                    "OrganismGroup = '" & cmbOrganismGroup & "' " & _
                    "and ( Site = 'Generic' or Site is Null ) " & _
                    "order by ListOrder"
70            Set tb = New Recordset
80            RecOpenServer 0, tb, sql
90        End If

100       Do While Not tb.EOF

110           strAB = vbTab & _
                      Trim$(tb!AntibioticName & "") & vbTab & _
                      vbTab & _
                      tb!AutoReportIf & "" & vbTab & _
                      IIf(tb!AutoPriority = 0, "", tb!AutoPriority)


120           Select Case tb!PriSec
              Case "P": gPrimary.AddItem strAB, gPrimary.Rows
130               gPrimary.Row = gPrimary.Rows - 1
140               gPrimary.Col = 2
150               gPrimary.CellPictureAlignment = flexAlignCenterCenter
160               Set gPrimary.CellPicture = IIf(tb!AutoReport = 1, imgSquareTick.Picture, imgSquareCross)
170           Case "S": gSecondary.AddItem strAB, gSecondary.Rows
180               gSecondary.Row = gSecondary.Rows - 1
190               gSecondary.Col = 2
200               gSecondary.CellPictureAlignment = flexAlignCenterCenter
210               Set gSecondary.CellPicture = IIf(tb!AutoReport = 1, imgSquareTick.Picture, imgSquareCross)
220           End Select

230           For Num = 0 To lstAvailable.ListCount - 1
240               If lstAvailable.List(Num) = Trim$(tb!AntibioticName & "") Then
250                   lstAvailable.RemoveItem Num
260                   Exit For
270               End If
280           Next

290           tb.MoveNext

300       Loop

310       Exit Sub

FillLists_Error:

          Dim strES As String
          Dim intEL As Integer

320       intEL = Erl
330       strES = Err.Description
340       LogError "frmAntibioticLists", "FillLists", intEL, strES, sql


End Sub

Private Sub FillOrganismGroups()

          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo FillOrganismGroups_Error

20        cmbOrganismGroup.Clear

30        sql = "SELECT * from lists WHERE listtype = 'OR'"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        Do While Not tb.EOF
70            cmbOrganismGroup.AddItem initial2upper(Trim(tb!Text))
80            tb.MoveNext
90        Loop

100       Exit Sub

FillOrganismGroups_Error:

          Dim strES As String
          Dim intEL As Integer

110       intEL = Erl
120       strES = Err.Description
130       LogError "frmAntibioticLists", "FillOrganismGroups", intEL, strES, sql


End Sub

Private Sub FillSites()
          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo FillSites_Error

20        cmbSite.Clear

30        sql = "SELECT * from lists WHERE listtype = 'SI'"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        Do While Not tb.EOF
70            cmbSite.AddItem Trim(tb!Text)
80            tb.MoveNext
90        Loop

100       Exit Sub

FillSites_Error:

          Dim strES As String
          Dim intEL As Integer

110       intEL = Erl
120       strES = Err.Description
130       LogError "frmAntibioticLists", "FillSites", intEL, strES, sql


End Sub

Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        bDontEnterCell = False
30        ClearLists

40        FillAvailable

50        FillOrganismGroups

60        FillSites
70        cmbSite = "Generic"

80        Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmAntibioticLists", "Form_Load", intEL, strES


End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

10        On Error GoTo Form_QueryUnload_Error

20        If cmdSave.Enabled Then
30            If iMsg("Cancel without Saving?", vbQuestion + vbYesNo) = vbNo Then
40                Cancel = True
50            End If
60        End If

70        Exit Sub

Form_QueryUnload_Error:

          Dim strES As String
          Dim intEL As Integer

80        intEL = Erl
90        strES = Err.Description
100       LogError "frmAntibioticLists", "Form_QueryUnload", intEL, strES


End Sub

Private Sub gPrimary_Click()

          Dim s As String

10        On Error GoTo gPrimary_Click_Error

20        cmbPriorityPri.Visible = False

30        With gPrimary
40            If bDontEnterCell = False Then
50                NewRowIndex = gPrimary.Row
60                UnHighlightGridRow gPrimary, OldRowIndex
70                HighlightGridRow gPrimary, NewRowIndex
80            End If
90            Select Case .MouseCol
              Case 2:
100               .Col = 2
110               .CellPictureAlignment = flexAlignCenterCenter
120               If .CellPicture = imgSquareTick.Picture Then
130                   Set .CellPicture = imgSquareCross.Picture
140                   .TextMatrix(.Row, 3) = ""
150               Else
160                   Set .CellPicture = imgSquareTick.Picture
170                   .TextMatrix(.Row, 3) = "R"
180               End If
190               EnableSave True
200           Case 3:
210               .Col = 3
220               .CellAlignment = flexAlignCenterCenter
230               s = Trim$(.TextMatrix(.Row, 3))
240               Select Case s
                  Case "": s = "R"
250               Case "R": s = "S"
260               Case "S": s = "I"
270               Case "I": s = ""
280               Case Else: s = ""
290               End Select
300               .TextMatrix(.Row, 3) = s
310               EnableSave True
320           Case 4:
330               .Row = .Row
340               .Col = 2
350               If .CellPicture = imgSquareTick Then
360                   FillPriorityCombo gPrimary, cmbPriorityPri
370                   .Col = 4
380                   cmbPriorityPri.Width = .CellWidth
390                   cmbPriorityPri.Left = .CellLeft + .Left + Frame2.Left
400                   cmbPriorityPri.Top = .CellTop + .Top + Frame2.Top
410                   cmbPriorityPri.Visible = True
420                   cmbPriorityPri.SetFocus
430                   EnableSave True
440               End If
450           End Select
460       End With

470       Exit Sub

gPrimary_Click_Error:

          Dim strES As String
          Dim intEL As Integer

480       intEL = Erl
490       strES = Err.Description
500       LogError "frmAntibioticLists", "gPrimary_Click", intEL, strES

End Sub


Private Sub HighlightGridRow(g As MSFlexGrid, RowIndex As Integer)

10        On Error GoTo HighlightGridRow_Error

20        bDontEnterCell = True
30        g.Row = RowIndex
40        g.Col = 1
50        g.CellBackColor = vbYellow
          'g.CellForeColor = &H8000000E

60        bDontEnterCell = False

70        Exit Sub

HighlightGridRow_Error:

          Dim strES As String
          Dim intEL As Integer

80        intEL = Erl
90        strES = Err.Description
100       LogError "frmAntibioticLists", "HighlightGridRow", intEL, strES

End Sub
Private Sub UnHighlightGridRow(g As MSFlexGrid, RowIndex As Integer)

10        On Error GoTo UnHighlightGridRow_Error

20        bDontEnterCell = True
30        g.Row = RowIndex
40        g.Col = 1
50        g.CellBackColor = 0
          'g.CellForeColor = &H80000012
60        g.Col = 1
70        bDontEnterCell = False

80        Exit Sub

UnHighlightGridRow_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "frmAntibioticLists", "UnHighlightGridRow", intEL, strES

End Sub

Private Sub gPrimary_LeaveCell()
10        OldRowIndex = gPrimary.Row
End Sub

Private Sub FillPriorityCombo(g As MSFlexGrid, cmb As ComboBox)
          Dim i As Integer
10        With cmb
20            .Clear
30            For i = 1 To g.Rows - 1
40                .AddItem i
50            Next i
60        End With
End Sub

Private Sub gPrimary_Scroll()
10        cmbPriorityPri.Visible = False
End Sub

Private Sub gSecondary_Click()

          Dim s As String

10        On Error GoTo gSecondary_Click_Error

20        cmbPrioritySec.Visible = False

30        With gSecondary
40            If bDontEnterCell = False Then
50                NewRowIndex = gSecondary.Row
60                UnHighlightGridRow gSecondary, OldRowIndex
70                HighlightGridRow gSecondary, NewRowIndex
80            End If
90            Select Case .MouseCol
              Case 2:
100               .Col = 2
110               .CellPictureAlignment = flexAlignCenterCenter
120               If .CellPicture = imgSquareTick.Picture Then
130                   Set .CellPicture = imgSquareCross.Picture
140               Else
150                   Set .CellPicture = imgSquareTick.Picture
160               End If
170               EnableSave True
180           Case 3:
190               .Col = 3
200               .CellAlignment = flexAlignCenterCenter
210               s = Trim$(.TextMatrix(.Row, 3))
220               Select Case s
                  Case "": s = "R"
230               Case "R": s = "S"
240               Case "S": s = "I"
250               Case "I": s = ""
260               Case Else: s = ""
270               End Select
280               .TextMatrix(.Row, 3) = s
290               EnableSave True
300           Case 4:
310               .Row = .Row
320               .Col = 2
330               If .CellPicture = imgSquareTick Then
340                   FillPriorityCombo gSecondary, cmbPrioritySec
350                   .Col = 4
360                   cmbPrioritySec.Width = .CellWidth
370                   cmbPrioritySec.Left = .CellLeft + .Left + Frame4.Left
380                   cmbPrioritySec.Top = .CellTop + .Top + Frame4.Top
390                   cmbPrioritySec.Visible = True
400                   cmbPrioritySec.SetFocus
410                   EnableSave True
420               End If
430           End Select
440       End With

450       Exit Sub

gSecondary_Click_Error:

          Dim strES As String
          Dim intEL As Integer

460       intEL = Erl
470       strES = Err.Description
480       LogError "frmAntibioticLists", "gSecondary_Click", intEL, strES

End Sub

Private Sub gSecondary_LeaveCell()
10        OldRowIndex = gSecondary.Row
End Sub

Private Sub gSecondary_Scroll()
10        cmbPrioritySec.Visible = False
End Sub
