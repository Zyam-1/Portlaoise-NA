VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form frmHaemErrors 
   Caption         =   "NetAcquire"
   ClientHeight    =   5505
   ClientLeft      =   1965
   ClientTop       =   6030
   ClientWidth     =   9255
   Icon            =   "frmHaemErrors.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   9255
   Begin VB.OptionButton optList 
      Caption         =   "List View"
      Height          =   255
      Left            =   7980
      TabIndex        =   8
      Top             =   1380
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.OptionButton optGrid 
      Caption         =   "Grid View"
      Height          =   255
      Left            =   7980
      TabIndex        =   7
      Top             =   1020
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid gFlags 
      Height          =   5295
      Left            =   120
      TabIndex        =   6
      Top             =   60
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   9340
      _Version        =   393216
      Cols            =   3
      RowHeightMin    =   325
      FormatString    =   "Flag                                                        |Flag Type                                 |Analyser Flag Type       "
   End
   Begin VB.ListBox List2 
      Height          =   4935
      Left            =   3480
      TabIndex        =   2
      Top             =   420
      Width           =   3465
   End
   Begin VB.CommandButton bCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   870
      Left            =   7800
      Picture         =   "frmHaemErrors.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4455
      Width           =   1245
   End
   Begin VB.ListBox List1 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   420
      Width           =   3255
   End
   Begin VB.Label Analyser 
      Height          =   195
      Left            =   2025
      TabIndex        =   5
      Top             =   1755
      Width           =   465
   End
   Begin VB.Label lblSystem 
      Caption         =   "System"
      Height          =   285
      Left            =   150
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.Label lblMorph 
      Caption         =   "Morphology"
      Height          =   285
      Left            =   3540
      TabIndex        =   3
      Top             =   90
      Visible         =   0   'False
      Width           =   1515
   End
End
Attribute VB_Name = "frmHaemErrors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mErrorNumber As Long

Public Property Let ErrorNumber(ByVal ErrorNumber As String)

10        On Error GoTo ErrorNumber_Error

20        mErrorNumber = Val(ErrorNumber)

30        Exit Property

ErrorNumber_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmHaemErrors", "ErrorNumber", intEL, strES


End Property

Private Sub FillListCell()

    Dim n As Long
    Dim Trial As Long

    On Error GoTo FillListCell_Error

    For n = 0 To 25
        Trial = 2 ^ n
        If mErrorNumber And Trial Then
            List1.AddItem Choose(n + 1, "Moving Average", "DFLT Flag", "Blast Flag", _
                                 "Variant Lymph", "DFLT (N)", "DFLT (E)", _
                                 "DFLT (L)", "IG Flag", "Band Flag", _
                                 "DFLT (M)", "DFLT (B)", "IG/Bands", _
                                 "FWBC", "WBC Count", "NRBC", _
                                 "DLTA", "NWBC", "RBC Morph", _
                                 "RRBC", "Plt Recount", "LRI", _
                                 "URI", "NOC Flow", "WOC Flow", _
                                 "RBC Flow", "Sampling Error")
            gFlags.AddItem List1.List(List1.ListIndex) & vbTab & "System"
        End If
    Next

    Exit Sub

FillListCell_Error:

    Dim strES As String
    Dim intEL As Integer



    intEL = Erl
    strES = Err.Description
    LogError "frmHaemErrors", "FillListCell", intEL, strES


End Sub

Private Sub bcancel_Click()

10        Unload Me

End Sub


Private Sub Form_Activate()



10    On Error GoTo Form_Activate_Error

      'If Val(Analyser) = 1 Then
20    Select Case SysOptHaemAn1(0)
          Case "CELLDYN"
              'Me.Width = 3030
              'bCancel.Left = 870
30            FillListCell
40        Case "ADVIA"
50            lblMorph.Visible = True
60            lblSystem.Visible = True
              'Me.Width = 5625
              'bCancel.Left = 2250
70            FillListAdvia
80            FillListExtendedIPU
90            Morph_Flags
100       Case "ADVIA60"
110           lblMorph.Visible = True
120           lblSystem.Visible = True
              'Me.Width = 5625
              'bCancel.Left = 2250
130           FillListAdvia60
140           Morph_Flags
150   End Select
      'Else
      '    Select Case SysOptHaemAn2(0)
      '      Case "CELLDYN"
      '        Me.Width = 3030
      '        bCancel.Left = 870
      '        FillListCell
      '      Case "ADVIA"
      '        lblMorph.Visible = True
      '        lblSystem.Visible = True
      '        Me.Width = 5460
      '        bCancel.Left = 2250
      '        FillListAdvia
      '        Morph_Flags
      '      Case "ADVIA60"
      '        lblMorph.Visible = True
      '        lblSystem.Visible = True
      '        Me.Width = 5460
      '        bCancel.Left = 2250
      '        FillListAdvia60
      '        Morph_Flags
      '    End Select
      'End If

160   Exit Sub

Form_Activate_Error:

      Dim strES As String
      Dim intEL As Integer



170   intEL = Erl
180   strES = Err.Description
190   LogError "frmHaemErrors", "Form_Activate", intEL, strES

End Sub

Private Sub FillListExtendedIPU()

      Dim tb As New Recordset
      Dim sql As String
      Dim n As Long
      Dim s As String
      Dim Flags As String
      Dim FlagType As String

10    On Error GoTo FillListExtendedIPU_Error

20    sql = "SELECT * from HaemFlags WHERE " & _
            "sampleid = '" & frmEditAll.txtSampleID & "'AND COALESCE(FlagType, '') <> ''"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    If tb.EOF Then Exit Sub

60    While Not tb.EOF
70        FlagType = GetLabLinkMapping("HaematologyFlagType", "Portlaoise", tb!Flags & "")
80        List1.AddItem tb!Flags & ""
90        gFlags.AddItem tb!Flags & vbTab & FlagType & vbTab & tb!FlagType & ""
100       tb.MoveNext
110   Wend

120   Exit Sub

FillListExtendedIPU_Error:
          Dim strES As String
          Dim intEL As Integer

130       intEL = Erl
140       strES = Err.Description
150       LogError "frmHaemErrors", "FillListExtendedIPU", intEL, strES

    End Sub

Private Sub FillListAdvia()

      Dim tb As New Recordset
      Dim sql As String
      Dim n As Long
      Dim s As String
      Dim Flags As String

10    On Error GoTo FillListAdvia_Error

20    sql = "SELECT * from HaemFlags WHERE " & _
            "sampleid = '" & frmEditAll.txtSampleID & "'AND COALESCE(FlagType, '') = ''"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    If tb.EOF Then Exit Sub

60    s = ""
70    Flags = Trim(tb!Flags & "")

80    If Flags = "" Then Exit Sub
90    s = ""
100   For n = 1 To Len(Flags) Step 3
110       s = FlagCodeToText(Mid(Flags, n, 2))
120       List1.AddItem s
130       gFlags.AddItem s & vbTab & "System"
140   Next



150   Exit Sub

FillListAdvia_Error:

      Dim strES As String
      Dim intEL As Integer



160   intEL = Erl
170   strES = Err.Description
180   LogError "frmHaemErrors", "FillListAdvia", intEL, strES, sql


End Sub

Private Function FlagCodeToText(ByVal Code As String) _
        As String

          Dim s As String

10        On Error GoTo FlagCodeToText_Error

20        Code = Trim(UCase(Code))

30        Select Case Code
          Case "AS": s = "CAS"
40        Case "L1": s = "Abnormal Lypm Count Check Film"
50        Case "G1": s = "Abnormal Gran Count Check Film"
60        Case "M2": s = "Abnoraml Mono Count Check Film"
70        Case "G2": s = "Abnormal Gran Count Check Film"
80        Case "G3": s = "Abnormal Gran Count Check Film"
90        Case "BC": s = "Baso Count Suspect"
100       Case "BR": s = "Baso Irregular Flow Rate"
110       Case "VB": s = "Baso No Valley"
120       Case "NB": s = "Baso Noise"
130       Case "BS": s = "Baso Saturation"
140       Case "TB": s = "Baso Temperature out of Range"
150       Case "CC": s = "Comparison Error MCHC/CHCM"
160       Case "WC": s = "Comparison Error WBCB/WBCP"
170       Case "HR": s = "Hgb Irregular Flow Rate"
180       Case "PH": s = "Hgb Power Low"
190       Case "PL": s = "Laser Power Low"
200       Case "MO": s = "Myeloperoxidase Deficiency"
210       Case "RB": s = "Nucleated Red Blood Cells"
220       Case "XR": s = "Perox Irregular Flow Rate"
230       Case "VX": s = "Perox No Valley"
240       Case "NX": s = "Perox Noise"
250       Case "PX": s = "Perox Power Low"
260       Case "XS": s = "Perox Saturation"
270       Case "TX": s = "Perox Temperature Out of Range"
280       Case "NW": s = "Platelet Clumps"
290       Case "NT": s = "Platelet Noise"
300       Case "OT": s = "Platelet Origin Noise"
310       Case "RR": s = "RBC Irregular Flow Rate"
320       Case "CT": s = "Retic - Platelet Interference"
330       Case "CA": s = "Retic Absorbtion Dist Abnormal"
340       Case "RF": s = "Retic Absorbtion Flatness"
350       Case "FC": s = "Retic Fit Suspect"
360       Case "CR": s = "Retic Irregular Flow Rate"
370       Case "NO": s = "Retic Noise Origin"
380       Case "CL": s = "Retic RBC Count Low"
390       Case "CS": s = "Retic Saturation Cells"
400       Case "SE": s = "Retic Slope Error"
410       Case "WS": s = "WBC Substitution"
420       Case Else: s = Code
430       End Select

440       FlagCodeToText = s

450       Exit Function

FlagCodeToText_Error:

          Dim strES As String
          Dim intEL As Integer



460       intEL = Erl
470       strES = Err.Description
480       LogError "frmHaemErrors", "FlagCodeToText", intEL, strES


End Function

Private Sub Morph_Flags()
          Dim sql As String
          Dim tb As New Recordset

10        On Error GoTo Morph_Flags_Error

20        sql = "SELECT * from Haemresults WHERE " & _
                "sampleid = '" & frmEditAll.txtSampleID & "'"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql

50        If Not tb.EOF Then
60            If Trim(tb!LS & "") <> "" Then
70                List2.AddItem "Left Shift " & Trim(tb!LS)
80                List2.Visible = True
90                gFlags.AddItem "Left Shift " & Trim(tb!LS) & vbTab & "Morphology"
100           End If
110           If Trim(tb!va & "") <> "" Then
120               List2.AddItem "HC Var " & Trim(tb!va)
130               List2.Visible = True
140               gFlags.AddItem "HC Var " & Trim(tb!va) & vbTab & "Morphology"
150           End If
160           If Trim(tb!At & "") <> "" Then
170               List2.AddItem "Atyp Lym " & Trim(tb!At)
180               List2.Visible = True
190               gFlags.AddItem "Atyp Lym " & Trim(tb!At) & vbTab & "Morphology"
200           End If
210           If Trim(tb!bl & "") <> "" Then
220               List2.AddItem "Blasts " & Trim(tb!bl)
230               List2.Visible = True
240               gFlags.AddItem "Blasts " & Trim(tb!bl) & vbTab & "Morphology"
250           End If
              'SUPRESS HYPT % FLAG AS ITS SENT WITH EVERY RESULTS
              '  If Trim(tb!hyp & "") <> "" Then
              '    List2.AddItem "Hypo% " & Trim(tb!hyp)
              '    List2.Visible = True
              '  End If
260           If Trim(tb!ho & "") <> "" Then
270               List2.AddItem "Hypo " & Trim(tb!ho)
280               List2.Visible = True
290               gFlags.AddItem "Hypo " & Trim(tb!ho) & vbTab & "Morphology"
300           End If
310           If Trim(tb!he & "") <> "" Then
320               List2.AddItem "Hyper " & Trim(tb!he)
330               List2.Visible = True
340               gFlags.AddItem "Hyper " & Trim(tb!he) & vbTab & "Morphology"
350           End If


360           If Trim(tb!An & "") <> "" Then
370               List2.AddItem "Anisocytosis " & Trim(tb!An)
380               List2.Visible = True
390               gFlags.AddItem "Anisocytosis " & Trim(tb!An) & vbTab & "Morphology"
400           End If

410           If Trim(tb!mi & "") <> "" Then
420               List2.AddItem "Microcytosis " & Trim(tb!mi)
430               List2.Visible = True
440               gFlags.AddItem "Microcytosis " & Trim(tb!mi) & vbTab & "Morphology"
450           End If

460           If Trim(tb!ca & "") <> "" Then
470               List2.AddItem "Macrocytosis " & Trim(tb!ca)
480               List2.Visible = True
490               gFlags.AddItem "Macrocytosis " & Trim(tb!ca) & vbTab & "Morphology"
500           End If

510           If Trim(tb!mpo & "") <> "" Then
520               List2.AddItem "Myeloperoxidase " & Trim(tb!mpo)
530               List2.Visible = True
540               gFlags.AddItem "Myeloperoxidase " & Trim(tb!mpo) & vbTab & "Morphology"
550           End If

560           If Trim(tb!Ig & "") <> "" Then
570               List2.AddItem "Immature Granulocytes " & Trim(tb!Ig)
580               List2.Visible = True
590               gFlags.AddItem "Immature Granulocytes " & vbTab & "Morphology"
600           End If

610           If Trim(tb!lplt & "") <> "" Then
620               List2.AddItem "Large Platelets " & Trim(tb!lplt)
630               List2.Visible = True
640               gFlags.AddItem "Large Platelets " & Trim(tb!lplt) & vbTab & "Morphology"
650           End If

660           If Trim(tb!pclm & "") <> "" Then
670               List2.AddItem "Platelet Clumps " & Trim(tb!pclm)
680               List2.Visible = True
690               gFlags.AddItem "Platelet Clumps " & Trim(tb!pclm) & vbTab & "Morphology"
700           End If

710           If Trim(tb!rbcf & "") <> "" Then
720               List2.AddItem "RBC Fragments " & Trim(tb!rbcf)
730               List2.Visible = True
740               gFlags.AddItem "RBC Fragments " & Trim(tb!rbcf) & vbTab & "Morphology"
750           End If

760           If Trim(tb!rbcg & "") <> "" Then
770               List2.AddItem "RBC Ghosts " & Trim(tb!rbcg)
780               List2.Visible = True
790               gFlags.AddItem "RBC Ghosts " & Trim(tb!rbcg) & vbTab & "Morphology"
800           End If
810       End If

820       Exit Sub

Morph_Flags_Error:

          Dim strES As String
          Dim intEL As Integer



830       intEL = Erl
840       strES = Err.Description
850       LogError "frmHaemErrors", "Morph_Flags", intEL, strES, sql


End Sub



Private Sub FillListAdvia60()

          Dim tb As New Recordset
          Dim sql As String
          Dim n As Long
          Dim s As String
          Dim Flags As String

10        On Error GoTo FillListAdvia60_Error

20        sql = "SELECT * from HaemFlags WHERE " & _
                "sampleid = '" & frmEditAll.txtSampleID & "'"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If tb.EOF Then Exit Sub
60        Flags = Trim(tb!Flags & "")
70        If Flags = "" Then Exit Sub

80        s = ""
90        For n = 1 To Len(Flags)
100           If Mid(Flags, n, 1) = vbCr Then
110               If Len(Trim(s)) = 2 Then
120                   s = FlagCodeToText(s)
130                   List1.AddItem s
140               Else
150                   List1.AddItem Trim(s)
160               End If
170               n = n + 1
180               s = ""
190           Else
200               s = s & (Mid(Flags, n, 1))
210           End If
220       Next


230       Exit Sub

FillListAdvia60_Error:

          Dim strES As String
          Dim intEL As Integer



240       intEL = Erl
250       strES = Err.Description
260       LogError "frmHaemErrors", "FillListAdvia60", intEL, strES, sql


End Sub



Private Sub Form_Load()

10    On Error GoTo Form_Load_Error

20    gFlags.Rows = 2
30    gFlags.FixedRows = 1
40    gFlags.FixedCols = 0
50    gFlags.Rows = 1

60    If UCase(GetOptionSetting("HaemFlagsDefaultView", "LIST")) = "LIST" Then
70        gFlags.Visible = False
80        optList.Value = True
90    Else
100       gFlags.Visible = True
110       optGrid.Value = True
120   End If


130   Exit Sub

Form_Load_Error:
      Dim strES As String
      Dim intEL As Integer

140   intEL = Erl
150   strES = Err.Description
160   LogError "frmHaemErrors", "Form_Load", intEL, strES

End Sub

Private Sub optGrid_Click()
gFlags.Visible = True
End Sub

Private Sub optList_Click()
gFlags.Visible = False
End Sub
