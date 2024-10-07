VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form frmHaemErrorsRep 
   Caption         =   "NetAcquire"
   ClientHeight    =   5505
   ClientLeft      =   1965
   ClientTop       =   6030
   ClientWidth     =   9255
   Icon            =   "frmHaemErrorsRep.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   9255
   Begin MSFlexGridLib.MSFlexGrid gFlags 
      Height          =   5295
      Left            =   60
      TabIndex        =   8
      Top             =   120
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   9340
      _Version        =   393216
      Cols            =   3
      RowHeightMin    =   325
      FormatString    =   "Flag                                                        |Flag Type                                 |Analyser Flag Type       "
   End
   Begin VB.OptionButton optGrid 
      Caption         =   "Grid View"
      Height          =   255
      Left            =   7860
      TabIndex        =   7
      Top             =   1200
      Width           =   1095
   End
   Begin VB.OptionButton optList 
      Caption         =   "List View"
      Height          =   255
      Left            =   7860
      TabIndex        =   6
      Top             =   1560
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.ListBox List2 
      Enabled         =   0   'False
      Height          =   4935
      Left            =   3480
      TabIndex        =   2
      Top             =   330
      Width           =   3765
   End
   Begin VB.CommandButton bCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   870
      Left            =   7800
      Picture         =   "frmHaemErrorsRep.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4440
      Width           =   1245
   End
   Begin VB.ListBox List1 
      Enabled         =   0   'False
      Height          =   4935
      Left            =   90
      TabIndex        =   0
      Top             =   330
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
      Left            =   2940
      TabIndex        =   3
      Top             =   90
      Visible         =   0   'False
      Width           =   1515
   End
End
Attribute VB_Name = "frmHaemErrorsRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mErrorNumber As Long
Private mSampleID As String
Private mDatetime As String

Public Property Let ErrorNumber(ByVal ErrorNumber As String)

10        On Error GoTo ErrorNumber_Error

20        mErrorNumber = Val(ErrorNumber)

30        Exit Property

ErrorNumber_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmHaemErrorsRep", "ErrorNumber", intEL, strES


End Property

Public Property Let SampleID(ByVal SampleID As String)

10        On Error GoTo SampleID_Error

20        mSampleID = SampleID

30        Exit Property

SampleID_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmHaemErrorsRep", "SampleID", intEL, strES


End Property

Public Property Let Datetime(ByVal Datetime As String)

10        On Error GoTo Datetime_Error

20        mDatetime = Datetime

30        Exit Property

Datetime_Error:

          Dim strES As String
          Dim intEL As Integer



40        intEL = Erl
50        strES = Err.Description
60        LogError "frmHaemErrorsRep", "Datetime", intEL, strES


End Property

Private Sub FillListCell()

          Dim n As Long
          Dim Trial As Long

10        On Error GoTo FillListCell_Error

20        For n = 0 To 25
30            Trial = 2 ^ n
40            If mErrorNumber And Trial Then
50                List1.AddItem Choose(n + 1, "Moving Average", "DFLT Flag", "Blast Flag", _
                                       "Variant Lymph", "DFLT (N)", "DFLT (E)", _
                                       "DFLT (L)", "IG Flag", "Band Flag", _
                                       "DFLT (M)", "DFLT (B)", "IG/Bands", _
                                       "FWBC", "WBC Count", "NRBC", _
                                       "DLTA", "NWBC", "RBC Morph", _
                                       "RRBC", "Plt Recount", "LRI", _
                                       "URI", "NOC Flow", "WOC Flow", _
                                       "RBC Flow", "Sampling Error")
60            End If
70            gFlags.AddItem List1.List(List1.ListIndex) & vbTab & "System"
80        Next

90        Exit Sub

FillListCell_Error:

          Dim strES As String
          Dim intEL As Integer



100       intEL = Erl
110       strES = Err.Description
120       LogError "frmHaemErrorsRep", "FillListCell", intEL, strES


End Sub

Private Sub bcancel_Click()

10        Unload Me

End Sub


Private Sub Form_Activate()



    On Error GoTo Form_Activate_Error

    If Val(Analyser) = 1 Then
        Select Case SysOptHaemAn1(0)
        Case "CELLDYN"
            'Me.Width = 3030
            'bCancel.Left = 870
            FillListCell
        Case "ADVIA"
            lblMorph.Visible = True
            lblSystem.Visible = True
            'Me.Width = 5460
            'bCancel.Left = 2250
            FillListAdvia
            FillListExtendedIPU
            Morph_Flags
        Case "ADVIA60"
            lblMorph.Visible = True
            lblSystem.Visible = True
            'Me.Width = 5460
            'bCancel.Left = 2250
            FillListAdvia60
            Morph_Flags
        End Select
    Else
        Select Case SysOptHaemAn2(0)
        Case "CELLDYN"
            'Me.Width = 3030
            'bCancel.Left = 870
            FillListCell
        Case "ADVIA"
            lblMorph.Visible = True
            lblSystem.Visible = True
            'Me.Width = 5460
            'bCancel.Left = 2250
            FillListAdvia
            Morph_Flags
        Case "ADVIA60"
            lblMorph.Visible = True
            lblSystem.Visible = True
            'Me.Width = 5460
            'bCancel.Left = 2250
            FillListAdvia60
            Morph_Flags
        End Select
    End If

    Exit Sub

Form_Activate_Error:

    Dim strES As String
    Dim intEL As Integer



    intEL = Erl
    strES = Err.Description
    LogError "frmHaemErrorsRep", "Form_Activate", intEL, strES

End Sub

Private Sub FillListAdvia()

    Dim tb As New Recordset
    Dim sql As String
    Dim n As Long
    Dim s As String
    Dim Flags As String

    On Error GoTo FillListAdvia_Error

    sql = "SELECT * from HaemFlagsrep WHERE " & _
          "sampleid = '" & mSampleID & "' and datetime between '" & Left(mDatetime, 17) & ":00' and '" & Left(mDatetime, 17) & ":59'"
    Set tb = New Recordset
    RecOpenServer 0, tb, sql
    If tb.EOF Then Exit Sub
    s = ""
    Flags = Trim(tb!Flags & "")

    If Flags = "" Then Exit Sub

    If tb!FlagType & "" = "" Then

        s = ""
        For n = 1 To Len(Flags) Step 3
            s = FlagCodeToText(Mid(Flags, n, 2))
            List1.AddItem s
            gFlags.AddItem s & vbTab & "System"
        Next
    Else
        List1.AddItem tb!Flags & ""
        gFlags.AddItem tb!Flags & vbTab & "System"
    End If
    

    Exit Sub

FillListAdvia_Error:

    Dim strES As String
    Dim intEL As Integer



    intEL = Erl
    strES = Err.Description
    LogError "frmHaemErrorsRep", "FillListAdvia", intEL, strES, sql


End Sub

Private Function FlagCodeToText(ByVal Code As String) _
        As String

          Dim s As String

10        On Error GoTo FlagCodeToText_Error

20        Code = Trim(UCase(Code))

30        Select Case Code
          Case "L1": s = "Abnormal Lypm Count Check Film"
40        Case "G1": s = "Abnormal Gran Count Check Film"
50        Case "M2": s = "Abnoraml Mono Count Check Film"
60        Case "G2": s = "Abnormal Gran Count Check Film"
70        Case "G3": s = "Abnormal Gran Count Check Film"
80        Case "BC": s = "Baso Count Suspect"
90        Case "BR": s = "Baso Irregular Flow Rate"
100       Case "VB": s = "Baso No Valley"
110       Case "NB": s = "Baso Noise"
120       Case "BS": s = "Baso Saturation"
130       Case "TB": s = "Baso Temperature out of Range"
140       Case "CC": s = "Comparison Error MCHC/CHCM"
150       Case "WC": s = "Comparison Error WBCB/WBCP"
160       Case "HR": s = "Hgb Irregular Flow Rate"
170       Case "PH": s = "Hgb Power Low"
180       Case "PL": s = "Laser Power Low"
190       Case "MO": s = "Myeloperoxidase Deficiency"
200       Case "RB": s = "Nucleated Red Blood Cells"
210       Case "XR": s = "Perox Irregular Flow Rate"
220       Case "VX": s = "Perox No Valley"
230       Case "NX": s = "Perox Noise"
240       Case "PX": s = "Perox Power Low"
250       Case "XS": s = "Perox Saturation"
260       Case "TX": s = "Perox Temperature Out of Range"
270       Case "NW": s = "Platelet Clumps"
280       Case "NT": s = "Platelet Noise"
290       Case "OT": s = "Platelet Origin Noise"
300       Case "RR": s = "RBC Irregular Flow Rate"
310       Case "CT": s = "Retic - Platelet Interference"
320       Case "CA": s = "Retic Absorbtion Dist Abnormal"
330       Case "RF": s = "Retic Absorbtion Flatness"
340       Case "FC": s = "Retic Fit Suspect"
350       Case "CR": s = "Retic Irregular Flow Rate"
360       Case "NO": s = "Retic Noise Origin"
370       Case "CL": s = "Retic RBC Count Low"
380       Case "CS": s = "Retic Saturation Cells"
390       Case "SE": s = "Retic Slope Error"
400       Case "WS": s = "WBC Substitution"
410       Case Else: s = Code
420       End Select

430       FlagCodeToText = s

440       Exit Function

FlagCodeToText_Error:

          Dim strES As String
          Dim intEL As Integer



450       intEL = Erl
460       strES = Err.Description
470       LogError "frmHaemErrorsRep", "FlagCodeToText", intEL, strES


End Function

Private Sub Morph_Flags()
      Dim sql As String
      Dim tb As New Recordset

10    On Error GoTo Morph_Flags_Error

20    sql = "SELECT * from Haemrepeats WHERE " & _
            "sampleid = '" & frmEditAll.txtSampleID & "' and rundatetime = '" & mDatetime & "'"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql

50    If Not tb.EOF Then
60        If Trim(tb!LS & "") <> "" Then
70            List2.AddItem "Left Shift " & Trim(tb!LS)
80            List2.Visible = True
90            gFlags.AddItem "Left Shift " & Trim(tb!LS) & vbTab & "Morphology"
100       End If
110       If Trim(tb!va & "") <> "" Then
120           List2.AddItem "HC Var " & Trim(tb!va)
130           List2.Visible = True
140           gFlags.AddItem "HC Var " & Trim(tb!va) & vbTab & "Morphology"
150       End If
160       If Trim(tb!At & "") <> "" Then
170           List2.AddItem "Atyp Lym " & Trim(tb!At)
180           List2.Visible = True
190           gFlags.AddItem "Atyp Lym " & Trim(tb!At) & vbTab & "Morphology"
200       End If
210       If Trim(tb!bl & "") <> "" Then
220           List2.AddItem "Blasts " & Trim(tb!bl)
230           List2.Visible = True
240           gFlags.AddItem "Blasts " & Trim(tb!bl) & vbTab & "Morphology"
250       End If
          'SUPRESS HYPT % FLAG AS ITS SENT WITH EVERY RESULTS
          '  If Trim(tb!hyp & "") <> "" Then
          '    List2.AddItem "Hypo% " & Trim(tb!hyp)
          '    List2.Visible = True
          '  End If
260       If Trim(tb!ho & "") <> "" Then
270           List2.AddItem "Hypo " & Trim(tb!ho)
280           List2.Visible = True
290           gFlags.AddItem "Hypo " & Trim(tb!ho) & vbTab & "Morphology"
300       End If
310       If Trim(tb!he & "") <> "" Then
320           List2.AddItem "Hyper " & Trim(tb!he)
330           List2.Visible = True
340           gFlags.AddItem "Hyper " & Trim(tb!he) & vbTab & "Morphology"
350       End If


360       If Trim(tb!An & "") <> "" Then
370           List2.AddItem "Anisocytosis " & Trim(tb!An)
380           List2.Visible = True
390           gFlags.AddItem "Anisocytosis " & Trim(tb!An) & vbTab & "Morphology"
400       End If

410       If Trim(tb!mi & "") <> "" Then
420           List2.AddItem "Microcytosis " & Trim(tb!mi)
430           List2.Visible = True
440           gFlags.AddItem "Microcytosis " & Trim(tb!mi) & vbTab & "Morphology"
450       End If

460       If Trim(tb!ca & "") <> "" Then
470           List2.AddItem "Macrocytosis " & Trim(tb!ca)
480           List2.Visible = True
490           gFlags.AddItem "Macrocytosis " & Trim(tb!ca) & vbTab & "Morphology"
500       End If

510       If Trim(tb!mpo & "") <> "" Then
520           List2.AddItem "Myeloperoxidase " & Trim(tb!mpo)
530           List2.Visible = True
540           gFlags.AddItem "Myeloperoxidase " & Trim(tb!mpo) & vbTab & "Morphology"
550       End If

560       If Trim(tb!Ig & "") <> "" Then
570           List2.AddItem "Immature Granulocytes " & Trim(tb!Ig)
580           List2.Visible = True
590           gFlags.AddItem "Immature Granulocytes " & vbTab & "Morphology"
600       End If

610       If Trim(tb!lplt & "") <> "" Then
620           List2.AddItem "Large Platelets " & Trim(tb!lplt)
630           List2.Visible = True
640           gFlags.AddItem "Large Platelets " & Trim(tb!lplt) & vbTab & "Morphology"
650       End If

660       If Trim(tb!pclm & "") <> "" Then
670           List2.AddItem "Platelet Clumps " & Trim(tb!pclm)
680           List2.Visible = True
690           gFlags.AddItem "Platelet Clumps " & Trim(tb!pclm) & vbTab & "Morphology"
700       End If

710       If Trim(tb!rbcf & "") <> "" Then
720           List2.AddItem "RBC Fragments " & Trim(tb!rbcf)
730           List2.Visible = True
740           gFlags.AddItem "RBC Fragments " & Trim(tb!rbcf) & vbTab & "Morphology"
750       End If

760       If Trim(tb!rbcg & "") <> "" Then
770           List2.AddItem "RBC Ghosts " & Trim(tb!rbcg)
780           List2.Visible = True
790           gFlags.AddItem "RBC Ghosts " & Trim(tb!rbcg) & vbTab & "Morphology"
800       End If
810   End If

      'If Not tb.EOF Then
      '    If Trim(tb!LS & "") <> "" Then
      '        List2.AddItem "Left Shift " & Trim(tb!LS)
      '        List2.Visible = True
      '    End If
      '    If Trim(tb!va & "") <> "" Then
      '        List2.AddItem "HC Var " & Trim(tb!va)
      '        List2.Visible = True
      '    End If
      '    If Trim(tb!At & "") <> "" Then
      '        List2.AddItem "Atyp Lym " & Trim(tb!At)
      '        List2.Visible = True
      '    End If
      '    If Trim(tb!bl & "") <> "" Then
      '        List2.AddItem "Blasts " & Trim(tb!bl)
      '        List2.Visible = True
      '    End If
      '    If Trim(tb!ho & "") <> "" Then
      '        List2.AddItem "Hypo " & Trim(tb!ho)
      '        List2.Visible = True
      '    End If
      '    If Trim(tb!he & "") <> "" Then
      '        List2.AddItem "Hyper " & Trim(tb!he)
      '        List2.Visible = True
      '    End If
      '
      '
      '    If Trim(tb!An & "") <> "" Then
      '        List2.AddItem "Anisocytosis " & Trim(tb!An)
      '        List2.Visible = True
      '    End If
      '
      '    If Trim(tb!mi & "") <> "" Then
      '        List2.AddItem "Microcytosis " & Trim(tb!mi)
      '        List2.Visible = True
      '    End If
      '
      '    If Trim(tb!ca & "") <> "" Then
      '        List2.AddItem "Macrocytosis " & Trim(tb!ca)
      '        List2.Visible = True
      '    End If
      '
      '    If Trim(tb!mpo & "") <> "" Then
      '        List2.AddItem "Myeloperoxidase " & Trim(tb!mpo)
      '        List2.Visible = True
      '    End If
      '
      '    If Trim(tb!Ig & "") <> "" Then
      '        List2.AddItem "Immature Granulocytes " & Trim(tb!Ig)
      '        List2.Visible = True
      '    End If
      '
      '    If Trim(tb!lplt & "") <> "" Then
      '        List2.AddItem "Large Platelets " & Trim(tb!lplt)
      '        List2.Visible = True
      '    End If
      '
      '    If Trim(tb!pclm & "") <> "" Then
      '        List2.AddItem "Platelet Clumps " & Trim(tb!pclm)
      '        List2.Visible = True
      '    End If
      '
      '    If Trim(tb!rbcf & "") <> "" Then
      '        List2.AddItem "RBC Fragments " & Trim(tb!pclm)
      '        List2.Visible = True
      '    End If
      '
      '    If Trim(tb!rbcg & "") <> "" Then
      '        List2.AddItem "RBC Ghosts " & Trim(tb!rbcg)
      '        List2.Visible = True
      '    End If
      'End If

820   Exit Sub

Morph_Flags_Error:

      Dim strES As String
      Dim intEL As Integer



830   intEL = Erl
840   strES = Err.Description
850   LogError "frmHaemErrorsRep", "Morph_Flags", intEL, strES, sql


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
260       LogError "frmHaemErrorsRep", "FillListAdvia60", intEL, strES


End Sub

Private Sub FillListExtendedIPU()

      Dim tb As New Recordset
      Dim sql As String
      Dim n As Long
      Dim s As String
      Dim Flags As String
      Dim FlagType As String

10    On Error GoTo FillListExtendedIPU_Error

20    sql = "SELECT * from HaemFlagsRep WHERE " & _
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
160   LogError "frmHaemErrorsRep", "Form_Load", intEL, strES

End Sub

Private Sub optGrid_Click()
gFlags.Visible = True
End Sub

Private Sub optList_Click()
gFlags.Visible = False
End Sub
