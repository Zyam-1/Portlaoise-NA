VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAudit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire"
   ClientHeight    =   8925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8925
   ScaleWidth      =   13590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0FFFF&
      Height          =   915
      Left            =   11640
      ScaleHeight     =   855
      ScaleWidth      =   1725
      TabIndex        =   6
      Top             =   6540
      Width           =   1785
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "Deletion"
         Height          =   195
         Left            =   570
         TabIndex        =   12
         Top             =   600
         Width           =   585
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "Changes Made"
         Height          =   195
         Left            =   570
         TabIndex        =   11
         Top             =   330
         Width           =   1080
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "Red"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   195
         TabIndex        =   10
         Top             =   600
         Width           =   300
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "Blue"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   180
         TabIndex        =   9
         Top             =   330
         Width           =   315
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "No Changes"
         Height          =   195
         Left            =   570
         TabIndex        =   8
         Top             =   60
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "Green"
         ForeColor       =   &H0000C000&
         Height          =   195
         Left            =   60
         TabIndex        =   7
         Top             =   60
         Width           =   435
      End
   End
   Begin MSComctlLib.ListView ListView 
      Height          =   5415
      Left            =   11640
      TabIndex        =   5
      Top             =   1050
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   9551
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.TextBox txtSampleID 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   11647
      TabIndex        =   3
      Top             =   600
      Width           =   1770
   End
   Begin RichTextLib.RichTextBox rtb 
      Height          =   8475
      Left            =   60
      TabIndex        =   2
      Top             =   270
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   14949
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmAudit.frx":0000
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   1100
      Left            =   11917
      Picture         =   "frmAudit.frx":008B
      Style           =   1  'Graphical
      TabIndex        =   0
      Tag             =   "101"
      ToolTipText     =   "Exit Screen"
      Top             =   7650
      Width           =   1200
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   165
      Left            =   90
      TabIndex        =   1
      Top             =   60
      Width           =   13365
      _ExtentX        =   23574
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ListView ListViewCodes 
      Height          =   5415
      Left            =   10740
      TabIndex        =   13
      Top             =   1080
      Visible         =   0   'False
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   9551
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Sample ID"
      Height          =   195
      Left            =   12150
      TabIndex        =   4
      Top             =   390
      Width           =   765
   End
End
Attribute VB_Name = "frmAudit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pTableName As String
Private pTableNameAudit As String

Private Dept As String

Private Sub LoadAudit()

          Dim sql As String
          Dim tb As Recordset
          Dim tbArc As Recordset
          Dim n As Integer
          Dim CurrentName() As String
          Dim current() As String
          Dim NameDisplayed As Boolean
          Dim AuditChanged As Boolean

10        On Error GoTo LoadAudit_Error

20        rtb.Text = ""
30        rtb.SelFontSize = 12

40        If Trim$(txtSampleID) = "" Then Exit Sub

50        rtb.SelFontSize = 16
60        rtb.SelColor = vbBlack
70        rtb.SelUnderline = True
80        rtb.SelBold = True
90        rtb.SelText = "Audit Trail for "
100       rtb.SelColor = vbRed
110       rtb.SelText = IIf(InStr(1, pTableName, "Coag"), CoagNameFor(ListView.SelectedItem.Text), ListView.SelectedItem.Text) & vbCrLf & vbCrLf
120       rtb.SelUnderline = False

130       rtb.SelFontSize = 12

140       sql = "SELECT * FROM " & pTableName & " WHERE " & _
                "SampleID = '" & txtSampleID & "' " & _
                "AND Code = '" & ListViewCodes.ListItems(ListView.SelectedItem.Index) & "' "
          '      "  (SELECT Distinct Code FROM " & Dept & "TestDefinitions WHERE " & _
                 '      "   ShortName = '" & ListView.SelectedItem.Text & "')"
150       If InStr(1, pTableName, "Coag") > 0 Then
160           sql = Replace(sql, "ShortName", "Code")
170       End If
180       Set tb = New Recordset
190       RecOpenServer 0, tb, sql
200       If tb.EOF Then
210           LoadNoCurrentRecord
220           Exit Sub
230       End If

240       AuditChanged = False

250       ReDim current(0 To tb.Fields.Count - 1)
260       ReDim CurrentName(0 To tb.Fields.Count - 1)
270       For n = 0 To tb.Fields.Count - 1
280           If Not tb.EOF Then
290               current(n) = tb.Fields(n).Value & ""
300           Else
310               current(n) = ""
320           End If
330           CurrentName(n) = tb.Fields(n).Name

340           sql = "SELECT ArchivedBy, ArchiveDateTime, [" & CurrentName(n) & "] FROM " & pTableNameAudit & " WHERE " & _
                    "SampleID = '" & txtSampleID & "' " & _
                    "AND Code = '" & ListViewCodes.ListItems(ListView.SelectedItem.Index) & "' " & _
                    "ORDER BY ArchiveDateTime DESC"
              '"  (SELECT DISTINCT Code FROM " & Dept & "TestDefinitions WHERE " & _
               '"   ShortName = '" & ListView.SelectedItem.Text & "') " &
350           If InStr(1, pTableNameAudit, "Coag") > 0 Then
360               sql = Replace(sql, "ShortName", "Code")
370           End If
380           Set tbArc = New Recordset
390           RecOpenServer 0, tbArc, sql
400           If tbArc.EOF Then
410               rtb.SelText = "No Changes Made."
420               AuditChanged = True
430               Exit For
440           Else
450               NameDisplayed = False
460               Do While Not tbArc.EOF
470                   If Trim$(current(n)) <> Trim$(tbArc.Fields(CurrentName(n)) & "") Then
480                       AuditChanged = True
490                       If Not NameDisplayed Then
500                           rtb.SelBold = True
510                           rtb.SelUnderline = False
520                           rtb.SelColor = vbBlue
530                           rtb.SelFontSize = 12
540                           rtb.SelText = CurrentName(n) & vbCrLf
550                           NameDisplayed = True
560                       End If
570                       rtb.SelFontSize = 12
580                       rtb.SelUnderline = False
590                       rtb.SelText = tbArc!ArchiveDateTime & " "
600                       rtb.SelColor = vbRed
610                       rtb.SelText = tbArc!ArchivedBy & ""
620                       rtb.SelColor = vbBlack
630                       rtb.SelText = " Changed "
640                       rtb.SelColor = vbGreen
650                       rtb.SelBold = True
660                       If Trim$(tbArc.Fields(CurrentName(n)) & "") = "" Then
670                           rtb.SelText = "<Blank> "
680                       Else
690                           rtb.SelText = Trim$(tbArc.Fields(CurrentName(n)))
700                       End If
710                       rtb.SelColor = vbBlack
720                       rtb.SelBold = False
730                       rtb.SelText = " to "
740                       rtb.SelBold = True
750                       If Trim$(current(n)) = "" Then
760                           rtb.SelText = "<Blank>" & vbCrLf
770                       Else
780                           rtb.SelText = Trim$(current(n)) & vbCrLf
790                       End If
800                       current(n) = Trim$(tbArc.Fields(CurrentName(n)) & "")
810                   End If
820                   tbArc.MoveNext
830               Loop
840           End If
850           If NameDisplayed Then
860               rtb.SelText = vbCrLf
870           End If
880       Next

890       If Not AuditChanged Then
900           rtb.SelText = "No Changes made"
910       End If

920       Exit Sub

LoadAudit_Error:

          Dim strES As String
          Dim intEL As Integer

930       intEL = Erl
940       strES = Err.Description
950       LogError "frmAudit", "LoadAudit", intEL, strES, sql

End Sub

Private Sub LoadNoCurrentRecord()

          Dim sql As String
          Dim tb As Recordset
          Dim tbArc As Recordset
          Dim n As Integer
          Dim CurrentName() As String
          Dim current() As String
          Dim NameDisplayed As Boolean
          Dim s As String
          Dim fld As Field
          Dim X As Integer
          Dim Y As Integer
          Dim ColumnADT As Integer

10        On Error GoTo LoadNoCurrentRecord_Error

20        rtb.Text = ""
30        rtb.SelFontSize = 12

40        If Trim$(txtSampleID) = "" Then Exit Sub

50        rtb.SelFontSize = 16
60        rtb.SelColor = vbBlack
70        rtb.SelUnderline = True
80        rtb.SelBold = True
90        rtb.SelText = "Audit Trail for "
100       rtb.SelColor = vbRed
110       rtb.SelText = IIf(InStr(1, pTableName, "Coag"), CoagNameFor(ListView.SelectedItem.Text), ListView.SelectedItem.Text) & vbCrLf & vbCrLf
120       rtb.SelUnderline = False

130       rtb.SelFontSize = 12

140       sql = "SELECT * FROM " & pTableName & "Audit " & _
                "WHERE SampleID = '" & txtSampleID & "' " & _
                "AND Code = '" & ListViewCodes.ListItems(ListView.SelectedItem.Index) & "' " & _
                "ORDER BY ArchiveDateTime DESC"
150       Set tb = New Recordset
160       Set tb = Cnxn(0).Execute(sql)
170       ReDim xy(1 To tb.Fields.Count, 0 To 0) As String
180       X = 0
190       For Each fld In tb.Fields
200           X = X + 1
210           xy(X, 0) = fld.Name
220           If fld.Name = "ArchiveDateTime" Then ColumnADT = X
230       Next

240       Y = 0
250       Do While Not tb.EOF
260           Y = Y + 1
270           ReDim Preserve xy(1 To tb.Fields.Count, 0 To Y)
280           X = 0
290           For Each fld In tb.Fields
300               X = X + 1
310               xy(X, Y) = tb.Fields(fld.Name) & ""
320           Next
330           tb.MoveNext
340       Loop

          Dim Differences As Boolean
          Dim Test As String
350       For X = 1 To UBound(xy, 1)
360           Differences = False
370           Test = xy(X, 1)
380           If UBound(xy, 2) = 1 Then
390               Differences = True
400           Else
410               For Y = 1 To UBound(xy, 2)
420                   If xy(X, Y) <> Test Then
430                       Differences = True
440                       Exit For
450                   End If
460                   If Differences Then
470                       Exit For
480                   End If
490               Next
500           End If

510           If Differences Then
520               s = "<BOLD 1><COLOUR RED><SIZE 14><UNDERLINE 1>" & xy(X, 0) & vbCrLf
530               Display s
540               For Y = 1 To UBound(xy, 2)
550                   s = "<COLOUR BLACK><UNDERLINE 0><SIZE 14>   " & xy(X, Y) & "   " & _
                          "<SIZE 10>{Retrieve record " & xy(ColumnADT, Y) & "}" & vbCrLf
560                   Display s
570               Next
580           End If
590       Next

          '  Set tbArc = New Recordset
          '  RecOpenServer 0, tbArc, sql
          '  If tbArc!Tot > 1 Then
          '    'differences
          '    ReDim current(0 To tbArc.Fields.Count - 1)
          '    ReDim CurrentName(0 To tbArc.Fields.Count - 1)
          '    For n = 0 To tbArc.Fields.Count - 1
          '      If Not tb.EOF Then
          '        current(n) = tb.Fields(n).Value & ""
          '      Else
          '        current(n) = ""
          '      End If
          '      CurrentName(n) = tb.Fields(n).Name
          '
          '      sql = "SELECT ArchivedBy, ArchiveDateTime, [" & CurrentName(n) & "] FROM " & pTableNameAudit & " WHERE " & _
                 '            "SampleID = '" & txtSampleID & "' " & _
                 '            "AND Code = '" & ListViewCodes.ListItems(ListView.SelectedItem.Index) & "' " & _
                 '            "ORDER BY ArchiveDateTime DESC"
          '      If InStr(1, pTableNameAudit, "Coag") > 0 Then
          '        sql = Replace(sql, "ShortName", "Code")
          '      End If
          '      Set tbArc = New Recordset
          '      RecOpenServer 0, tbArc, sql
          '      If tbArc.EOF Then
          '        rtb.SelText = "No Changes Made."
          '        Exit For
          '      Else
          '        NameDisplayed = False
          '        Do While Not tbArc.EOF
          '          If Trim$(current(n)) <> Trim$(tbArc.Fields(CurrentName(n)) & "") Then
          '            If Not NameDisplayed Then
          '              s = "<UNDERLINE 0><BOLD 1><COLOUR BLUE><SIZE 12>" & CurrentName(n) & vbCrLf
          '              Display s
          '              NameDisplayed = True
          '            End If
          '
          '            s = "<SIZE 12>" & tbArc!ArchiveDateTime & " " & _
                       '                "<COLOUR RED>" & tbArc!ArchivedBy & _
                       '                "<COLOUR BLACK> Changed " & _
                       '                "<COLOUR GREEN><BOLD 1>"
          '            If Trim$(tbArc.Fields(CurrentName(n)) & "") = "" Then
          '              s = s & "[Blank] "
          '            Else
          '              s = s & Trim$(tbArc.Fields(CurrentName(n)))
          '            End If
          '            s = s & "<COLOUR BLACK><BOLD 0> to <BOLD 1>"
          '            If Trim$(current(n)) = "" Then
          '              s = s & "[Blank]" & vbCrLf
          '            Else
          '              s = s & Trim$(current(n)) & vbCrLf
          '            End If
          '            Display s
          '            current(n) = Trim$(tbArc.Fields(CurrentName(n)) & "")
          '
          '          End If
          '          tbArc.MoveNext
          '        Loop
          '      End If
          '      If NameDisplayed Then
          '        rtb.SelText = vbCrLf
          '      End If
          '    Next
          '  End If
          'Next
600       Exit Sub

          '  rtb.SelText = "No Current Record found." & vbCrLf
          '    rtb.SelFontSize = 12
          '    rtb.SelFontSize = 12
          '    rtb.SelText = "{Retrieve record 15/11/2010 15:37:28}" & vbCrLf
          'End If

610       ReDim current(0 To tb.Fields.Count - 1)
620       ReDim CurrentName(0 To tb.Fields.Count - 1)
630       For n = 0 To tb.Fields.Count - 1
640           If Not tb.EOF Then
650               current(n) = tb.Fields(n).Value & ""
660           Else
670               current(n) = ""
680           End If
690           CurrentName(n) = tb.Fields(n).Name

700           sql = "SELECT ArchivedBy, ArchiveDateTime, [" & CurrentName(n) & "] FROM " & pTableNameAudit & " WHERE " & _
                    "SampleID = '" & txtSampleID & "' " & _
                    "AND Code = '" & ListViewCodes.ListItems(ListView.SelectedItem.Index) & "' " & _
                    "ORDER BY ArchiveDateTime DESC"
              '"  (SELECT DISTINCT Code FROM " & Dept & "TestDefinitions WHERE " & _
               '"   ShortName = '" & ListView.SelectedItem.Text & "') " &
710           If InStr(1, pTableNameAudit, "Coag") > 0 Then
720               sql = Replace(sql, "ShortName", "Code")
730           End If
740           Set tbArc = New Recordset
750           RecOpenServer 0, tbArc, sql
760           If tbArc.EOF Then
770               rtb.SelText = "No Changes Made."
780               Exit For
790           Else
800               NameDisplayed = False
810               Do While Not tbArc.EOF
820                   If Trim$(current(n)) <> Trim$(tbArc.Fields(CurrentName(n)) & "") Then
830                       If Not NameDisplayed Then
                              '460                       rtb.SelBold = True
                              '470                       rtb.SelColor = vbBlue
                              '480                       rtb.SelFontSize = 12
                              '490                       rtb.SelText = CurrentName(n) & vbCrLf
840                           s = "<UNDERLINE 0><BOLD 1><COLOUR BLUE><SIZE 12>" & CurrentName(n) & vbCrLf
850                           Display s
860                           NameDisplayed = True
870                       End If

880                       s = "<SIZE 12>" & tbArc!ArchiveDateTime & " " & _
                              "<COLOUR RED>" & tbArc!ArchivedBy & _
                              "<COLOUR BLACK> Changed " & _
                              "<COLOUR GREEN><BOLD 1>"
890                       If Trim$(tbArc.Fields(CurrentName(n)) & "") = "" Then
900                           s = s & "[Blank] "
910                       Else
920                           s = s & Trim$(tbArc.Fields(CurrentName(n)))
930                       End If
940                       s = s & "<COLOUR BLACK><BOLD 0> to <BOLD 1>"
950                       If Trim$(current(n)) = "" Then
960                           s = s & "[Blank]" & vbCrLf
970                       Else
980                           s = s & Trim$(current(n)) & vbCrLf
990                       End If
1000                      Display s
1010                      current(n) = Trim$(tbArc.Fields(CurrentName(n)) & "")


                          '                rtb.SelFontSize = 12
                          '                rtb.SelText = tbArc!ArchiveDateTime & " "
                          '                rtb.SelColor = vbRed
                          '                rtb.SelText = tbArc!ArchivedBy & ""
                          '                rtb.SelColor = vbBlack
                          '                rtb.SelText = " Changed "
                          '                rtb.SelColor = vbGreen
                          '                rtb.SelBold = True
                          '                If Trim$(tbArc.Fields(CurrentName(n)) & "") = "" Then
                          '                    rtb.SelText = "<Blank> "
                          '                Else
                          '                    rtb.SelText = Trim$(tbArc.Fields(CurrentName(n)))
                          '                End If
                          '                rtb.SelColor = vbBlack
                          '                rtb.SelBold = False
                          '                rtb.SelText = " to "
                          '                rtb.SelBold = True
                          '                If Trim$(current(n)) = "" Then
                          '                    rtb.SelText = "<Blank>" & vbCrLf
                          '                Else
                          '                    rtb.SelText = Trim$(current(n)) & vbCrLf
                          '                End If
1020                      current(n) = Trim$(tbArc.Fields(CurrentName(n)) & "")
1030                  End If
1040                  tbArc.MoveNext
1050              Loop
1060          End If
1070          If NameDisplayed Then
1080              rtb.SelText = vbCrLf
1090          End If
1100      Next

1110      Exit Sub

LoadNoCurrentRecord_Error:

          Dim strES As String
          Dim intEL As Integer

1120      intEL = Erl
1130      strES = Err.Description
1140      LogError "frmAudit", "LoadNoCurrentRecord", intEL, strES, sql

End Sub

Private Sub Display(ByVal s As String)

          Dim R() As String
          Dim n As Integer
          Dim c As Long
          Dim A As String
          Dim b As String
          Dim X As Integer

          's = "<BOLD 1><COLOUR BLUE><SIZE 12><UNDERLINE 0>" & CurrentName(n) & vbCrLf

10        R = Split(s, ">")

20        For n = 0 To UBound(R)
30            X = InStr(R(n), "<")
40            If X > 1 Then
50                A = Left$(R(n), X - 1)
60                b = Mid$(R(n), X + 1)
70            Else
80                A = R(n)
90                b = ""
100           End If

110           A = Replace(A, "<", "")
120           If Left$(A, 4) = "BOLD" Then
130               rtb.SelBold = IIf(Mid$(A, 6, 1) = "1", True, False)
140           ElseIf Left$(A, 6) = "COLOUR" Then
150               Select Case Mid$(A, 8)
                  Case "BLACK": c = vbBlack
160               Case "BLUE": c = vbBlue
170               Case "GREEN": c = vbGreen
180               Case "RED": c = vbRed
190               Case Else: c = vbBlack
200               End Select
210               rtb.SelColor = c
220           ElseIf Left$(A, 4) = "SIZE" Then
230               rtb.SelFontSize = Val(Mid$(A, 6))
240           ElseIf Left$(A, 9) = "UNDERLINE" Then
250               rtb.SelUnderline = IIf(Mid$(A, 11, 1) = "1", True, False)
260           Else
270               rtb.SelText = A
280           End If

290           If Left$(b, 4) = "BOLD" Then
300               rtb.SelBold = IIf(Mid$(b, 6, 1) = "1", True, False)
310           ElseIf Left$(b, 6) = "COLOUR" Then
320               Select Case Mid$(b, 8)
                  Case "BLACK": c = vbBlack
330               Case "BLUE": c = vbBlue
340               Case "GREEN": c = vbGreen
350               Case "RED": c = vbRed
360               Case Else: c = vbBlack
370               End Select
380               rtb.SelColor = c
390           ElseIf Left$(b, 4) = "SIZE" Then
400               rtb.SelFontSize = Val(Mid$(b, 6))
410           ElseIf Left$(b, 9) = "UNDERLINE" Then
420               rtb.SelUnderline = IIf(Mid$(b, 11, 1) = "1", True, False)
430           Else
440               rtb.SelText = b
450           End If
460       Next

End Sub

Private Sub SelectDisplay()

          Dim sql As String
          Dim tb As Recordset
          Dim clmX As ColumnHeader
          Dim itmX As ListItem
          Dim itmC As ListItem

10        On Error GoTo SelectDisplay_Error

20        rtb.TextRTF = ""

30        ListView.ListItems.Clear
40        Set clmX = ListView.ColumnHeaders.Add()
50        clmX.Text = "Parameter"

60        txtSampleID = Val(txtSampleID)
70        If Val(txtSampleID) = 0 Then Exit Sub

80        Select Case UCase$(pTableName)
          Case "BIORESULTS": Dept = "Bio"
90        Case "ENDRESULTS": Dept = "End"
100       Case "IMMRESULTS": Dept = "Imm"
110       Case "COAGRESULTS": Dept = "Coag"
120       End Select

130       If Dept <> "" Then

              '    sql = "SELECT DISTINCT ShortName, Code Cod, '65280' Colour FROM " & Dept & "TestDefinitions WHERE " & _
                   '          "  Code IN (  SELECT DISTINCT Code FROM " & Dept & "Results WHERE " & _
                   '          "             SampleID = '" & txtSampleID & "' ) " & _
                   '          "  AND Code NOT IN ( SELECT DISTINCT Code FROM " & Dept & "ResultsAudit WHERE " & _
                   '          "           SampleID = '" & txtSampleID & "' ) " & _
                   '          "UNION " & _
                   '          "SELECT DISTINCT ShortName, Code Cod, '255' Colour FROM " & Dept & "TestDefinitions WHERE " & _
                   '          "  Code IN ( SELECT DISTINCT Code FROM " & Dept & "ResultsAudit WHERE " & _
                   '          "            SampleID = '" & txtSampleID & "' ) " & _
                   '          "  AND Code NOT IN (  SELECT DISTINCT Code FROM " & Dept & "Results WHERE " & _
                   '          "             SampleID = '" & txtSampleID & "' ) " & _
                   '          "UNION " & _
                   '          "SELECT DISTINCT ShortName, Code Cod, '16711680' Colour FROM " & Dept & "TestDefinitions WHERE " & _
                   '          "  Code IN ( SELECT DISTINCT Code FROM " & Dept & "ResultsAudit WHERE " & _
                   '          "            SampleID = '" & txtSampleID & "' ) " & _
                   '          "  AND Code IN (  SELECT DISTINCT Code FROM " & Dept & "Results WHERE " & _
                   '          "             SampleID = '" & txtSampleID & "' ) " & _
                   '          "GROUP BY ShortName, Code ORDER BY ShortName"

140           sql = "SELECT DISTINCT T.ShortName, T.Code Cod, 'Colour' = CASE " & _
                    "WHEN R.Code Is Not Null AND A.Code Is Null THEN '65280' " & _
                    "WHEN R.Code Is Null AND A.Code Is Not Null THEN '255' " & _
                    "WHEN R.Code Is Not Null AND A.Code Is not Null THEN '16711680' END " & _
                    "FROM " & Dept & "TestDefinitions T " & _
                    "LEFT JOIN " & Dept & "Results R On T.Code = R.Code And R.SampleID = '" & txtSampleID & "' " & _
                    "LEFT JOIN " & Dept & "ResultsAudit A On T.Code = A.Code And A.SampleID = '" & txtSampleID & "' " & _
                    "Where (R.Code Is Not Null And A.Code Is Null) " & _
                    "OR (R.Code Is Null And A.Code Is Not Null) " & _
                    "OR (R.Code Is Not Null AND A.Code Is not Null) " & _
                    "ORDER BY T.ShortName"

150           If Dept = "Coag" Then
160               sql = Replace(sql, "ShortName", "Code")
170           End If
180           Set tb = New Recordset
190           RecOpenServer 0, tb, sql
200           If Not tb.EOF Then
210               Do While Not tb.EOF

220                   Set itmX = ListView.ListItems.Add()
230                   Set itmC = ListViewCodes.ListItems.Add()
240                   If Dept = "Coag" Then
250                       itmX.Text = CoagNameFor(tb!Code) & ""
260                   Else
270                       itmX.Text = tb!ShortName & ""
280                   End If
290                   itmC.Text = tb!Cod & ""
300                   'itmX.ForeColor = tb!Colour

310                   tb.MoveNext
320               Loop
330               If ListView.ListItems.Count > 0 Then
340                   ListView.ListItems(1).Selected = True
350                   LoadAudit
360               End If
370           Else
380               rtb.SelText = "No Record."
390           End If

400       End If

410       Exit Sub

SelectDisplay_Error:

          Dim strES As String
          Dim intEL As Integer

420       intEL = Erl
430       strES = Err.Description
440       LogError "frmAudit", "SelectDisplay", intEL, strES, sql

End Sub

Private Sub cmdCancel_Click()

10        Unload Me

End Sub


Private Sub Form_Activate()

10        SelectDisplay

End Sub
Public Property Let TableName(ByVal sNewValue As String)

10        pTableName = sNewValue
20        pTableNameAudit = sNewValue & "Audit"

End Property
Public Property Let SampleID(ByVal sNewValue As String)

10        txtSampleID = sNewValue

End Property

Private Sub ListView_Click()

10        LoadAudit

End Sub

Private Sub ListView_KeyPress(KeyAscii As Integer)

10        KeyAscii = 0

End Sub


Private Sub rtb_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

          Dim sStart As Long
          Dim sStop As Long
          Dim current As Long
          Dim sql As String
          Dim adt As String
          Dim SelectList As String

10        current = rtb.SelStart

20        rtb.upto "{", False, False
30        sStart = rtb.SelStart

40        rtb.upto "}", True, False
50        sStop = rtb.SelStart

60        rtb.SelStart = sStart
70        rtb.SelLength = sStop - sStart
80        If rtb.SelLength <> 35 Then
90            rtb.SelLength = 0
100       End If

110       If current < sStart Or current > sStop Then
120           rtb.SelLength = 0
130       End If

140       adt = Format$(Mid$(rtb.SelText, 17), "dd/MMM/yyyy HH:nn:ss")

150       If rtb.SelLength = 35 Then
160           Screen.MousePointer = vbDefault

170           Select Case UCase$(pTableName)
              Case "BIORESULTS", "ENDRESULTS"
180               SelectList = "SampleID, Code, Result, Valid, Printed, RunTime, RunDate, Operator, Flags, Units, SampleType, " & _
                               "Analyser, Faxed, Authorised, Comment, PC, HealthLink"
190           Case "IMMRESULTS"
200               SelectList = "SampleID, Code, Result, Valid, Printed, RunTime, RunDate, Operator, Flags, Units, SampleType, " & _
                               "Analyser, Faxed, Authorised, Comment, PC, HealthLink, Method"
210           Case "COAGRESULTS"
220               SelectList = "SampleID, Code, Units, Result, Valid, Printed, RunTime, RunDate, UserName, " & _
                               "Analyser, Faxed, Authorised, Released, HealthLink"
230           End Select

240           If iMsg(rtb.SelText, vbYesNo) = vbYes Then
250               sql = "IF NOT EXISTS ( SELECT * FROM " & pTableName & " " & _
                        "                WHERE SampleID = '" & txtSampleID & "' " & _
                        "                AND Code = '" & ListViewCodes.ListItems(ListView.SelectedItem.Index) & "') " & _
                        "  INSERT INTO " & pTableName & " " & _
                        "  ( " & SelectList & ") " & _
                        "  SELECT TOP 1 " & SelectList & " " & _
                        "  FROM " & pTableName & "Audit " & _
                        "  WHERE ArchiveDateTime BETWEEN '" & Format$(DateAdd("s", -1, adt), "dd/MMM/yyyy HH:nn:ss") & "' AND '" & Format$(DateAdd("s", 1, adt), "dd/MMM/yyyy HH:nn:ss") & "' AND Code = '" & ListViewCodes.ListItems(ListView.SelectedItem.Index) & "'"
260               Cnxn(0).Execute sql

270           End If
280       End If

End Sub

Private Sub txtSampleID_LostFocus()

10        rtb.TextRTF = ""

20        SelectDisplay

End Sub

