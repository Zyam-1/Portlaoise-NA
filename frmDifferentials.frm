VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmDifferentials 
   Caption         =   "NetAcquire - Differentials"
   ClientHeight    =   5835
   ClientLeft      =   1455
   ClientTop       =   1875
   ClientWidth     =   7485
   Icon            =   "frmDifferentials.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   7485
   Begin VB.CommandButton bSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   795
      Left            =   5355
      Picture         =   "frmDifferentials.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   74
      ToolTipText     =   "Save Changes"
      Top             =   2340
      Width           =   1785
   End
   Begin VB.CommandButton bCancel 
      Appearance      =   0  'Flat
      Caption         =   "Exit"
      Height          =   795
      Left            =   5355
      Picture         =   "frmDifferentials.frx":0316
      Style           =   1  'Graphical
      TabIndex        =   73
      Tag             =   "101"
      ToolTipText     =   "Exit Screen"
      Top             =   3240
      Width           =   1785
   End
   Begin VB.CommandButton bClear 
      Caption         =   "Clear All Results"
      Height          =   795
      Left            =   5355
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   1020
      Width           =   1785
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   5160
      TabIndex        =   65
      Top             =   4260
      Width           =   2115
      Begin VB.CheckBox cPrint 
         Caption         =   "Print"
         Height          =   225
         Left            =   60
         TabIndex        =   70
         Top             =   1110
         Width           =   1665
      End
      Begin Threed.SSOption oEnter 
         Height          =   225
         Index           =   1
         Left            =   90
         TabIndex        =   67
         Top             =   540
         Width           =   1965
         _Version        =   65536
         _ExtentX        =   3466
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Change keys or wording"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption oEnter 
         Height          =   225
         Index           =   0
         Left            =   60
         TabIndex        =   66
         Top             =   210
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Enter Results"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   -1  'True
      End
      Begin Threed.SSCheck oSaveSettings 
         Height          =   195
         Left            =   60
         TabIndex        =   69
         Top             =   840
         Width           =   1665
         _Version        =   65536
         _ExtentX        =   2937
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "Save Key Settings"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   -1  'True
      End
   End
   Begin VB.TextBox tCount 
      BackColor       =   &H80000018&
      Height          =   285
      Index           =   29
      Left            =   4230
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   64
      Top             =   5100
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox tKey 
      Height          =   285
      Index           =   14
      Left            =   300
      MaxLength       =   1
      TabIndex        =   63
      Top             =   5100
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox tCount 
      BackColor       =   &H80000018&
      Height          =   285
      Index           =   28
      Left            =   4230
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   62
      Top             =   4770
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox tKey 
      Height          =   285
      Index           =   13
      Left            =   300
      MaxLength       =   1
      TabIndex        =   61
      Top             =   4770
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox tCount 
      BackColor       =   &H80000018&
      Height          =   285
      Index           =   27
      Left            =   4230
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   60
      Top             =   4440
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox tKey 
      Height          =   285
      Index           =   12
      Left            =   300
      MaxLength       =   1
      TabIndex        =   59
      Top             =   4440
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox tCount 
      BackColor       =   &H80000018&
      Height          =   285
      Index           =   26
      Left            =   4230
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   58
      Top             =   4110
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox tKey 
      Height          =   285
      Index           =   11
      Left            =   300
      MaxLength       =   1
      TabIndex        =   57
      Top             =   4110
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox tCount 
      BackColor       =   &H80000018&
      Height          =   285
      Index           =   25
      Left            =   4230
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   56
      Top             =   3780
      Width           =   705
   End
   Begin VB.TextBox tKey 
      Height          =   285
      Index           =   10
      Left            =   300
      MaxLength       =   1
      TabIndex        =   55
      Top             =   3780
      Width           =   525
   End
   Begin VB.TextBox tCount 
      BackColor       =   &H80000018&
      Height          =   285
      Index           =   24
      Left            =   4230
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   54
      Top             =   3450
      Width           =   705
   End
   Begin VB.TextBox tCount 
      BackColor       =   &H80000018&
      Height          =   285
      Index           =   23
      Left            =   4230
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   53
      Top             =   3120
      Width           =   705
   End
   Begin VB.TextBox tCount 
      BackColor       =   &H80000018&
      Height          =   285
      Index           =   22
      Left            =   4230
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   52
      Top             =   2790
      Width           =   705
   End
   Begin VB.TextBox tCount 
      BackColor       =   &H80000018&
      Height          =   285
      Index           =   21
      Left            =   4230
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   51
      Top             =   2460
      Width           =   705
   End
   Begin VB.TextBox tCount 
      BackColor       =   &H80000018&
      Height          =   285
      Index           =   20
      Left            =   4230
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   50
      Top             =   2130
      Width           =   705
   End
   Begin VB.TextBox tCell 
      Height          =   285
      Index           =   14
      Left            =   930
      MaxLength       =   20
      TabIndex        =   49
      Top             =   5100
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.TextBox tCell 
      Height          =   285
      Index           =   13
      Left            =   930
      TabIndex        =   48
      Top             =   4770
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.TextBox tCell 
      Height          =   285
      Index           =   12
      Left            =   930
      TabIndex        =   47
      Top             =   4440
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.TextBox tCell 
      Height          =   285
      Index           =   11
      Left            =   930
      TabIndex        =   46
      Top             =   4110
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.TextBox tCell 
      Height          =   285
      Index           =   10
      Left            =   930
      TabIndex        =   45
      Top             =   3780
      Width           =   2265
   End
   Begin VB.TextBox tCell 
      Height          =   285
      Index           =   9
      Left            =   930
      TabIndex        =   44
      Top             =   3450
      Width           =   2265
   End
   Begin VB.TextBox tCell 
      Height          =   285
      Index           =   8
      Left            =   930
      TabIndex        =   43
      Top             =   3120
      Width           =   2265
   End
   Begin VB.TextBox tCell 
      Height          =   285
      Index           =   7
      Left            =   930
      TabIndex        =   42
      Top             =   2790
      Width           =   2265
   End
   Begin VB.TextBox tCell 
      Height          =   285
      Index           =   6
      Left            =   930
      TabIndex        =   41
      Top             =   2460
      Width           =   2265
   End
   Begin VB.TextBox tCell 
      Height          =   285
      Index           =   5
      Left            =   930
      TabIndex        =   40
      Text            =   "Lucocytes"
      Top             =   2130
      Width           =   2265
   End
   Begin VB.TextBox tCell 
      BackColor       =   &H80000018&
      Height          =   285
      Index           =   4
      Left            =   930
      TabIndex        =   39
      Text            =   "Basophils"
      Top             =   1800
      Width           =   2265
   End
   Begin VB.TextBox tCell 
      BackColor       =   &H80000018&
      Height          =   285
      Index           =   3
      Left            =   930
      TabIndex        =   38
      Text            =   "Eosinophils"
      Top             =   1470
      Width           =   2265
   End
   Begin VB.TextBox tCell 
      BackColor       =   &H80000018&
      Height          =   285
      Index           =   2
      Left            =   930
      TabIndex        =   37
      Text            =   "Monocytes"
      Top             =   1140
      Width           =   2265
   End
   Begin VB.TextBox tCell 
      BackColor       =   &H80000018&
      Height          =   285
      Index           =   1
      Left            =   930
      TabIndex        =   36
      Text            =   "Lymphocytes"
      Top             =   810
      Width           =   2265
   End
   Begin VB.TextBox tCell 
      BackColor       =   &H80000018&
      Height          =   285
      Index           =   0
      Left            =   930
      TabIndex        =   35
      Text            =   "Neutrophils"
      Top             =   480
      Width           =   2265
   End
   Begin VB.TextBox tCount 
      BackColor       =   &H80000018&
      Height          =   285
      Index           =   19
      Left            =   4230
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   31
      Top             =   1800
      Width           =   705
   End
   Begin VB.TextBox tCount 
      BackColor       =   &H80000018&
      Height          =   285
      Index           =   18
      Left            =   4230
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   30
      Top             =   1470
      Width           =   705
   End
   Begin VB.TextBox tCount 
      BackColor       =   &H80000018&
      Height          =   285
      Index           =   17
      Left            =   4230
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   29
      Top             =   1140
      Width           =   705
   End
   Begin VB.TextBox tCount 
      BackColor       =   &H80000018&
      Height          =   285
      Index           =   16
      Left            =   4230
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   28
      Top             =   810
      Width           =   705
   End
   Begin VB.TextBox tCount 
      BackColor       =   &H80000018&
      Height          =   285
      Index           =   15
      Left            =   4230
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   27
      Top             =   480
      Width           =   705
   End
   Begin VB.TextBox tCount 
      BackColor       =   &H80000018&
      Height          =   285
      Index           =   14
      Left            =   3270
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   26
      Top             =   5100
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox tCount 
      BackColor       =   &H80000018&
      Height          =   285
      Index           =   13
      Left            =   3270
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   25
      Top             =   4770
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox tCount 
      BackColor       =   &H80000018&
      Height          =   285
      Index           =   12
      Left            =   3270
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   24
      Top             =   4440
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox tCount 
      BackColor       =   &H80000018&
      Height          =   285
      Index           =   11
      Left            =   3270
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   23
      Top             =   4110
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox tCount 
      BackColor       =   &H80000018&
      Height          =   285
      Index           =   10
      Left            =   3270
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   22
      Top             =   3780
      Width           =   705
   End
   Begin VB.TextBox tKey 
      Height          =   285
      Index           =   9
      Left            =   300
      MaxLength       =   1
      TabIndex        =   19
      Top             =   3450
      Width           =   525
   End
   Begin VB.TextBox tCount 
      BackColor       =   &H80000018&
      Height          =   285
      Index           =   9
      Left            =   3270
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   18
      Top             =   3450
      Width           =   705
   End
   Begin VB.TextBox tKey 
      Height          =   285
      Index           =   8
      Left            =   300
      MaxLength       =   1
      TabIndex        =   17
      Top             =   3120
      Width           =   525
   End
   Begin VB.TextBox tCount 
      BackColor       =   &H80000018&
      Height          =   285
      Index           =   8
      Left            =   3270
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   16
      Top             =   3120
      Width           =   705
   End
   Begin VB.TextBox tKey 
      Height          =   285
      Index           =   7
      Left            =   300
      MaxLength       =   1
      TabIndex        =   15
      Top             =   2790
      Width           =   525
   End
   Begin VB.TextBox tCount 
      BackColor       =   &H80000018&
      Height          =   285
      Index           =   7
      Left            =   3270
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   14
      Top             =   2790
      Width           =   705
   End
   Begin VB.TextBox tKey 
      Height          =   285
      Index           =   6
      Left            =   300
      MaxLength       =   1
      TabIndex        =   13
      Top             =   2460
      Width           =   525
   End
   Begin VB.TextBox tCount 
      BackColor       =   &H80000018&
      Height          =   285
      Index           =   6
      Left            =   3270
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   12
      Top             =   2460
      Width           =   705
   End
   Begin VB.TextBox tKey 
      Height          =   285
      Index           =   5
      Left            =   300
      MaxLength       =   1
      TabIndex        =   11
      Top             =   2130
      Width           =   525
   End
   Begin VB.TextBox tCount 
      BackColor       =   &H80000018&
      Height          =   285
      Index           =   5
      Left            =   3270
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   10
      Top             =   2130
      Width           =   705
   End
   Begin VB.TextBox tKey 
      Height          =   285
      Index           =   4
      Left            =   300
      MaxLength       =   1
      TabIndex        =   9
      Top             =   1800
      Width           =   525
   End
   Begin VB.TextBox tCount 
      BackColor       =   &H80000018&
      Height          =   285
      Index           =   4
      Left            =   3270
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   8
      Top             =   1800
      Width           =   705
   End
   Begin VB.TextBox tKey 
      Height          =   285
      Index           =   3
      Left            =   300
      MaxLength       =   1
      TabIndex        =   7
      Top             =   1470
      Width           =   525
   End
   Begin VB.TextBox tCount 
      BackColor       =   &H80000018&
      Height          =   285
      Index           =   3
      Left            =   3270
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   6
      Top             =   1470
      Width           =   705
   End
   Begin VB.TextBox tKey 
      Height          =   285
      Index           =   2
      Left            =   300
      MaxLength       =   1
      TabIndex        =   5
      Top             =   1140
      Width           =   525
   End
   Begin VB.TextBox tCount 
      BackColor       =   &H80000018&
      Height          =   285
      Index           =   2
      Left            =   3270
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   4
      Top             =   1140
      Width           =   705
   End
   Begin VB.TextBox tKey 
      Height          =   285
      Index           =   1
      Left            =   300
      MaxLength       =   1
      TabIndex        =   3
      Top             =   810
      Width           =   525
   End
   Begin VB.TextBox tCount 
      BackColor       =   &H80000018&
      Height          =   285
      Index           =   1
      Left            =   3270
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   2
      Top             =   810
      Width           =   705
   End
   Begin VB.TextBox tKey 
      Height          =   285
      Index           =   0
      Left            =   300
      MaxLength       =   1
      TabIndex        =   1
      Top             =   480
      Width           =   525
   End
   Begin VB.TextBox tCount 
      BackColor       =   &H80000018&
      Height          =   285
      Index           =   0
      Left            =   3270
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   0
      Top             =   480
      Width           =   705
   End
   Begin VB.Label lblOper 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2220
      TabIndex        =   72
      Top             =   5490
      Width           =   2055
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Differentiated By :"
      Height          =   255
      Left            =   840
      TabIndex        =   71
      Top             =   5490
      Width           =   1365
   End
   Begin VB.Label lWBC 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5610
      TabIndex        =   34
      Top             =   660
      Width           =   1095
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "WBC"
      Height          =   195
      Left            =   5610
      TabIndex        =   33
      Top             =   450
      Width           =   375
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Count #"
      Height          =   195
      Left            =   4320
      TabIndex        =   32
      Top             =   180
      Width           =   570
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Count %"
      Height          =   195
      Left            =   3330
      TabIndex        =   21
      Top             =   180
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Key"
      Height          =   195
      Left            =   420
      TabIndex        =   20
      Top             =   180
      Width           =   270
   End
End
Attribute VB_Name = "frmDifferentials"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Total As Long

Private pLoadDiff As Boolean

Private mSampleID As String

Private Sub bcancel_Click()

10        Unload Me

End Sub

Private Sub bClear_Click()

          Dim n As Long

10        On Error GoTo bClear_Click_Error

20        For n = 0 To 29
30            tCount(n) = ""
40        Next

50        Exit Sub

bClear_Click_Error:

          Dim strES As String
          Dim intEL As Integer

60        intEL = Erl
70        strES = Err.Description
80        LogError "frmDifferentials", "bClear_Click", intEL, strES

End Sub

Private Sub bSave_Click()

          Dim n As Long
          Dim sql As String
          Dim s As Long

10        On Error GoTo bSave_Click_Error

20        s = 0

30        For n = 0 To 14
40            s = s + Val(Trim(tCount(n)))
50        Next

60        If s = 0 Then
70            iMsg "You must enter a" & vbCrLf & "Differential first."
80            Exit Sub
90        End If

100       If s < 100 Then
110           If iMsg("Diff less than 100%. Continue", vbYesNo) = vbNo Then
120               Exit Sub
130           End If
140       End If

150       If oSaveSettings Then
160           sql = "If Exists(Select 1 From DifferentialTitles) " & _
                    "Begin " & _
                    "Update DifferentialTitles Set " & _
                    "K0 = '@K00', " & _
                    "K1 = '@K11', " & _
                    "K2 = '@K22', " & _
                    "K3 = '@K33', " & _
                    "K4 = '@K44', " & _
                    "K5 = '@K55', " & _
                    "K6 = '@K66', " & _
                    "K7 = '@K77', " & _
                    "K8 = '@K88', " & _
                    "K9 = '@K99', " & _
                    "K10 = '@K1010', " & _
                    "K11 = '@K1111', " & _
                    "K12 = '@K1212', " & _
                    "K13 = '@K1313', " & _
                    "K14 = '@K1414', "

170           sql = sql & _
                    "C5 = '@C515', " & _
                    "C6 = '@C616', " & _
                    "C7 = '@C717', " & _
                    "C8 = '@C818', " & _
                    "C9 = '@C919', " & _
                    "C10 = '@C1020', " & _
                    "C11 = '@C1121', " & _
                    "C12 = '@C1222', " & _
                    "C13 = '@C1323', " & _
                    "C14 = '@C1424' "
180           sql = sql & _
                    "End  " & _
                    "Else " & _
                    "Begin  " & _
                    "Insert Into DifferentialTitles (K0, K1, K2, K3, K4, K5, K6, K7, K8, K9, K10, K11, K12, K13, K14, C5, C6, " & _
                    "C7, C8, C9, C10, C11, C12, C13, C14) Values " & _
                    "('@K00', '@K11', '@K22', '@K33', '@K44', '@K55', '@K66', '@K77', '@K88', '@K99', '@K1010', '@K1111', " & _
                    "'@K1212', '@K1313', '@K1414', '@C515', '@C616', '@C717', '@C818', '@C919', '@C1020', '@C1121', '@C1222', " & _
                    "'@C1323', '@C1424') " & _
                    "End"

              'Please don't mind the code. it works. :)  (Babar Shahzad)

190           For n = 14 To 0 Step -1
200               sql = Replace(sql, "@K" & n & n, tKey(n).Text)
210           Next n
220           For n = 14 To 5 Step -1
230               sql = Replace(sql, "@C" & (n) & n + 10, tCell(n))
240           Next n


250           Cnxn(0).Execute sql


              '************Cursor operation conflict error. Old Code
              '    sql = "SELECT * from DifferentialTitles"
              '    Set tb = New Recordset
              '    RecOpenServer 0, tb, sql
              '    If tb.EOF Then tb.AddNew
              '    For n = 0 To 14
              '        tb("K" & Format(n)) = tKey(n)
              '    Next
              '    For n = 5 To 14
              '        tb("C" & Format(n)) = tCell(n)
              '    Next
              '    tb.Update
              '*******************************************
260       End If

270       sql = "DELETE from differentials WHERE runnumber = '" & frmEditAll.txtSampleID & "'"
280       Cnxn(0).Execute sql



          '*********************************************
          'sql = "SELECT * from Differentials WHERE " & _
           '      "RunNumber = '" & frmEditAll.txtSampleID & "'"
          'Set tb = New Recordset
          'RecOpenServer 0, tb, sql
          'If tb.EOF Then
          '    tb.AddNew
          'End If
          'tb!RunNumber = frmEditAll.txtSampleID
          'tb!Operator = UserCode
          'If cPrint.Value = 1 Then tb!prndiff = True Else tb!prndiff = False
          'For n = 0 To 14
          '    If Trim(tCell(n)) <> "" Then
          '        tb("Key" & Format(n)) = tKey(n)
          '        tb("Wording" & Format(n)) = tCell(n)
          '        tb("P" & Format(n)) = Val(tCount(n))
          '        tb("A" & Format(n)) = Val(tCount(n + 15))
          '    End If
          'Next
          'tb.Update
          '**************************************************

290       sql = "If Exists(Select 1 From Differentials " & _
                "Where RunNumber = @RunNumber0 ) " & _
                "Begin " & _
                "Update Differentials Set " & _
                "RunNumber = @RunNumber0, " & _
                "Key0 = '@Key01', " & _
                "Key1 = '@Key12', " & _
                "Key2 = '@Key23', " & _
                "Key3 = '@Key34', " & _
                "Key4 = '@Key45', " & _
                "Key5 = '@Key56', " & _
                "Key6 = '@Key67', " & _
                "Key7 = '@Key78', " & _
                "Key8 = '@Key89', " & _
                "Key9 = '@Key910', " & _
                "Key10 = '@Key1011', " & _
                "Key11 = '@Key1112', " & _
                "Key12 = '@Key1213', " & _
                "Key13 = '@Key1314', " & _
    "Key14 = '@Key1415', "
300       sql = sql & _
                "Wording0 = '@Wording016', " & _
                "Wording1 = '@Wording117', " & _
                "Wording2 = '@Wording218', " & _
                "Wording3 = '@Wording319', " & _
                "Wording4 = '@Wording420', " & _
                "Wording5 = '@Wording521', " & _
                "Wording6 = '@Wording622', " & _
                "Wording7 = '@Wording723', " & _
                "Wording8 = '@Wording824', " & _
                "Wording9 = '@Wording925', " & _
                "Wording10 = '@Wording1026', " & _
                "Wording11 = '@Wording1127', " & _
                "Wording12 = '@Wording1228', " & _
                "Wording13 = '@Wording1329', " & _
    "Wording14 = '@Wording1430', "
310       sql = sql & _
                "P0 = @P031, " & _
                "P1 = @P132, " & _
                "P2 = @P233, " & _
                "P3 = @P334, " & _
                "P4 = @P435, " & _
                "P5 = @P536, " & _
                "P6 = @P637, " & _
                "P7 = @P738, " & _
                "P8 = @P839, " & _
                "P9 = @P940, " & _
                "P10 = @P1041, " & _
                "P11 = @P1142, " & _
                "P12 = @P1243, " & _
                "P13 = @P1344, " & _
    "P14 = @P1445, "
320       sql = sql & _
                "A0 = @A046, " & _
                "A1 = @A147, " & _
                "A2 = @A248, " & _
                "A3 = @A349, " & _
                "A4 = @A450, " & _
                "A5 = @A551, " & _
                "A6 = @A652, " & _
                "A7 = @A753, " & _
                "A8 = @A854, " & _
                "A9 = @A955, " & _
                "A10 = @A1056, " & _
                "A11 = @A1157, " & _
                "A12 = @A1258, " & _
                "A13 = @A1359, " & _
    "A14 = @A1460, "
330       sql = sql & _
                "prndiff = @prndiff61, " & _
                "Operator = '@Operator62' " & _
                "Where RunNumber = @RunNumber0  " & _
                "End  " & _
                "Else " & _
                "Begin  " & _
                "Insert Into Differentials (RunNumber, Key0, Key1, Key2, Key3, Key4, Key5, Key6, Key7, Key8, Key9, Key10, Key11, " & _
                "Key12, Key13, Key14, Wording0, Wording1, Wording2, Wording3, Wording4, Wording5, Wording6, Wording7, Wording8, " & _
                "Wording9, Wording10, Wording11, Wording12, Wording13, Wording14, P0, P1, P2, P3, P4, P5, P6, P7, P8, P9, P10, " & _
                "P11, P12, P13, P14, A0, A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, A12, A13, A14, prndiff, Operator) Values " & _
                "(@RunNumber0, '@Key01', '@Key12', '@Key23', '@Key34', '@Key45', '@Key56', '@Key67', '@Key78', '@Key89', '@Key910', " & _
                "'@Key1011', '@Key1112', '@Key1213', '@Key1314', '@Key1415', '@Wording016', '@Wording117', '@Wording218', " & _
                "'@Wording319', '@Wording420', '@Wording521', '@Wording622', '@Wording723', '@Wording824', '@Wording925', " & _
                "'@Wording1026', '@Wording1127', '@Wording1228', '@Wording1329', '@Wording1430', @P031, @P132, @P233, @P334, @P435, " & _
                "@P536, @P637, @P738, @P839, @P940, @P1041, @P1142, @P1243, @P1344, @P1445, @A046, @A147, @A248, @A349, @A450, " & _
                "@A551, @A652, @A753, @A854, @A955, @A1056, @A1157, @A1258, @A1359, @A1460, @prndiff61, '@Operator62') " & _
                "End"


340       sql = Replace(sql, "@RunNumber0", frmEditAll.txtSampleID)
350       For n = 14 To 0 Step -1
360           If Trim(tCell(n).Text) <> "" Then
370               sql = Replace(sql, "@Key" & n & (n + 1), tKey(n))
380               sql = Replace(sql, "@Wording" & n & (n + 16), tCell(n))
390               sql = Replace(sql, "@P" & n & (n + 31), Val(tCount(n)))
400               sql = Replace(sql, "@A" & n & (n + 46), Val(tCount(n + 15)))
410           Else
420               sql = Replace(sql, "'@Key" & n & (n + 1) & "'", "Null")
430               sql = Replace(sql, "'@Wording" & n & (n + 16) & "'", "Null")
440               sql = Replace(sql, "@P" & n & (n + 31), "Null")
450               sql = Replace(sql, "@A" & n & (n + 46), "Null")
460           End If
470       Next n
480       sql = Replace(sql, "@prndiff61", IIf(cPrint.Value = 1, 1, 0))
490       sql = Replace(sql, "@Operator62", UserCode)

500       Cnxn(0).Execute sql

510       frmEditAll.bFilm.BackColor = vbBlue
520       Unload Me

530       Exit Sub

bSave_Click_Error:

          Dim strES As String
          Dim intEL As Integer

540       intEL = Erl
550       strES = Err.Description
560       LogError "frmDifferentials", "bSave_Click", intEL, strES

End Sub





Private Sub Form_KeyPress(KeyAscii As Integer)

          Dim n As Long



10        On Error GoTo Form_KeyPress_Error

20        If Total = 100 Then
30            Beep
40            KeyAscii = 0
50            Exit Sub
60        End If

70        If oEnter(0) Then
80            For n = 0 To 14
90                If UCase(Chr(KeyAscii)) = tKey(n) Then
100                   tCount(n) = Format(Val(tCount(n)) + 1)
110                   tCount(n + 15) = Format(Val(tCount(n)) * Val(lWBC) / 100, "0.0")
120                   Exit For
130               End If
140           Next
150           KeyAscii = 0

160           Total = 0
170           For n = 0 To 14
180               Total = Total + Val(tCount(n))
190           Next
200           If Total = 100 Then
210               Beep
220           End If
230       End If

240       Exit Sub

Form_KeyPress_Error:

          Dim strES As String
          Dim intEL As Integer

250       intEL = Erl
260       strES = Err.Description
270       LogError "frmDifferentials", "Form_KeyPress", intEL, strES

End Sub

Private Sub Form_Load()

          Dim n As Long
          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo Form_Load_Error

20        LoadDefaultDiffTitles

30        If pLoadDiff Then
40            sql = "SELECT * from Differentials WHERE " & _
                    "RunNumber = '" & mSampleID & "'"
50            Set tb = New Recordset
60            RecOpenServer 0, tb, sql
70            If Not tb.EOF Then
80                lblOper = TechNameFor(tb!Operator & "")
90                If tb!prndiff = True Then cPrint.Value = 1 Else cPrint.Value = 0
100               For n = 0 To 14
110                   tKey(n) = Trim(tb("Key" & Format(n)) & "")
120                   tCell(n) = Trim(tb("Wording" & Format(n)) & "")
130                   tCount(n) = IIf(Val(tb("P" & Format(n)) & "") = 0, "", tb("P" & Format(n)))
140                   tCount(n + 15) = IIf(Val(tb("A" & Format(n)) & "") = 0, "", tb("A" & Format(n)))
150               Next
160           End If
170       End If

180       If frmEditAll.bValidateHaem.Caption = "VALID" Then
190           bsave.Enabled = False
200           Frame1.Enabled = False
210           bclear.Enabled = False
220           For n = 0 To 14
230               tCount(n).Enabled = False
240               tKey(n).Enabled = False
250               tCell(n).Enabled = False
260           Next
270           For n = 15 To 29
280               tCount(n).Enabled = False
290           Next
300       Else
310           bsave.Enabled = True
320           Frame1.Enabled = True
330           bclear.Enabled = True
340       End If

350       Exit Sub

Form_Load_Error:

          Dim strES As String
          Dim intEL As Integer

360       intEL = Erl
370       strES = Err.Description
380       LogError "frmDifferentials", "Form_Load", intEL, strES, sql

End Sub

Private Sub Form_Unload(Cancel As Integer)

10        Total = 0
20        pLoadDiff = False

End Sub

Private Sub LoadDefaultDiffTitles()

          Dim n As Long
          Dim tb As New Recordset
          Dim sql As String

10        On Error GoTo LoadDefaultDiffTitles_Error

20        For n = 0 To 29
30            tCount(n) = ""
40        Next

50        sql = "SELECT * from DifferentialTitles"
60        Set tb = New Recordset
70        RecOpenServer 0, tb, sql
80        If Not tb.EOF Then
90            For n = 0 To 14
100               tKey(n) = Trim(tb("K" & Format(n)))
110           Next
120           For n = 5 To 14
130               tCell(n) = Trim(tb("C" & Format(n)))
140           Next
150       End If

160       Total = 0

170       Exit Sub

LoadDefaultDiffTitles_Error:

          Dim strES As String
          Dim intEL As Integer

180       intEL = Erl
190       strES = Err.Description
200       LogError "frmDifferentials", "LoadDefaultDiffTitles", intEL, strES

End Sub

Public Property Let LoadDiff(ByVal b As Boolean)

10        pLoadDiff = b

End Property

Public Property Let SampleID(ByVal SID As String)

10        mSampleID = SID

End Property

Private Sub tKey_Change(Index As Integer)

10        On Error GoTo tKey_Change_Error

20        tKey(Index) = UCase(tKey(Index))

30        Exit Sub

tKey_Change_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "frmDifferentials", "tKey_Change", intEL, strES

End Sub
