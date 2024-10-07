VERSION 5.00
Begin VB.Form frmUpdatePrinted 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NetAcquire - Master Buttons"
   ClientHeight    =   3525
   ClientLeft      =   2340
   ClientTop       =   3525
   ClientWidth     =   8085
   Icon            =   "frmUpdatePrinted.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   8085
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdRemoveEndo 
      Caption         =   "Remove All Endocrinology Requests"
      Height          =   525
      Left            =   4140
      TabIndex        =   9
      Top             =   2145
      Width           =   3735
   End
   Begin VB.CommandButton cmdSetEndo 
      Caption         =   "Set All Endocrinology Status to  Printed"
      Height          =   525
      Left            =   4140
      TabIndex        =   8
      Top             =   1485
      Width           =   3735
   End
   Begin VB.CommandButton bBGA 
      Caption         =   "Set All Blood Gas Status to  Printed"
      Height          =   525
      Left            =   4140
      TabIndex        =   7
      Top             =   2790
      Width           =   3735
   End
   Begin VB.CommandButton bImm 
      Caption         =   "Set All Immunology Status to  Printed"
      Height          =   525
      Left            =   4140
      TabIndex        =   6
      Top             =   195
      Width           =   3735
   End
   Begin VB.CommandButton bRemoveImmRequests 
      Caption         =   "Remove All Immunology Requests"
      Height          =   525
      Left            =   4140
      TabIndex        =   5
      Top             =   855
      Width           =   3735
   End
   Begin VB.CommandButton bRemoveCoagRequests 
      Caption         =   "Remove All Coagulation Requests"
      Height          =   525
      Left            =   120
      TabIndex        =   4
      Top             =   2820
      Width           =   3735
   End
   Begin VB.CommandButton bRemoveBioRequests 
      Caption         =   "Remove All Biochemistry Requests"
      Height          =   525
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   3735
   End
   Begin VB.CommandButton bCoag 
      Caption         =   "Set All Coagulation Status to  Printed"
      Height          =   525
      Left            =   120
      TabIndex        =   2
      Top             =   2190
      Width           =   3735
   End
   Begin VB.CommandButton bBio 
      Caption         =   "Set All Biochemistry Status to  Printed"
      Height          =   525
      Left            =   120
      TabIndex        =   1
      Top             =   180
      Width           =   3735
   End
   Begin VB.CommandButton bHaem 
      Caption         =   "Set All Haematology Status to  Printed"
      Height          =   525
      Left            =   120
      TabIndex        =   0
      Top             =   1530
      Width           =   3735
   End
End
Attribute VB_Name = "frmUpdatePrinted"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub bBGA_Click()

          Dim sql As String


10        On Error GoTo bBGA_Click_Error

20        sql = "UPDATE BgaResults " & _
                "Set Printed = 1 " & _
                "WHERE Printed = 0 or printed = '' or printed is null"
30        Cnxn(0).Execute sql




40        Exit Sub

bBGA_Click_Error:

          Dim strES As String
          Dim intEL As Integer



50        intEL = Erl
60        strES = Err.Description
70        LogError "frmUpdatePrinted", "bBGA_Click", intEL, strES, sql


End Sub

Private Sub bBio_Click()

          Dim sql As String


10        On Error GoTo bBio_Click_Error

20        sql = "UPDATE BioResults " & _
                "Set Printed = 1 " & _
                "WHERE Printed = 0 or printed = '' or printed is null"
30        Cnxn(0).Execute sql




40        Exit Sub

bBio_Click_Error:

          Dim strES As String
          Dim intEL As Integer



50        intEL = Erl
60        strES = Err.Description
70        LogError "frmUpdatePrinted", "bBio_Click", intEL, strES, sql


End Sub


Private Sub bCoag_Click()

          Dim sql As String


10        On Error GoTo bCoag_Click_Error

20        sql = "UPDATE CoagResults " & _
                "Set Printed = 1 " & _
                "WHERE COALESCE(Printed, 0) = 0"
30        Cnxn(0).Execute sql



40        Exit Sub

bCoag_Click_Error:

          Dim strES As String
          Dim intEL As Integer



50        intEL = Erl
60        strES = Err.Description
70        LogError "frmUpdatePrinted", "bCoag_Click", intEL, strES, sql


End Sub

Private Sub bHaem_Click()

          Dim sql As String


10        On Error GoTo bHaem_Click_Error

20        sql = "UPDATE HaemResults " & _
                "Set Printed = 1 " & _
                "WHERE Printed = 0 or printed = '' or printed is null"

30        Cnxn(0).Execute sql




40        Exit Sub

bHaem_Click_Error:

          Dim strES As String
          Dim intEL As Integer



50        intEL = Erl
60        strES = Err.Description
70        LogError "frmUpdatePrinted", "bHaem_Click", intEL, strES, sql


End Sub





Private Sub bImm_Click()

          Dim sql As String


10        On Error GoTo bImm_Click_Error

20        sql = "UPDATE ImmResults " & _
                "Set Printed = 1 " & _
                "WHERE Printed = 0 or printed = '' or printed is null"
30        Cnxn(0).Execute sql




40        Exit Sub

bImm_Click_Error:

          Dim strES As String
          Dim intEL As Integer



50        intEL = Erl
60        strES = Err.Description
70        LogError "frmUpdatePrinted", "bImm_Click", intEL, strES, sql


End Sub

Private Sub bRemoveBioRequests_Click()

          Dim sql As String


10        On Error GoTo bRemoveBioRequests_Click_Error

20        sql = "DELETE from BioRequests"

30        Cnxn(0).Execute sql



40        Exit Sub

bRemoveBioRequests_Click_Error:

          Dim strES As String
          Dim intEL As Integer



50        intEL = Erl
60        strES = Err.Description
70        LogError "frmUpdatePrinted", "bRemoveBioRequests_Click", intEL, strES, sql


End Sub


Private Sub bRemoveCoagRequests_Click()

          Dim sql As String


10        On Error GoTo bRemoveCoagRequests_Click_Error

20        sql = "DELETE from CoagRequests"

30        Cnxn(0).Execute sql



40        Exit Sub

bRemoveCoagRequests_Click_Error:

          Dim strES As String
          Dim intEL As Integer



50        intEL = Erl
60        strES = Err.Description
70        LogError "frmUpdatePrinted", "bRemoveCoagRequests_Click", intEL, strES, sql


End Sub


Private Sub bRemoveImmRequests_Click()

          Dim sql As String


10        On Error GoTo bRemoveImmRequests_Click_Error

20        sql = "DELETE from ImmRequests"

30        Cnxn(0).Execute sql


40        Exit Sub

bRemoveImmRequests_Click_Error:

          Dim strES As String
          Dim intEL As Integer



50        intEL = Erl
60        strES = Err.Description
70        LogError "frmUpdatePrinted", "bRemoveImmRequests_Click", intEL, strES, sql


End Sub


Private Sub cmdRemoveEndo_Click()
          Dim sql As String


10        On Error GoTo cmdRemoveEndo_Click_Error

20        sql = "DELETE from endRequests"

30        Cnxn(0).Execute sql


40        Exit Sub

cmdRemoveEndo_Click_Error:

          Dim strES As String
          Dim intEL As Integer



50        intEL = Erl
60        strES = Err.Description
70        LogError "frmUpdatePrinted", "cmdRemoveEndo_Click", intEL, strES, sql


End Sub

Private Sub cmdSetEndo_Click()
          Dim sql As String


10        On Error GoTo cmdSetEndo_Click_Error

20        sql = "UPDATE endResults " & _
                "Set Printed = 1 " & _
                "WHERE Printed = 0 or printed = '' or printed is null"
30        Cnxn(0).Execute sql




40        Exit Sub

cmdSetEndo_Click_Error:

          Dim strES As String
          Dim intEL As Integer



50        intEL = Erl
60        strES = Err.Description
70        LogError "frmUpdatePrinted", "cmdSetEndo_Click", intEL, strES, sql


End Sub
