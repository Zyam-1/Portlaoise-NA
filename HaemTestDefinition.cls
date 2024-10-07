VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HaemTestDefinition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public AnalyteName As String
Public MaleLow As Single
Public MaleHigh As Single
Public FemaleLow As Single
Public FemaleHigh As Single
Public FlagMaleLow As Single
Public FlagMaleHigh As Single
Public FlagFemaleLow As Single
Public FlagFemaleHigh As Single
Public LowControlMean As Single
Public NormalControlMean As Single
Public HighControlMean As Single
Public LowControl1SD As Single
Public NormalControl1SD As Single
Public HighControl1SD As Single

Public DoRM As Boolean
Public DoDelta As Boolean
Public DeltaValue As Single
Public Printformat As Long

Public PlausibleLow As Single
Public PlausibleHigh As Single
Public Category As String

Public AgeFromDays As Long
Public AgeToDays As Long

Public DP As Long
Public Units As String

Public InUse As Boolean
Public Hospital As String
Public Eod As Boolean
