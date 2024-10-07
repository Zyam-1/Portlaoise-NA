Attribute VB_Name = "modDbDesign"


Option Explicit

Public Sub CheckPatientNotepadInDb()

      Dim sql As String

10    On Error GoTo CheckPatientNotepadInDb_Error

20    If Not IsTableInDatabase("PatientNotePad") Then
30        sql = "CREATE TABLE [dbo].[PatientNotePad] " & _
                "([SampleID] [numeric](18, 0) NOT NULL, " & _
                "[DateTimeofRecord] [datetime] NOT NULL, " & _
                "[Comment] [nvarchar](4000) NOT NULL, " & _
                "[UserName] [nvarchar](20) NOT NULL, " & _
                "[Descipline] [nvarchar](20), " & _
                "[LabNo] [numeric](18, 0) )" & _
                "ON [PRIMARY]"
40        Cnxn(0).Execute sql
50    End If

60    Exit Sub

CheckPatientNotepadInDb_Error:

       Dim strES As String
       Dim intEL As Integer

70     intEL = Erl
80     strES = Err.Description
90     LogError "modDbDesign", "CheckPatientNotepadInDb", intEL, strES, sql
          
End Sub

'Public Sub CheckIQ200RepeatsInDb()
'
'      Dim sql As String
'
'10    On Error GoTo CheckIQ200RepeatsInDb_Error
'
'20    If IsTableInDatabase("IQ200Repeats") = False Then 'There is no table  in database
'30      sql = "CREATE TABLE IQ200Repeats " & _
 '              "( SampleID  numeric(18, 0) NOT NULL, " & _
 '              "  TestCode  nvarchar(50), " & _
 '              "  ShortName nvarchar(50), " & _
 '              "  LongName  nvarchar(50), " & _
 '              "  Range nvarchar(50), " & _
 '              "  Result nvarchar(50), " & _
 '              "  WorklistPrinted bit, " & _
 '              "  DateTimeOfRecord datetime NOT NULL DEFAULT getdate(), " & _
 '              "  Validated bit, " & _
 '              "  ValidatedBy nvarchar(50), " & _
 '              "  Printed bit, " & _
 '              "  PrintedBy nvarchar(50), " & _
 '              "  Counter numeric(18, 0) IDENTITY(1,1) NOT NULL )"
'40      Cnxn(0).Execute sql
'50    End If
'
'60    Exit Sub
'
'CheckIQ200RepeatsInDb_Error:
'
'Dim strES As String
'Dim intEL As Integer
'
'70    intEL = Erl
'80    strES = Err.Description
'90    LogError "modDbDesign", "CheckIQ200RepeatsInDb", intEL, strES, sql
'
'End Sub
'
'
'Public Sub CheckPhoresisRequestsInDb()
'
'      Dim sql As String
'
'10    On Error GoTo CheckPhoresisRequestsInDb_Error
'
'20    If IsTableInDatabase("PhoresisRequests") = False Then 'There is no table  in database
'30      sql = "CREATE TABLE [dbo].[PhoresisRequests] ( " & _
 '              "[AnalysisProgramCode] [char](1) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL, " & _
 '              "[PhoresisSampleNumber] [char](4) COLLATE SQL_Latin1_General_CP1_CI_AS NULL, " & _
 '              "[PatientID] [nvarchar](15) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL, " & _
 '              "[PatientName] [nvarchar](30) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL, " & _
 '              "[DoB] [smalldatetime] NULL, " & _
 '              "[Sex] [char](1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL, " & _
 '              "[Age] [char](3) COLLATE SQL_Latin1_General_CP1_CI_AS NULL, " & _
 '              "[Department] [nvarchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL, " & _
 '              "[SampleDate] [smalldatetime] NULL, " & _
 '              "[Concentration] [float] NULL, " & _
 '              "[DateTimeOfRecord] [datetime] NOT NULL CONSTRAINT [DF_PhoresisRequests_DateTimeOfRecord]  DEFAULT (getdate()), " & _
 '              "[Counter] [bigint] IDENTITY(1,1) NOT NULL, " & _
 '              "[UserName] [nvarchar](50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL, " & _
 '              "[Programmed] [tinyint] NOT NULL )"
'40      Cnxn(0).Execute sql
'50    End If
'
'60    Exit Sub
'
'CheckPhoresisRequestsInDb_Error:
'
'      Dim strES As String
'      Dim intEL As Integer
'
'70    intEL = Erl
'80    strES = Err.Description
'90    LogError "modDbDesign", "CheckPhoresisRequestsInDb", intEL, strES, sql
'
'End Sub
'
'
'
'
'
'Public Sub CheckPrintValidLogInDb()
'
'      Dim sql As String
'10    On Error GoTo CheckPrintValidLogInDb_Error
'
'20    If IsTableInDatabase("PrintValidLog") = False Then 'There is no table  in database
'30      sql = "CREATE TABLE PrintValidLog " & _
 '              "( SampleID  numeric(9), " & _
 '              "  Department nvarchar(1), " & _
 '              "  Printed tinyint, " & _
 '              "  Valid tinyint, " & _
 '              "  PrintedBy nvarchar(50), " & _
 '              "  PrintedDateTime datetime, " & _
 '              "  ValidatedBy nvarchar(50), " & _
 '              "  ValidatedDateTime datetime )"
'40      Cnxn(0).Execute sql
'50    End If
'
'60    Exit Sub
'
'CheckPrintValidLogInDb_Error:
'
'      Dim strES As String
'      Dim intEL As Integer
'
'70    intEL = Erl
'80    strES = Err.Description
'90    LogError "modDbDesign", "CheckPrintValidLogInDb", intEL, strES, sql
'
'
'End Sub
'
'Public Sub CheckMicroExternalsInDb()
'
'      Dim sql As String
'
'10    On Error GoTo CheckMicroExternalsInDb_Error
'
'20    If IsTableInDatabase("MicroExternals") = False Then 'There is no table  in database
'30      sql = "CREATE TABLE MicroExternals " & _
 '              "( MicroSID  numeric(9) NOT NULL, " & _
 '              "  InHouseSID  numeric(9) NOT NULL, " & _
 '              "  OrderGlu bit NOT NULL, " & _
 '              "  OrderTP bit NOT NULL, " & _
 '              "  OrderAlb bit NOT NULL, " & _
 '              "  OrderGlo bit NOT NULL, " & _
 '              "  OrderLDH bit NOT NULL, " & _
 '              "  OrderAmy bit NOT NULL, " & _
 '              "  DateTimeOfRecord datetime NOT NULL DEFAULT getdate(), " & _
 '              "  UserName nvarchar(50) NOT NULL, " & _
 '              "  RowGUID uniqueidentifier NOT NULL DEFAULT newid() )"
'40      Cnxn(0).Execute sql
'50    End If
'
'60    Exit Sub
'
'CheckMicroExternalsInDb_Error:
'
'      Dim strES As String
'      Dim intEL As Integer
'
'70    intEL = Erl
'80    strES = Err.Description
'90    LogError "modDbDesign", "CheckMicroExternalsInDb", intEL, strES, sql
'
'
'End Sub
'
'Public Sub CheckMicroExternalResultsInDb()
'
'      Dim sql As String
'
'10    On Error GoTo CheckMicroExternalResultsInDb_Error
'
'20    If IsTableInDatabase("MicroExternalResults") = False Then 'There is no table  in database
'30      sql = "CREATE TABLE MicroExternalResults " & _
 '              "( SampleID  numeric(9) NOT NULL, " & _
 '              "  TestName  nvarchar(50) NOT NULL, " & _
 '              "  SentTo nvarchar(50), " & _
 '              "  SentDate datetime, " & _
 '              "  InterimReportDate datetime, " & _
 '              "  InterimReportComment nvarchar(50), " & _
 '              "  FinalReportDate datetime, " & _
 '              "  FinalReportComment nvarchar(50), " & _
 '              "  DateTimeOfRecord datetime NOT NULL DEFAULT getdate(), " & _
 '              "  UserName nvarchar(50) NOT NULL, " & _
 '              "  RowGUID uniqueidentifier NOT NULL DEFAULT newid() )"
'40      Cnxn(0).Execute sql
'50    End If
'
'60    Exit Sub
'
'CheckMicroExternalResultsInDb_Error:
'
'      Dim strES As String
'      Dim intEL As Integer
'
'70    intEL = Erl
'80    strES = Err.Description
'90    LogError "modDbDesign", "CheckMicroExternalResultsInDb", intEL, strES, sql
'
'End Sub
'
'Public Sub CheckMicroExternalResultsArcInDb()
'
'      Dim sql As String
'
'10    On Error GoTo CheckMicroExternalResultsArcInDb_Error
'
'20    If IsTableInDatabase("MicroExternalResultsArc") = False Then 'There is no table  in database
'30      sql = "CREATE TABLE MicroExternalResultsArc " & _
 '              "( SampleID  numeric(9) NOT NULL, " & _
 '              "  TestName  nvarchar(50) NOT NULL, " & _
 '              "  SentTo nvarchar(50), " & _
 '              "  SentDate datetime, " & _
 '              "  InterimReportDate datetime, " & _
 '              "  InterimReportComment nvarchar(50), " & _
 '              "  FinalReportDate datetime, " & _
 '              "  FinalReportComment nvarchar(50), " & _
 '              "  DateTimeOfRecord datetime, " & _
 '              "  UserName nvarchar(50) NOT NULL, " & _
 '              "  ArchiveDateTime datetime NOT NULL DEFAULT getdate(), " & _
 '              "  ArchivedBy nvarchar(50) , " & _
 '              "  RowGUID uniqueidentifier NOT NULL DEFAULT newid() )"
'40      Cnxn(0).Execute sql
'50    End If
'
'60    Exit Sub
'
'CheckMicroExternalResultsArcInDb_Error:
'
'      Dim strES As String
'      Dim intEL As Integer
'
'70    intEL = Erl
'80    strES = Err.Description
'90    LogError "modDbDesign", "CheckMicroExternalResultsArcInDb", intEL, strES, sql
'
'End Sub
'Public Sub CheckUrineRequestsInDb()
'
'      Dim sql As String
'
'10    On Error GoTo CheckUrineRequestsInDb_Error
'
'20    If IsTableInDatabase("UrineRequests") = False Then 'There is no table  in database
'30      sql = "CREATE TABLE UrineRequests " & _
 '              "( SampleID  numeric(9) NOT NULL, " & _
 '              "  CS bit, " & _
 '              "  Pregnancy bit, " & _
 '              "  RedSub bit, " & _
 '              "  DoNotDisplayInBatchEntry bit, " & _
 '              "  DateTimeOfRecord datetime NOT NULL DEFAULT getdate() )"
'40      Cnxn(0).Execute sql
'50    End If
'
'60    Exit Sub
'
'CheckUrineRequestsInDb_Error:
'
'      Dim strES As String
'      Dim intEL As Integer
'
'70    intEL = Erl
'80    strES = Err.Description
'90    LogError "modDbDesign", "CheckUrineRequestsInDb", intEL, strES, sql
'
'End Sub
'
'Public Sub CheckUrineRequestsArcInDb()
'
'      Dim sql As String
'
'10    On Error GoTo CheckUrineRequestsArcInDb_Error
'
'20    If IsTableInDatabase("UrineRequestsArc") = False Then 'There is no table  in database
'30      sql = "CREATE TABLE UrineRequestsArc " & _
 '              "( SampleID  numeric(9) NOT NULL, " & _
 '              "  CS bit, " & _
 '              "  Pregnancy bit, " & _
 '              "  RedSub bit, " & _
 '              "  DoNotDisplayInBatchEntry bit, " & _
 '              "  DateTimeOfRecord datetime NOT NULL DEFAULT getdate(), " & _
 '              "  UserName nvarchar(50), " & _
 '              "  ArchivedBy nvarchar(50) NOT NULL, " & _
 '              "  ArchiveDateTime datetime NOT NULL DEFAULT getdate(), " & _
 '              "  RowGUID uniqueidentifier NOT NULL DEFAULT newid() )"
'40      Cnxn(0).Execute sql
'50    End If
'
'60    Exit Sub
'
'CheckUrineRequestsArcInDb_Error:
'
'      Dim strES As String
'      Dim intEL As Integer
'
'70    intEL = Erl
'80    strES = Err.Description
'90    LogError "modDbDesign", "CheckUrineRequestsArcInDb", intEL, strES, sql
'
'End Sub
'
'Public Sub CheckLoggedOnUsersInDb()
'
'      Dim sql As String
'
'10    On Error GoTo CheckLoggedOnUsersInDb_Error
'
'20    If IsTableInDatabase("LoggedOnUsers") = False Then 'There is no table  in database
'30      sql = "CREATE TABLE dbo.[LoggedOnUsers](" & _
 '              "  [MachineName] [nvarchar](50) NULL, " & _
 '              "  [UserName] [nvarchar](50) NULL, " & _
 '              "  [AppName] [nvarchar](50) NULL)"
'40      Cnxn(0).Execute sql
'50    End If
'
'60    Exit Sub
'
'CheckLoggedOnUsersInDb_Error:
'
'Dim strES As String
'Dim intEL As Integer
'
'70    intEL = Erl
'80    strES = Err.Description
'90    LogError "modDbDesign", "CheckLoggedOnUsersInDb", intEL, strES, sql
'
'End Sub
'
'Public Sub CheckMicroExtLabNameInDb()
'
'      Dim sql As String
'10    On Error GoTo CheckMicroExtLabNameInDb_Error
'
'20    If IsTableInDatabase("MicroExtLabName") = False Then 'There is no table  in database
'30      sql = "CREATE TABLE MicroExtLabName " & _
 '              "( LabName nvarchar(50), " & _
 '              "  Address0 nvarchar(50), " & _
 '              "  Address1 nvarchar(50), " & _
 '              "  Address2 nvarchar(50), " & _
 '              "  DateTimeOfRecord datetime DEFAULT getdate(), " & _
 '              "  RowGUID uniqueidentifier DEFAULT newid() )"
'40      Cnxn(0).Execute sql
'50    End If
'
'60    Exit Sub
'
'CheckMicroExtLabNameInDb_Error:
'
'      Dim strES As String
'      Dim intEL As Integer
'
'70    intEL = Erl
'80    strES = Err.Description
'90    LogError "modDbDesign", "CheckMicroExtLabNameInDb", intEL, strES, sql
'
'
'End Sub
'
'
'Public Sub CheckIsolatesRepeatsInDb()
'
'      Dim sql As String
'
'10    On Error GoTo CheckIsolatesRepeatsInDb_Error
'
'20    If IsTableInDatabase("IsolatesRepeats") = False Then 'There is no table  in database
'30      sql = "CREATE TABLE IsolatesRepeats " & _
 '              "( SampleID  numeric(9), " & _
 '              "  IsolateNumber int, " & _
 '              "  OrganismGroup nvarchar(50), " & _
 '              "  OrganismName nvarchar(50), " & _
 '              "  Qualifier nvarchar(50) )"
'40      Cnxn(0).Execute sql
'50    End If
'
'60    Exit Sub
'
'CheckIsolatesRepeatsInDb_Error:
'
'      Dim strES As String
'      Dim intEL As Integer
'
'70    intEL = Erl
'80    strES = Err.Description
'90    LogError "modDbDesign", "CheckIsolatesRepeatsInDb", intEL, strES, sql
'
'
'End Sub
'
'Public Sub CheckLockStatusInDb()
'
'      Dim sql As String
'10    On Error GoTo CheckLockStatusInDb_Error
'
'20    If IsTableInDatabase("LockStatus") = False Then 'There is no table  in database
'30      sql = "CREATE TABLE LockStatus " & _
 '              "( SampleID  numeric(9), " & _
 '              "  Lock bit, " & _
 '              "  DeptIndex int, " & _
 '              "  RowGUID uniqueidentifier default newid() )"
'40      Cnxn(0).Execute sql
'50    End If
'
'60    Exit Sub
'
'CheckLockStatusInDb_Error:
'
'      Dim strES As String
'      Dim intEL As Integer
'
'70    intEL = Erl
'80    strES = Err.Description
'90    LogError "modDbDesign", "CheckLockStatusInDb", intEL, strES, sql
'
'
'End Sub
'Public Sub CheckSensitivitiesRepeatsInDb()
'
'      Dim sql As String
'10    On Error GoTo CheckSensitivitiesRepeatsInDb_Error
'
'20    If IsTableInDatabase("SensitivitiesRepeats") = False Then 'There is no table  in database
'30      sql = "CREATE TABLE SensitivitiesRepeats " & _
 '              "( SampleID  numeric(9), " & _
 '              "  IsolateNumber int, " & _
 '              "  AntibioticCode nvarchar(50), " & _
 '              "  Result nvarchar(50), " & _
 '              "  Report bit, " & _
 '              "  CPOFlag nvarchar(1), " & _
 '              "  RunDate datetime, " & _
 '              "  RunDateTime datetime, " & _
 '              "  RSI char(1), " & _
 '              "  UserName nvarchar(50), " & _
 '              "  Forced bit, " & _
 '              "  Secondary bit, " & _
 '              "  Valid bit, " & _
 '              "  AuthoriserCode nvarchar(50) )"
'40      Cnxn(0).Execute sql
'50    End If
'
'60    Exit Sub
'
'CheckSensitivitiesRepeatsInDb_Error:
'
'      Dim strES As String
'      Dim intEL As Integer
'
'70    intEL = Erl
'80    strES = Err.Description
'90    LogError "modDbDesign", "CheckSensitivitiesRepeatsInDb", intEL, strES, sql
'
'
'End Sub
'
'Public Sub CheckFaecesResultsInDb()
'
'      Dim sql As String
'
'10    On Error GoTo CheckFaecesResultsInDb_Error
'
'20    If IsTableInDatabase("FaecesResults") = False Then 'There is no table  in database
'30      sql = "CREATE TABLE FaecesResults " & _
 '              "( SampleID  numeric NOT NULL, " & _
 '              "  TestName nvarchar(50) NOT NULL, " & _
 '              "  Result nvarchar(50) NOT NULL, " & _
 '              "  UserName nvarchar(50) NOT NULL, " & _
 '              "  Valid bit NOT NULL, " & _
 '              "  HealthLink tinyint NOT NULL, " & _
 '              "  DateTimeOfRecord datetime NOT NULL )"
'40      Cnxn(0).Execute sql
'
'50    End If
'
'60    Exit Sub
'
'CheckFaecesResultsInDb_Error:
'
'Dim strES As String
'Dim intEL As Integer
'
'70    intEL = Erl
'80    strES = Err.Description
'90    LogError "modDbDesign", "CheckFaecesResultsInDb", intEL, strES, sql
'
'End Sub
'
'Public Sub CheckGenericResultsInDb()
'
'      Dim sql As String
'
'10    On Error GoTo CheckGenericResultsInDb_Error
'
'20    If IsTableInDatabase("GenericResults") = False Then 'There is no table  in database
'30      sql = "CREATE TABLE GenericResults " & _
 '              "( SampleID  numeric(9), " & _
 '              "  TestName nvarchar(50), " & _
 '              "  Result nvarchar(50) )"
'40      Cnxn(0).Execute sql
'50    End If
'
'60    Exit Sub
'
'CheckGenericResultsInDb_Error:
'
'      Dim strES As String
'      Dim intEL As Integer
'
'70    intEL = Erl
'80    strES = Err.Description
'90    LogError "modDbDesign", "CheckGenericResultsInDb", intEL, strES, sql
'
'End Sub
'
'
'Public Sub CheckFaecesWorksheetInDb()
'
'      Dim sql As String
'
'10    On Error GoTo CheckFaecesWorksheetInDb_Error
'
'20    If IsTableInDatabase("FaecesWorksheet") = False Then 'There is no table  in database
'30      sql = "CREATE TABLE FaecesWorksheet " & _
 '              "( SampleID  numeric(9), " & _
 '              "  Day111 nvarchar(50), Day112 nvarchar(50), Day113 nvarchar(50), " & _
 '              "  Day121 nvarchar(50), Day122 nvarchar(50), Day123 nvarchar(50), " & _
 '              "  Day131 nvarchar(50), Day132 nvarchar(50), Day133 nvarchar(50), " & _
 '              "  Day211 nvarchar(50), Day212 nvarchar(50), Day213 nvarchar(50), " & _
 '              "  Day221 nvarchar(50), Day222 nvarchar(50), Day223 nvarchar(50), " & _
 '              "  Day231 nvarchar(50), Day232 nvarchar(50), Day233 nvarchar(50), " & _
 '              "  Day31 nvarchar(50), Day32 nvarchar(50), Day33 nvarchar(50), " & _
 '              "  TimeOfRecord datetime, " & _
 '              "  Operator nvarchar(50) )"
'40      Cnxn(0).Execute sql
'50    End If
'
'60    Exit Sub
'
'CheckFaecesWorksheetInDb_Error:
'
'      Dim strES As String
'      Dim intEL As Integer
'
'70    intEL = Erl
'80    strES = Err.Description
'90    LogError "modDbDesign", "CheckFaecesWorksheetInDb", intEL, strES, sql
'
'
'End Sub
'
'Public Sub CheckGenericResultsArcInDb()
'
'      Dim sql As String
'
'10    On Error GoTo CheckGenericResultsArcInDb_Error
'
'20    If IsTableInDatabase("GenericResultsArc") = False Then 'There is no table  in database
'30      sql = "CREATE TABLE GenericResultsArc " & _
 '              "( SampleID numeric NOT NULL, " & _
 '              "  TestName nvarchar(50), " & _
 '              "  Result nvarchar(50), " & _
 '              "  UserName nvarchar(50), " & _
 '              "  ArchivedBy nvarchar(50), " & _
 '              "  ArchiveDateTime datetime default getdate(), " & _
 '              "  RowGUID uniqueidentifier default newid() )"
'40      Cnxn(0).Execute sql
'50    End If
'
'60    Exit Sub
'
'CheckGenericResultsArcInDb_Error:
'
'      Dim strES As String
'      Dim intEL As Integer
'
'70    intEL = Erl
'80    strES = Err.Description
'90    LogError "modDbDesign", "CheckGenericResultsArcInDb", intEL, strES, sql
'
'End Sub
'
'Public Sub CheckMicroSiteDetailsArcInDb()
'
'      Dim sql As String
'
'10    On Error GoTo CheckMicroSiteDetailsArcInDb_Error
'
'20    If IsTableInDatabase("MicroSiteDetailsArc") = False Then 'There is no table  in database
'30      sql = "CREATE TABLE MicroSiteDetailsArc " & _
 '              "( SampleID numeric NOT NULL, " & _
 '              "  Site nvarchar(50), " & _
 '              "  SiteDetails nvarchar(50), " & _
 '              "  PCA0 nvarchar(50), " & _
 '              "  PCA1 nvarchar(50), " & _
 '              "  PCA2 nvarchar(50), " & _
 '              "  PCA3 nvarchar(50), " & _
 '              "  ArchiveDateTime datetime default getdate(), " & _
 '              "  ArchivedBy nvarchar(50), " & _
 '              "  rowguid uniqueidentifier default newid() )"
'40      Cnxn(0).Execute sql
'50    End If
'
'60    Exit Sub
'
'CheckMicroSiteDetailsArcInDb_Error:
'
'      Dim strES As String
'      Dim intEL As Integer
'
'70    intEL = Erl
'80    strES = Err.Description
'90    LogError "modDbDesign", "CheckMicroSiteDetailsArcInDb", intEL, strES, sql
'
'End Sub
'
'Public Sub CheckSemenResultsArcInDb()
'
'      Dim sql As String
'
'10    On Error GoTo CheckSemenResultsArcInDb_Error
'
'20    If IsTableInDatabase("SemenResultsArc") = False Then 'There is no table  in database
'30      sql = "CREATE TABLE SemenResultsArc " & _
 '              "( SampleID numeric NOT NULL, " & _
 '              "  Volume nvarchar(50), " & _
 '              "  SemenCount nvarchar(50), " & _
 '              "  MotilityPro nvarchar(50), " & _
 '              "  MotilityNonPro nvarchar(50), " & _
 '              "  MotilityNonMotile nvarchar(50), " & _
 '              "  Consistency nvarchar(50), " & _
 '              "  Valid int, " & _
 '              "  Operator nvarchar(50), " & _
 '              "  Printed int, " & _
 '              "  Motility nvarchar(50), " & _
 '              "  ArchiveDateTime datetime default getdate(), " & _
 '              "  ArchivedBy nvarchar(50), " & _
 '              "  rowguid uniqueidentifier default newid() )"
'40      Cnxn(0).Execute sql
'50    End If
'
'60    Exit Sub
'
'CheckSemenResultsArcInDb_Error:
'
'      Dim strES As String
'      Dim intEL As Integer
'
'70    intEL = Erl
'80    strES = Err.Description
'90    LogError "modDbDesign", "CheckSemenResultsArcInDb", intEL, strES, sql
'
'End Sub
'
'Public Sub CheckUrineArcInDb()
'
'      Dim sql As String
'
'10    On Error GoTo CheckUrineArcInDb_Error
'
'20    If IsTableInDatabase("UrineArc") = False Then 'There is no table  in database
'30      sql = "CREATE TABLE UrineArc " & _
 '              "( SampleID numeric NOT NULL, " & _
 '              "  Pregnancy nvarchar(50), " & _
 '              "  HCGLevel nvarchar(50), " & _
 '              "  BenceJones nvarchar(50), " & _
 '              "  SG nvarchar(50), " & _
 '              "  FatGlobules nvarchar(50), " & _
 '              "  pH nvarchar(50), " & _
 '              "  Protein nvarchar(50), " & _
 '              "  Glucose nvarchar(50), " & _
 '              "  Ketones nvarchar(50), " & _
 '              "  Urobilinogen nvarchar(50), " & _
 '              "  Bilirubin nvarchar(50), " & _
 '              "  BloodHb nvarchar(50), " & _
 '              "  WCC nvarchar(50), " & _
 '              "  RCC nvarchar(50), " & _
 '              "  Crystals nvarchar(50), " & _
 '              "  Casts nvarchar(50), " & _
 '              "  Misc0 nvarchar(50), " & _
 '              "  Misc1 nvarchar(50), " & _
 '              "  Misc2 nvarchar(50), " & _
 '              "  Valid bit, " & _
 '              "  Bacteria nvarchar(50), " & _
 '              "  [Count] nvarchar(50), " & _
 '              "  HealthLink tinyint, "
'
'40      sql = sql & "Printed int, " & _
 '              "  ArchiveDateTime datetime default getdate(), " & _
 '              "  ArchivedBy nvarchar(50), " & _
 '              "  rowguid uniqueidentifier default newid(), " & _
 '              "  UserName nvarchar(50) )"
'50      Cnxn(0).Execute sql
'60    End If
'
'70    Exit Sub
'
'CheckUrineArcInDb_Error:
'
'      Dim strES As String
'      Dim intEL As Integer
'
'80    intEL = Erl
'90    strES = Err.Description
'100   LogError "modDbDesign", "CheckUrineArcInDb", intEL, strES, sql
'
'End Sub
'
'Public Sub CheckUrineIdentArcInDb()
'
'      Dim sql As String
'
'10    On Error GoTo CheckUrineIdentArcInDb_Error
'
'20    If IsTableInDatabase("UrineIdentArc") = False Then 'There is no table  in database
'30      sql = "CREATE TABLE UrineIdentArc " & _
 '              "( SampleID numeric NOT NULL, " & _
 '              "  Gram nvarchar(50), " & _
 '              "  WetPrep nvarchar(50), " & _
 '              "  Coagulase nvarchar(50), " & _
 '              "  Catalase nvarchar(50), " & _
 '              "  Oxidase nvarchar(50), " & _
 '              "  API0 nvarchar(50), " & _
 '              "  API1 nvarchar(50), " & _
 '              "  Ident0 nvarchar(50), " & _
 '              "  Ident1 nvarchar(50), " & _
 '              "  Rapidec nvarchar(50), " & _
 '              "  Chromogenic nvarchar(50), " & _
 '              "  Reincubation nvarchar(50), " & _
 '              "  UrineSensitivity nvarchar(50), " & _
 '              "  ExtraSensitivity nvarchar(50), " & _
 '              "  Valid bit, " & _
 '              "  Isolate int, " & _
 '              "  Notes nvarchar(500), " & _
 '              "  UserName nvarchar(50), " & _
 '              "  ArchiveDateTime datetime default getdate(), " & _
 '              "  ArchivedBy nvarchar(50), " & _
 '              "  RowGUID uniqueidentifier default newid() )"
'40      Cnxn(0).Execute sql
'50    End If
'
'60    Exit Sub
'
'CheckUrineIdentArcInDb_Error:
'
'      Dim strES As String
'      Dim intEL As Integer
'
'70    intEL = Erl
'80    strES = Err.Description
'90    LogError "modDbDesign", "CheckUrineIdentArcInDb", intEL, strES, sql
'
'End Sub
Public Function IsTableInDatabase(ByVal TableName As String) As Boolean

          Dim tbExists As Recordset
          Dim sql As String
          Dim RetVal As Boolean

          'How to find if a table exists in a database
          'open a recordset with the following sql statement:
          'Code:SELECT name FROM sysobjects WHERE xtype = 'U' AND name = 'MyTable'
          'If the recordset it at eof then the table doesn't exist
          'if it has a record then the table does exist.

10        On Error GoTo IsTableInDatabase_Error

20        sql = "SELECT OBJECT_ID('dbo." & TableName & "', 'U') E"
          '20    sql = "SELECT name FROM sysobjects WHERE " & _
           '            "xtype = 'U' " & _
           '            "AND name = 'dbo." & TableName & "'"
30        Set tbExists = New Recordset
40        Set tbExists = Cnxn(0).Execute(sql)

50        RetVal = True

60        If tbExists.EOF Then    'There is no table <TableName> in database
70            RetVal = False
80        Else
90            If IsNull(tbExists!e) Then
100               RetVal = False
110           Else
120               RetVal = True
130           End If
140       End If
150       IsTableInDatabase = RetVal

160       Exit Function

IsTableInDatabase_Error:

          Dim strES As String
          Dim intEL As Integer

170       intEL = Erl
180       strES = Err.Description
190       LogError "modDbDesign", "IsTableInDatabase", intEL, strES, sql

End Function

Public Function EnsureColumnExists(ByVal TableName As String, _
                                   ByVal ColumnName As String, _
                                   ByVal Definition As String) _
                                   As Boolean

      'Return 1 if column created
      '       0 if column already exists

          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo EnsureColumnExists_Error

20        sql = "IF NOT EXISTS " & _
                "    (SELECT * FROM syscolumns WHERE " & _
                "    id = object_id('" & TableName & "') " & _
                "    AND name = '" & ColumnName & "') " & _
                "  BEGIN " & _
                "    ALTER TABLE " & TableName & " " & _
                "    ADD " & ColumnName & " " & Definition & " " & _
                "    SELECT 1 AS RetVal " & _
                "  END " & _
                "ELSE " & _
                "  SELECT 0 AS RetVal"

30        Set tb = Cnxn(0).Execute(sql)

40        EnsureColumnExists = tb!RetVal

50        Exit Function

EnsureColumnExists_Error:

          Dim strES As String
          Dim intEL As Integer

60        intEL = Erl
70        strES = Err.Description
80        LogError "modDbDesign", "EnsureColumnExists", intEL, strES, sql

End Function

Public Sub CheckAutoCommentsInDb()

          Dim sql As String

10        On Error GoTo CheckAutoCommentsInDb_Error

20        If IsTableInDatabase("AutoComments") = False Then    'There is no table  in database
30            sql = "CREATE TABLE AutoComments " & _
                    "( Discipline nvarchar(50) NOT NULL, " & _
                    "  Parameter nvarchar(50) NOT NULL, " & _
                    "  Criteria nvarchar(50) NOT NULL, " & _
                    "  Value0 nvarchar(50), " & _
                    "  Value1 nvarchar(50), " & _
                    "  Comment nvarchar(80), " & _
                    "  DateStart smalldatetime, " & _
                    "  DateEnd smalldatetime, " & _
                    "  ListOrder tinyint, " & _
                    "  DateTimeOfRecord datetime NOT NULL DEFAULT getdate() )"
40            Cnxn(0).Execute sql
50        End If

60        Exit Sub

CheckAutoCommentsInDb_Error:

          Dim strES As String
          Dim intEL As Integer

70        intEL = Erl
80        strES = Err.Description
90        LogError "modDbDesign", "CheckAutoCommentsInDb", intEL, strES, sql

End Sub

Public Function EnsureIndexExists(ByVal TableName As String, _
                                  ByVal ColumnName As String, _
                                  ByVal IndexName As String) _
                                  As Boolean

          Dim sql As String
          Dim tb As Recordset

10        On Error GoTo EnsureIndexExists_Error

20        sql = "IF NOT EXISTS " & _
                "    (SELECT name FROM sysindexes WHERE " & _
                "    name = '" & IndexName & "')" & _
                "  BEGIN " & _
                "    CREATE index [" & IndexName & "] " & _
                "    ON [" & TableName & "] ([" & ColumnName & "]) " & _
                "    SELECT 1 AS RetVal " & _
                "  END " & _
                "ELSE " & _
                "  SELECT 0 AS RetVal"

30        Set tb = Cnxn(0).Execute(sql)

40        EnsureIndexExists = tb!RetVal

50        Exit Function

EnsureIndexExists_Error:

          Dim strES As String
          Dim intEL As Integer

60        intEL = Erl
70        strES = Err.Description
80        LogError "modDbDesign", "EnsureIndexExists", intEL, strES, sql

End Function

Public Function EnsureOptionExists(ByVal Description As String, _
                                   ByVal Contents As String) _
                                   As Boolean

      'Return 1 if column created
      '       0 if column already exists

          Dim sql As String
          Dim tb As Recordset



10        On Error GoTo EnsureOptionExists_Error

20        sql = "IF NOT EXISTS " & _
                "    (SELECT * FROM Options WHERE " & _
                "    Description = '" & Description & "') " & _
                "  BEGIN " & _
                "    INSERT INTO Options (Description, Contents, UserName) " & _
                "    VALUES ('" & Description & "', '" & Contents & "', 'System') " & _
                "  END " & _
                "ELSE " & _
                "  SELECT 0 AS RetVal"

30        Set tb = Cnxn(0).Execute(sql)

40        EnsureOptionExists = True




50        Exit Function

EnsureOptionExists_Error:

          Dim strES As String
          Dim intEL As Integer

60        intEL = Erl
70        strES = Err.Description
80        LogError "modDbDesign", "EnsureOptionExists", intEL, strES, sql


End Function

Public Sub CheckAnalyserTestCodeMappingInDb()

          Dim sql As String

10        On Error GoTo CheckAnalyserTestCodeMappingInDb_Error

20        If IsTableInDatabase("AnalyserTestCodeMapping") = False Then

30            sql = "CREATE TABLE [dbo].[AnalyserTestCodeMapping]( " & _
                    "[NetAcquireTestCode] [nvarchar](50) NULL, " & _
                    "[EquipmentAnalyserCode] [nvarchar](50) NOT NULL, " & _
                    "[TestName] [nvarchar](100) NOT NULL, " & _
                    "[AnalyserName] [nvarchar](50) NOT NULL, " & _
                    "[DateTimeOfRecord] [datetime] NULL, " & _
                    "[Counter] [numeric](18, 0) IDENTITY(1,1) NOT NULL, " & _
                    "[Department] [nvarchar](50) NULL " & _
                    ") ON [PRIMARY]"

40            Cnxn(0).Execute sql
50        End If
60        Exit Sub

CheckAnalyserTestCodeMappingInDb_Error:

          Dim strES As String
          Dim intEL As Integer

70        intEL = Erl
80        strES = Err.Description
90        LogError "modDbDesign", "CheckAnalyserTestCodeMappingInDb", intEL, strES, sql

End Sub

Public Sub CheckPrintingRulesInDb()

          Dim sql As String

10        On Error GoTo CheckPrintingRulesInDb_Error

20        If IsTableInDatabase("PrintingRules") = False Then
30            sql = "CREATE TABLE [dbo].[PrintingRules]( " & _
                    "[TestName] [nvarchar](200) NULL, " & _
                    "[Criteria] [nvarchar](200) NULL, " & _
                    "[Type] [nvarchar](50) NULL, " & _
                    "[Bold] [bit] NULL, " & _
                    "[Italic] [bit] NULL, " & _
                    "[Underline] [bit] NULL, " & _
                    "[Counter] [int] IDENTITY(1,1) NOT NULL " & _
                    ") ON [PRIMARY] "

40            Cnxn(0).Execute sql
50        End If
60        Exit Sub

CheckPrintingRulesInDb_Error:

          Dim strES As String
          Dim intEL As Integer

70        intEL = Erl
80        strES = Err.Description
90        LogError "modDbDesign", "CheckPrintingRulesInDb", intEL, strES, sql

End Sub

'---------------------------------------------------------------------------------------
' Procedure : CheckUnauthorisedReportsInDb
' Author    : XPMUser
' Date      : 2/19/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub CheckUnauthorisedReportsInDb()    'Masood 19_Feb_2013

          Dim sql As String
10        On Error GoTo CheckUnauthorisedReportsInDb_Error


20        If IsTableInDatabase("UnauthorisedReports") = False Then
30            sql = " CREATE TABLE [dbo].[UnauthorisedReports]( " & _
                    " [Sampleid] [decimal](18, 0) NULL, " & _
                    " [Name] [nvarchar](50) NULL, " & _
                    " [Dept] [char](2) NULL, " & _
                    " [Initiator] [nvarchar](50) NULL, " & _
                    " [PrintTime] [datetime] NULL, " & _
                    " [RepNo] [char](30) NULL, " & _
                    " [PageOne] Text NULL, " & _
                    " [PageTwo] Text NULL, " & _
                    " [Printer] [nvarchar](100) NULL, " & _
                    " [Printed] [tinyint] NULL, " & _
                    " [PageThree] [text] NULL, " & _
                    " [PageFour] [text] NULL, " & _
                    " [PageNumber] [tinyint] NULL, " & _
                    " [Report] Text NULL, " & _
                    " [Notes] [nvarchar](1000) NULL, " & _
                    " [Year] [smallint] NULL " & _
                    " ) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]"

40            Cnxn(0).Execute sql
50        End If


60        Exit Sub


CheckUnauthorisedReportsInDb_Error:

          Dim strES As String
          Dim intEL As Integer

70        intEL = Erl
80        strES = Err.Description
90        LogError "modDbDesign", "CheckUnauthorisedReportsInDb", intEL, strES, sql
End Sub

Public Sub CheckImmDefIndexInDb()

10    On Error GoTo CheckImmDefIndexInDb_Error

      Dim sql As String
20    If IsTableInDatabase("ImmDefIndex") = False Then
30        sql = "CREATE TABLE [dbo].[ImmDefIndex]( " & _
                "[DefIndex] [numeric](18, 0) IDENTITY(1,1) NOT NULL, " & _
                "[NormalLow] [real] NOT NULL, " & _
                "[NormalHigh] [real] NOT NULL, " & _
                "[FlagLow] [real] NOT NULL, " & _
                "[FlagHigh] [real] NOT NULL, " & _
                "[PlausibleLow] [real] NOT NULL, " & _
                "[PlausibleHigh] [real] NOT NULL, " & _
                "[AutoValLow] [real] NULL, " & _
                "[AutoValHigh] [Real] NULL " & _
                ") ON [PRIMARY]"
40        Cnxn(0).Execute sql
50    End If


60    Exit Sub

CheckImmDefIndexInDb_Error:

       Dim strES As String
       Dim intEL As Integer

70     intEL = Erl
80     strES = Err.Description
90     LogError "modDbDesign", "CheckImmDefIndexInDb", intEL, strES, sql
          
End Sub

Public Sub CheckEndDefIndexInDb()

10    On Error GoTo CheckEndDefIndexInDb_Error

      Dim sql As String
20    If IsTableInDatabase("EndDefIndex") = False Then
30        sql = "CREATE TABLE [dbo].[EndDefIndex]( " & _
                "[DefIndex] [numeric](18, 0) IDENTITY(1,1) NOT NULL, " & _
                "[NormalLow] [real] NOT NULL, " & _
                "[NormalHigh] [real] NOT NULL, " & _
                "[FlagLow] [real] NOT NULL, " & _
                "[FlagHigh] [real] NOT NULL, " & _
                "[PlausibleLow] [real] NOT NULL, " & _
                "[PlausibleHigh] [real] NOT NULL, " & _
                "[AutoValLow] [real] NULL, " & _
                "[AutoValHigh] [Real] NULL " & _
                ") ON [PRIMARY]"
40        Cnxn(0).Execute sql
50    End If


60    Exit Sub

CheckEndDefIndexInDb_Error:

       Dim strES As String
       Dim intEL As Integer

70     intEL = Erl
80     strES = Err.Description
90     LogError "modDbDesign", "CheckEndDefIndexInDb", intEL, strES, sql
          
End Sub

Public Sub CheckBioDefIndexInDb()

10    On Error GoTo CheckBioDefIndexInDb_Error

      Dim sql As String
20    If IsTableInDatabase("BioDefIndex") = False Then
30        sql = "CREATE TABLE [dbo].[BioDefIndex]( " & _
                "[DefIndex] [numeric](18, 0) IDENTITY(1,1) NOT NULL, " & _
                "[NormalLow] [real] NOT NULL, " & _
                "[NormalHigh] [real] NOT NULL, " & _
                "[FlagLow] [real] NOT NULL, " & _
                "[FlagHigh] [real] NOT NULL, " & _
                "[PlausibleLow] [real] NOT NULL, " & _
                "[PlausibleHigh] [real] NOT NULL, " & _
                "[AutoValLow] [real] NULL, " & _
                "[AutoValHigh] [Real] NULL " & _
                ") ON [PRIMARY]"
40        Cnxn(0).Execute sql
50    End If

60    Exit Sub

CheckBioDefIndexInDb_Error:

      Dim strES As String
      Dim intEL As Integer

70    intEL = Erl
80    strES = Err.Description
90    LogError "modDbDesign", "CheckBioDefIndexInDb", intEL, strES, sql

End Sub

'---------------------------------------------------------------------------------------
' Sub       : SubDesign
' Author    : Trevor
' Date      : 13/11/2015
' Purpose   :New printing disable facility to exclude Doctors by Department
'---------------------------------------------------------------------------------------
Public Sub CheckDisablePrintingInDb()
    Dim sql As String

If IsTableInDatabase("DisablePrinting") = False Then
    sql = "CREATE TABLE [dbo].[DisablePrinting]( " & _
          "[GPCode] [nvarchar](20) NULL, " & _
          "[GPName] [nvarchar](50) NULL, " & _
          "[Type] [nvarchar](50) NULL, " & _
          "[Disable] [bit] NULL, " & _
          "[Department] [nvarchar](50) NULL, " & _
          ") ON [PRIMARY] "
    Cnxn(0).Execute sql
End If

Exit Sub
End Sub

Public Sub CheckLabLinkMappingInDb()

On Error GoTo CheckLabLinkMappingInDb_Error

Dim sql As String

If IsTableInDatabase("LabLinkMapping") = False Then
    sql = "CREATE TABLE [dbo].[LabLinkMapping]( " & _
          "[MappingType] [nvarchar](20) NULL, " & _
          "[TargetHospital] [nvarchar](50) NULL, " & _
          "[SourceValue] [nvarchar](50) NULL, " & _
          "[TargetValue] [nvarchar](50) NULL, " & _
          "[Counter] [int] IDENTITY(1,1) NOT NULL " & _
          ") ON [PRIMARY] "
    Cnxn(0).Execute sql
End If

Exit Sub

CheckLabLinkMappingInDb_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "modDbDesign", "CheckLabLinkMappingInDb", intEL, strES, sql

End Sub


'---------------------------------------------------------------------------------------
' Procedure : CheckBiomnisRequestsInDb
' Author    : Masood
' Date      : 26/May/2016
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub CheckBiomnisRequestsInDb()


      Dim sql As String

10      On Error GoTo CheckBiomnisRequestsInDb_Error


20    If IsTableInDatabase("BiomnisRequests") = False Then
30        sql = "CREATE TABLE [dbo].[BiomnisRequests]( " & _
                "[SampleID] [nvarchar](50) NULL, " & _
                "[TestCode] [nvarchar](50) NULL, " & _
                "[TestName] [nvarchar](200) NULL, " & _
                "[SampleType] [nvarchar](200) NULL, " & _
                "[SampleDateTime] [datetime] NULL, " & _
                "[Department] [nvarchar](50) NULL, " & _
                "[RequestedBy] [nvarchar](50) NULL, " & _
                "[SendTo] [nvarchar](100) NULL, " & _
                "[Status] [nvarchar](50) NULL, " & _
                "[DateTimeOfRecord] [datetime] NULL " & _
                ") ON [PRIMARY] "
40        Cnxn(0).Execute sql
          
50       sql = " ALTER TABLE [dbo].[BiomnisRequests] ADD  CONSTRAINT [DF_BiomnisRequests_DateTimeOfRecord]  DEFAULT (getdate()) FOR [DateTimeOfRecord]"
60       Cnxn(0).Execute (sql)

70    End If


       
80    Exit Sub

       
CheckBiomnisRequestsInDb_Error:

      Dim strES As String
      Dim intEL As Integer

90    intEL = Erl
100   strES = Err.Description
110   LogError "modDbDesign", "CheckBiomnisRequestsInDb", intEL, strES, sql

End Sub


Public Sub CheckLabLinkConnectionConfigInDb()


10    On Error GoTo CheckLabLinkConnectionConfigInDb_Error

      Dim sql As String

20    If IsTableInDatabase("LabLinkConnectionConfig") = False Then
30        sql = "CREATE TABLE [dbo].[LabLinkConnectionConfig]( " & _
                "[LabName] [nvarchar](50) NULL, " & _
                "[LocalIPAddress] [nvarchar](20) NULL, " & _
                "[LocalPort] [nvarchar](10) NULL, " & _
                "[RemoteIPAddress] [nvarchar](20) NULL, " & _
                "[RemotePortIN] [nvarchar](10) NULL, " & _
                "[RemotePortOut] [nvarchar](10) NULL, " & _
                "[RequestFolder] [nvarchar](50) NULL, " & _
                "[RequestFolderCopy] [nvarchar](50) NULL, " & _
                "[ResultFolder] [nvarchar](50) NULL, " & _
                "[ResultFolderCopy] [nvarchar](50) NULL, " & _
                "[Counter] [int] IDENTITY(1,1) NOT NULL " & _
                ") ON [PRIMARY]"
40        Cnxn(0).Execute sql
50    End If

60    Exit Sub

CheckLabLinkConnectionConfigInDb_Error:

      Dim strES As String
      Dim intEL As Integer

70    intEL = Erl
80    strES = Err.Description
90    LogError "modDbDesign", "CheckLabLinkConnectionConfigInDb", intEL, strES, sql

End Sub

Public Sub CheckLabLinkCommunicationInDb()

On Error GoTo CheckBioDefIndexInDb_Error

Dim sql As String
If IsTableInDatabase("LabLinkCommunication") = False Then
    sql = "CREATE TABLE [dbo].[LabLinkCommunication]( " & _
          "[SampleID] [nvarchar](50) NOT NULL, " & _
          "[Department] [nvarchar](50) NOT NULL, " & _
          "[SourceHospital] [nvarchar](50) NULL, " & _
          "[TargetHospital] [nvarchar](50) NULL, " & _
          "[MessageType] [nvarchar](50) NULL, " & _
          "[Status] [nvarchar](50) NULL, " & _
          "[MessageState] [tinyint] NULL, " & _
          "[DateTimeOfRecord] [datetime] NULL, " & _
          "[Counter] [int] IDENTITY(1,1) NOT NULL " & _
          ") ON [PRIMARY]"
    Cnxn(0).Execute sql
End If

Exit Sub

CheckBioDefIndexInDb_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "modDbDesign", "CheckLabLinkCommunicationInDb", intEL, strES, sql

End Sub

'---------------------------------------------------------------------------------------
' Procedure : SampleAddedtoConsultantList
' Author    : XPMUser
' Date      : 2/19/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function SampleAddedtoConsultantList(SampleID As String, Dept As String) As Boolean    'Masood 19_Feb_2013
          Dim sql As String
          Dim tb As New ADODB.Recordset
10        On Error GoTo SampleAddedtoConsultantList_Error

20        sql = "SELECT * from ConsultantList WHERE " & _
                "SampleID = '" & SampleID & "' And Department ='" & Dept & "'"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If tb.EOF Then
60            SampleAddedtoConsultantList = False
70        Else
80            SampleAddedtoConsultantList = True
90        End If


100       Exit Function


SampleAddedtoConsultantList_Error:

          Dim strES As String
          Dim intEL As Integer

110       intEL = Erl
120       strES = Err.Description
130       LogError "modDbDesign", "SampleAddedtoConsultantList", intEL, strES, sql
End Function


Public Function SampleRelasedtoConsultant(SampleID As String, Dept As String) As Boolean    'Farhan 17_un_2015
          Dim sql As String
          Dim tb As New ADODB.Recordset
10        On Error GoTo SampleRelasedtoConsultant_Error

20        sql = "SELECT * from ConsultantList WHERE " & _
                "SampleID = '" & SampleID & "' And Department ='" & Dept & "' And status =0"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If tb.EOF Then
60            SampleRelasedtoConsultant = False
70        Else
80            SampleRelasedtoConsultant = True
90        End If


100       Exit Function


SampleRelasedtoConsultant_Error:

          Dim strES As String
          Dim intEL As Integer

110       intEL = Erl
120       strES = Err.Description
130       LogError "modDbDesign", "SampleRelasedtoConsultant", intEL, strES, sql
End Function
