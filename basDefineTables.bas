Attribute VB_Name = "basDefineTables"
                         
Option Explicit


Public Type FieldDefs
    ColumnName As String
    DataType As String
    Length As Long
    NoNull As Boolean
    DirectionASC As Boolean
End Type

Public Design() As FieldDefs
Public Sub DefineTableDefinitions()

10        On Error GoTo DefineTableDefinitions_Error

20        ReDim Design(0 To 4) As FieldDefs

30        FillDesignL 0, "TableName", "nvarchar", 50
40        FillDesignL 1, "ColumnName", "nvarchar", 50
50        FillDesignL 2, "Size", "int"
60        FillDesignL 3, "ColumnType", "nvarchar", 50
70        FillDesignL 4, "IsNullable", "int"

80        If IsTableInDB("TableDefinitions") = False Then    'There is no table  in database
90            CreateTable "TableDefinitions"
100       Else
110           DoTableAnalysis "TableDefinitions"
120       End If

130       Exit Sub

DefineTableDefinitions_Error:

          Dim strES As String
          Dim intEL As Integer

140       intEL = Erl
150       strES = Err.Description
160       LogError "basDefineTables", "DefineTableDefinitions", intEL, strES


End Sub

Public Function IsTableInDB(ByVal TableName As String) As Boolean

          Dim tbExists As Recordset
          Dim sql As String
          Dim RetVal As Boolean

          'How to find if a table exists in a database
          'open a recordset with the following sql statement:
          'Code:SELECT name FROM sysobjects WHERE xtype = 'U' AND name = 'MyTable'
          'If the recordset it at eof then the table doesn't exist
          'if it has a record then the table does exist.

10        On Error GoTo IsTableInDB_Error

20        sql = "SELECT name FROM sysobjects WHERE " & _
                "xtype = 'U' " & _
                "AND name = '" & TableName & "'"
30        Set tbExists = Cnxn(0).Execute(sql)

40        RetVal = True

50        If tbExists.EOF Then    'There is no table <TableName> in database
60            RetVal = False
70        End If
80        IsTableInDB = RetVal

90        Exit Function

IsTableInDB_Error:

          Dim strES As String
          Dim intEL As Integer

100       intEL = Erl
110       strES = Err.Description
120       LogError "modDbDesign", "IsTableInDB", intEL, strES, sql

End Function

Public Sub DoTableAnalysis(ByVal TableName As String)

          Dim n As Long
          Dim f As Long
          Dim Found As Boolean
          Dim Matching As Boolean
          Dim sql As String
          Dim tb As Recordset
          Dim s As String

10        On Error GoTo DoTableAnalysis_Error

20        sql = "Select top 1 * from [" & TableName & "]"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        For n = 0 To UBound(Design)
60            Found = False
70            Matching = False
80            For f = 0 To tb.Fields.Count - 1
90                If UCase$(tb.Fields(f).Name) = UCase$(Design(n).ColumnName) Then
100                   Found = True
110                   If ((tb.Fields(f).Type = 2 And UCase$(Design(n).DataType) = "SMALLINT") Or _
                          (tb.Fields(f).Type = 4 And UCase$(Design(n).DataType) = "REAL") Or _
                          (tb.Fields(f).Type = 5 And UCase$(Design(n).DataType) = "FLOAT") Or _
                          (tb.Fields(f).Type = 16 And UCase$(Design(n).DataType) = "TINYINT") Or _
                          (tb.Fields(f).Type = 17 And UCase$(Design(n).DataType) = "TINYINT") Or _
                          (tb.Fields(f).Type = 203 And UCase$(Design(n).DataType) = "NTEXT") Or _
                          (tb.Fields(f).Type = 205 And UCase$(Design(n).DataType) = "IMAGE") Or _
                          (tb.Fields(f).Type = 135 And UCase$(Design(n).DataType) = "DATETIME") Or _
                          (tb.Fields(f).Type = 131 And UCase$(Design(n).DataType) = "NUMERIC") Or _
                          (tb.Fields(f).Type = 11 And UCase$(Design(n).DataType) = "BIT") Or _
                          (tb.Fields(f).Type = 129 And UCase$(Design(n).DataType) = "CHAR" And tb.Fields(f).DefinedSize = Design(n).Length) Or _
                          (tb.Fields(f).Type = 200 And UCase$(Design(n).DataType) = "VARCHAR" And tb.Fields(f).DefinedSize = Design(n).Length) Or _
                          (tb.Fields(f).Type = 130 And UCase$(Design(n).DataType) = "NCHAR" And tb.Fields(f).DefinedSize = Design(n).Length) Or _
                          (tb.Fields(f).Type = 202 And UCase$(Design(n).DataType) = "NVARCHAR" And tb.Fields(f).DefinedSize = Design(n).Length) Or _
                          (tb.Fields(f).Type = 3 And UCase$(Design(n).DataType) = "INT")) Then
120                       Matching = True
130                   End If
140                   Exit For
150               End If
160           Next
170           s = ""
180           If Not Found Then
190               sql = "ALTER TABLE [" & TableName & "] " & _
                        "ADD [" & Design(n).ColumnName & "] " & Design(n).DataType & " "
200               If Design(n).Length <> 0 Then
210                   sql = sql & "(" & Design(n).Length & ") "
220               End If
230               If Design(n).NoNull = True Then
240                   sql = sql & "NOT "
250               End If
260               sql = sql & "NULL "
270               Cnxn(0).Execute sql
280           ElseIf Not Matching Then
290               If tb.Fields(f).DefinedSize < Design(n).Length Then
300                   s = TableName & vbCrLf & "Column '" & tb.Fields(f).Name & "' " & _
                          "should be " & Design(n).DataType & "(" & Design(n).Length & ")"
310                   MsgBox s
320               End If
330           End If
340       Next

350       If Right$(TableName, 3) = "Arc" Then
360           Found = False
370           For f = 0 To tb.Fields.Count - 1
380               If tb.Fields(f).Name = "ArchiveDateTime" Then
390                   Found = True
400                   Exit For
410               End If
420           Next
430           If Not Found Then
440               sql = "ALTER TABLE [" & TableName & "] ADD " & _
                        "[ArchiveDateTime] datetime"
450               Cnxn(0).Execute sql
460           End If
470           Found = False
480           For f = 0 To tb.Fields.Count - 1
490               If tb.Fields(f).Name = "ArchivedBy" Then
500                   Found = True
510                   Exit For
520               End If
530           Next
540           If Not Found Then
550               sql = "ALTER TABLE [" & TableName & "] ADD " & _
                        "[ArchivedBy] nvarchar (50)"
560               Cnxn(0).Execute sql
570           End If
580       End If


590       Exit Sub

DoTableAnalysis_Error:

          Dim strES As String
          Dim intEL As Integer

600       intEL = Erl
610       strES = Err.Description
620       LogError "basDefineTables", "DoTableAnalysis", intEL, strES


End Sub

Public Sub CreateTable(ByVal TableName As String)

          Dim sql As String
          Dim n As Integer

10        On Error GoTo CreateTable_Error

20        sql = "CREATE TABLE " & TableName & " ( "
30        For n = 0 To UBound(Design)
40            sql = sql & "[" & Design(n).ColumnName & "] " & _
                    Design(n).DataType & " "
50            If Design(n).Length <> 0 Then
60                sql = sql & "(" & Design(n).Length & ") "
70            End If
80            sql = sql & IIf(Design(n).NoNull, " NOT NULL, ", "NULL, ")
90        Next
100       sql = Left$(sql, Len(sql) - 2) & ")"
110       Cnxn(0).Execute sql

          'If Right$(TableName, 3) = "Arc" Then
          '  sql = "ALTER TABLE [" & TableName & "] ADD " & _
             '        "[ArchiveDateTime] datetime, " & _
             '        "[ArchivedBy] nvarchar (50)"
          '  Cnxn(0).Execute sql
          'End If

120       Exit Sub

CreateTable_Error:

          Dim strES As String
          Dim intEL As Integer

130       intEL = Erl
140       strES = Err.Description
150       LogError "basDefineTables", "CreateTable", intEL, strES


End Sub


Public Sub CreateIndex(ByVal TableName As String, _
                       ByVal IndexName As String, _
                       ByRef cD() As FieldDefs, _
                       ByVal Unique As Boolean, _
                       ByVal Clustered As Boolean)

          Dim sql As String
          Dim n As Integer

10        On Error GoTo CreateIndex_Error

20        sql = "IF NOT EXISTS (SELECT * FROM sysindexes WHERE " & _
                "               name = '" & TableName & "_" & IndexName & "') " & _
                "BEGIN " & _
                "  CREATE " & IIf(Unique, "UNIQUE", "") & " " & _
                "  " & IIf(Clustered, "CLUSTERED", "NONCLUSTERED") & " " & _
                "  INDEX " & TableName & "_" & IndexName & " " & _
                "  ON " & TableName & " " & _
                "  ("
30        For n = 0 To UBound(Design)
40            sql = sql & cD(n).ColumnName & " " & IIf(cD(n).DirectionASC, "ASC", "DESC") & ","
50        Next
60        sql = Left$(sql, Len(sql) - 1) & ") END"
70        Cnxn(0).Execute sql

80        Exit Sub

CreateIndex_Error:

          Dim strES As String
          Dim intEL As Integer

90        intEL = Erl
100       strES = Err.Description
110       LogError "basDefineTables", "CreateIndex", intEL, strES


End Sub

Public Sub CheckBuild()

10        On Error GoTo CheckBuild_Error

20        DefineTableDefinitions

30        Exit Sub

CheckBuild_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "basDefineTables", "CheckBuild", intEL, strES


End Sub
Public Sub CheckBuildDatabase()

10        On Error GoTo CheckBuildDatabase_Error

20        DefineTableGeneric

30        Exit Sub

CheckBuildDatabase_Error:

          Dim strES As String
          Dim intEL As Integer

40        intEL = Erl
50        strES = Err.Description
60        LogError "basDefineTables", "CheckBuildDatabase", intEL, strES


End Sub


Public Sub FillDesignL(ByVal ColumnIndex As Long, _
                       ByVal ColumnName As String, _
                       ByVal DataType As String, _
                       Optional ByVal DataLength As Integer, _
                       Optional ByVal NoNull As Boolean)

10        On Error GoTo FillDesignL_Error

20        Design(ColumnIndex).ColumnName = ColumnName
30        Design(ColumnIndex).DataType = DataType
40        Design(ColumnIndex).Length = DataLength
50        Design(ColumnIndex).NoNull = NoNull

60        Exit Sub

FillDesignL_Error:

          Dim strES As String
          Dim intEL As Integer

70        intEL = Erl
80        strES = Err.Description
90        LogError "basDefineTables", "FillDesignL", intEL, strES


End Sub

Public Sub CheckConfigFile()

          Dim f As Integer
          Dim s As String
          Dim Lines() As String
          Dim sql As String
          Dim DirName As String
          Dim R As String
          Dim n As Long


          'SELECT     SO.name AS TableName, SC.name AS ColumnName, sc.prec AS Length,
          '                      CASE sc.type WHEN 39 THEN 'nvarchar' WHEN 38 THEN 'int' WHEN 56 THEN 'int' WHEN 63 THEN 'numeric' WHEN 52 THEN 'smallint' WHEN 50 THEN
          '                       'bit' WHEN 109 THEN 'real' WHEN 111 THEN 'datetime' WHEN 239 THEN 'nchar' WHEN 106 THEN 'decimal' WHEN 47 THEN 'char' WHEN 61 THEN 'datetime'
          '                       WHEN 108 THEN 'numeric' WHEN 48 THEN 'tinyint' WHEN 35 THEN 'ntext' WHEN 34 THEN 'image' WHEN 37 THEN 'uniqueidentifier' END AS type,
          '                      sc.isnullable AS isnullable
          'FROM         sysobjects AS SO, syscolumns AS SC
          'WHERE     So.id IN
          '                          (SELECT     object_id(name)
          '                            From sysobjects
          '                            WHERE      type = 'U' AND status > 0) AND so.id = sc.id
          'ORDER BY tablename

10        On Error GoTo CheckConfigFile_Error

20        R = App.Path

30        n = InStr(App.Path, "NetAcquire")

40        R = Left(R, n - 1)

50        R = R & "NetAcquire Config\"



60        If Dir(R & "dbConfig.txt", 0) <> "" Then
70            sql = "Delete from TableDefinitions"
80            Cnxn(0).Execute sql
90            f = FreeFile
100           Open R & "dbConfig.txt" For Input As f
110           Do While Not EOF(f)
120               Line Input #f, s
130               Lines = Split(s, vbTab)
140               sql = "INSERT INTO TableDefinitions (Tablename, ColumnName, Size, ColumnType, IsNullable) VALUES " & _
                        "('" & Lines(1) & "', " & _
                        " '" & Lines(2) & "', " & _
                        " '" & Lines(3) & "', " & _
                        " '" & Lines(4) & "', " & _
                        " '" & Lines(5) & "')"
150               Cnxn(0).Execute sql
160           Loop
170           Close f
180           DirName = Dir(R & "Processed", vbDirectory)
190           If DirName = "" Then
200               MkDir R & "Processed"
210           End If
220           FileCopy R & "dbConfig.txt", R & "Processed\dbConfig" & Format$(Now, "yyyyMMddhhmm") & ".txt"
230           FileCopy R & "dbConfig.txt", R & "Processed\dbConfig.txt"
240           Kill R & "dbConfig.txt"
250       End If


260       CheckBuildDatabase
270       Debug.Print Now


280       Exit Sub

CheckConfigFile_Error:

          Dim strES As String
          Dim intEL As Integer

290       intEL = Erl
300       strES = Err.Description
310       LogError "basDefineTables", "CheckConfigFile", intEL, strES


End Sub

Public Sub DefineTableGeneric()

          Dim sql As String
          Dim tb As Recordset
          Dim sn As Recordset
          Dim TableName As String
          Dim n As Integer
          Dim pr As Double
          Dim prCount As Double

10        On Error GoTo DefineTableGeneric_Error
20        prCount = 24
30        sql = "SELECT count(*) as tot FROM TableDefinitions"
40        Set tb = New Recordset
50        RecOpenServer 0, tb, sql
60        If Not tb.EOF Then
70            If tb!Tot = 0 Then
80                pr = 0.03
90            Else
100               pr = Format(74 / tb!Tot, "0.###")
110           End If
120       Else
130           pr = 0.03
140       End If

150       sql = "SELECT DISTINCT TableName FROM TableDefinitions order by tablename"
160       Set tb = New Recordset
170       RecOpenServer 0, tb, sql
180       Do While Not tb.EOF
190           sql = "SELECT COUNT(*) AS Tot FROM TableDefinitions WHERE " & _
                    "TableName = '" & tb!TableName & "'"
200           Set sn = New Recordset
210           RecOpenServer 0, sn, sql
220           ReDim Design(0 To sn!Tot - 1) As FieldDefs
230           sql = "SELECT  * FROM TableDefinitions WHERE " & _
                    "TableName = '" & tb!TableName & "'"
240           Set sn = New Recordset
250           RecOpenServer 0, sn, sql

260           n = 0
270           Do While Not sn.EOF
280               prCount = prCount + pr
290               Select Case sn!ColumnType
                  Case "nvarchar"
300                   If sn!isnullable = 0 Then
310                       FillDesignL n, sn!ColumnName, sn!ColumnType, sn!Size, True
320                   Else
330                       FillDesignL n, sn!ColumnName, sn!ColumnType, sn!Size
340                   End If
350               Case Else
360                   If sn!isnullable = 0 Then
370                       FillDesignL n, sn!ColumnName, sn!ColumnType, , True
380                   Else
390                       FillDesignL n, sn!ColumnName, sn!ColumnType
400                   End If
410               End Select
420               n = n + 1
430               sn.MoveNext
440           Loop
450           If IsTableInDB(tb!TableName) = False Then    'There is no table  in database
460               CreateTable tb!TableName
470           Else
480               DoTableAnalysis tb!TableName
490           End If
500           tb.MoveNext
510       Loop

520       Exit Sub

DefineTableGeneric_Error:

          Dim strES As String
          Dim intEL As Integer

530       intEL = Erl
540       strES = Err.Description
550       LogError "basDefineTables", "DefineTableGeneric", intEL, strES, sql

End Sub

