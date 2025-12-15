Attribute VB_Name = "modAddDataMacrosDM"
Option Compare Database
Option Explicit

'=====================================================================================================================
' COMPLETE AUDIT DATA MACRO SYSTEM
'
' This module provides a complete auditing solution using Access Data Macros
' with special handling for Long Text fields via LookupRecord.
'NOTE: Because data macros are created on tables in the same accdb, this code must be run in the BE, not in the FE.
'            The main audit table, tblAuditLog, must be created in the BE and MUST be linked to the FE for use in restoring records when required.
'            The tblLongTextBackup table must be created in the BE and it MUST be linked to the FE
'              because both the Data Macros in the BE and the VBA in the FE must have read/write access to this table.
'             The tblAuditLogConfig table must be created in the BE and MUST be linked to the FE
'              because both the Data Macros in the BE and the VBA in the FE must have read/write access to this table.

' SETUP INSTRUCTIONS:
'  BACK END Set up
' 1. Run CreateAuditTables() to create the 3 required tables- - tblAuditLog, tblLongTextBackup, and tblAuditLogConfig
' 2. Run PopulateConfigTable() to populate the config with your auditable tables/fields
' 3. Run GenerateAllAuditDataMacros() to create all Data Macros
'  FRONT END Set up
'  Link to the audit tables
'=====================================================================================================================

'-----------------------------------------------------------------------------
' STEP 1: Create 3 required tables
'-----------------------------------------------------------------------------
Public Sub One_CreateAuditTables()
100       On Error GoTo ErrorHandler
          
      Dim db As DAO.Database
      Dim tdf As DAO.TableDef
      Dim fld As DAO.Field
      Dim idx As DAO.Index
          
110       Set db = CurrentDb
          
          ' ========== Create tblAuditLog ==========
120       On Error Resume Next
130       Set tdf = db.TableDefs("tblAuditLog")
140       If Not tdf Is Nothing Then
150   Debug.Print "tblAuditLog already exists"
160           GoTo CreateLongTextBackup
170       End If
180       On Error GoTo ErrorHandler
          
190       Set tdf = db.CreateTableDef("tblAuditLog")
          
200       Set fld = tdf.CreateField("AuditLogID", dbLong)
210       fld.Attributes = dbAutoIncrField
220       tdf.Fields.Append fld
          
230       Set fld = tdf.CreateField("TableName", dbText, 50)
240       fld.Required = True
250       tdf.Fields.Append fld
          
260       Set fld = tdf.CreateField("PrimaryKey", dbLong)
270       tdf.Fields.Append fld
          
280       Set fld = tdf.CreateField("FieldName", dbText, 50)
290       fld.Required = True
300       tdf.Fields.Append fld
          
310       Set fld = tdf.CreateField("OldValue", dbMemo)
320       tdf.Fields.Append fld
          
330       Set fld = tdf.CreateField("NewValue", dbMemo)
340       tdf.Fields.Append fld
          
350       Set fld = tdf.CreateField("DateChanged", dbDate)
360       fld.Required = True
370       tdf.Fields.Append fld
          
380       Set fld = tdf.CreateField("ChangedBy", dbText, 50)
390       fld.Required = True
400       tdf.Fields.Append fld
          
410       db.TableDefs.Append tdf
          
420       Set idx = tdf.CreateIndex("PrimaryKey")
430       idx.Primary = True
440       idx.Required = True
450       Set fld = idx.CreateField("AuditLogID")
460       idx.Fields.Append fld
470       tdf.Indexes.Append idx
          
480   Debug.Print "tblAuditLog created"
          
CreateLongTextBackup:
          ' ========== Create tblLongTextBackup ==========
          'Optional, useful only when Long Text field data is <= ~2034 characters long and can be edited by VBA
          
490       Set tdf = Nothing
500       On Error Resume Next
510       Set tdf = db.TableDefs("tblLongTextBackup")
520       If Not tdf Is Nothing Then
530   Debug.Print "tblLongTextBackup already exists"
540           GoTo CreateConfig
550       End If
560       On Error GoTo ErrorHandler
          
570       Set tdf = db.CreateTableDef("tblLongTextBackup")
          
580       Set fld = tdf.CreateField("BackupID", dbLong)
590       fld.Attributes = dbAutoIncrField
600       tdf.Fields.Append fld
          
610       Set fld = tdf.CreateField("TableName", dbText, 50)
620       fld.Required = True
630       tdf.Fields.Append fld
          
640       Set fld = tdf.CreateField("PrimaryKey", dbLong)
650       fld.Required = True
660       tdf.Fields.Append fld
          
670       Set fld = tdf.CreateField("FieldName", dbText, 50)
680       fld.Required = True
690       tdf.Fields.Append fld
          
700       Set fld = tdf.CreateField("OldValue", dbMemo)
710       tdf.Fields.Append fld
          
720       Set fld = tdf.CreateField("DateChanged", dbDate)
730       fld.Required = True
740       tdf.Fields.Append fld
          
750       Set fld = tdf.CreateField("ChangedBy", dbText, 50)
760       fld.Required = True
770       tdf.Fields.Append fld
          
780       db.TableDefs.Append tdf
          
790       Set idx = tdf.CreateIndex("PrimaryKey")
800       idx.Primary = True
810       idx.Required = True
820       Set fld = idx.CreateField("BackupID")
830       idx.Fields.Append fld
840       tdf.Indexes.Append idx
          
850   Debug.Print "tblLongTextBackup created"
          
CreateConfig:
          ' ========== Create tblAuditLogConfig ==========
860       Set tdf = Nothing
870       On Error Resume Next
880       Set tdf = db.TableDefs("tblAuditLogConfig")
890       If Not tdf Is Nothing Then
900   Debug.Print "tblAuditLogConfig already exists"
910           GoTo Cleanup
920       End If
930       On Error GoTo ErrorHandler
          
940       Set tdf = db.CreateTableDef("tblAuditLogConfig")
          
950       Set fld = tdf.CreateField("ConfigID", dbLong)
960       fld.Attributes = dbAutoIncrField
970       tdf.Fields.Append fld
          
980       Set fld = tdf.CreateField("TableName", dbText, 50)
990       fld.Required = True
1000      tdf.Fields.Append fld
          
1010      Set fld = tdf.CreateField("FieldName", dbText, 50)
1020      fld.Required = True
1030      tdf.Fields.Append fld
          
1040      Set fld = tdf.CreateField("FieldPosition", dbLong)
1050      fld.Required = True
1060      tdf.Fields.Append fld
          
1070      Set fld = tdf.CreateField("DataType", dbLong)
1080      fld.Required = True
1090      tdf.Fields.Append fld
          
1100      Set fld = tdf.CreateField("IsPrimaryKey", dbBoolean)
1110      fld.Required = True
1120      fld.DefaultValue = "False"
1130      tdf.Fields.Append fld
          
1140      Set fld = tdf.CreateField("IsAuditable", dbBoolean)
1150      fld.Required = True
1160      fld.DefaultValue = "True"
1170      tdf.Fields.Append fld
1180      db.TableDefs.Append tdf
          
1190      Set idx = tdf.CreateIndex("PrimaryKey")
1200      idx.Primary = True
1210      idx.Required = True
1220      Set fld = idx.CreateField("ConfigID")
1230      idx.Fields.Append fld
1240      tdf.Indexes.Append idx
          
1250  Debug.Print "tblAuditLogConfig created"
          
Cleanup:
1260      Set fld = Nothing
1270      Set idx = Nothing
1280      Set tdf = Nothing
1290      Set db = Nothing
          
1300      MsgBox "Audit tables created successfully!", vbInformation
1310      Exit Sub
          
ErrorHandler: 'If desired, replace with your own Error Handling
1320      MsgBox "Error creating tables: " & Err.Number & " - " & Err.Description, vbCritical
1330      Resume Cleanup
End Sub


'-----------------------------------------------------------------------------
' STEP 2: Populate configuration table with your tables and fields
'         Customize by including/excluding specific tables and fields in your database.
'-----------------------------------------------------------------------------
Public Sub Two_PopulateConfigTable()
100       On Error GoTo ErrorHandler
          
      Dim db As DAO.Database
      Dim tdef As DAO.TableDef
      Dim fld As DAO.Field
      Dim idx As DAO.index
      Dim pkField As DAO.Field
      Dim strSQL As String
      Dim isPK As Boolean
      Dim pkFieldName As String
          
110       Set db = CurrentDb
          
          ' Clear existing config
120       db.Execute "DELETE * FROM tblDataMacroConfig"
          
          ' Loop through all tables
130       For Each tdef In db.TableDefs
              ' Filter: Include only tables starting with "tbl"
              ' Exclude: System tables, audit tables, other tables you don't want to audit
              'This should be implemented as a table driven solution, with all auditable tables and auditable fields flagged or unflagged as appropriate.
              'The audit log tables could also be prefixed "USys" to avoid hard-coding them here or flagging them in the config table.
140           If Left(tdef.Name, 3) = "tbl" 
                    '_
                  'And tdef.Name <> "tblAuditLog" _
                  'And tdef.Name <> "tblLongTextBackup" _
                  'And tdef.Name <> "tblDataMacroConfig" _
                  'And tdef.Name <> "tblLoadTime" _
                  'And tdef.Name <> "tblPublicationHistory" _
                  'And Left(tdef.Name, 7) <> "tblPUTT" 
Then
                  
                  ' Get primary key field name for this table
150               pkFieldName = ""
160               For Each idx In tdef.Indexes
170                   If idx.Primary Then
180                       For Each pkField In idx.Fields
190                           pkFieldName = pkField.Name
200                           Exit For
210                       Next pkField
220                       Exit For
230                   End If
240               Next idx
                  
                  ' Add each field to be audited (excluding specific field names you don't want to audit)
                  ' Instead of hard-coding excluded fields, this step should be modified to add all fields.
                  ' Then, an additional "IsAudited" field in the config table can be used to 
                  '  review and manually flag excluded fields.
250               For Each fld In tdef.Fields
260                  ' If fld.Name <> "AccessTS" _
                     '     And fld.Name <> "SSMA_TimeStamp" _
                     '     And fld.Name <> "ValidFrom" _
                     '     And fld.Name <> "ValidTo" Then
                          
                          ' Check if this field is the primary key
270                       isPK = (fld.Name = pkFieldName)
                          
280                       strSQL = "INSERT INTO tblDataMacroConfig (TableName, FieldName, DataType, IsPrimaryKey, IsAuditable) " & _
                              "VALUES ('" & tdef.Name & "', '" & fld.Name & "', " & fld.type & ", " & isPK & ", " & -1 & ")" 'default to IsAuditable = true
290                       db.Execute strSQL
300                   End If
310               Next fld
320           End If
330       Next tdef
          
340       Set pkField = Nothing
350       Set idx = Nothing
360       Set fld = Nothing
370       Set tdef = Nothing
380       Set db = Nothing
          
390       MsgBox "Configuration table populated successfully!", vbInformation
Cleanup:
400       Exit Sub
          
ErrorHandler: 'If desired, replace with your own Error Handling
410       MsgBox "Error populating config: " & Err.Number & " - " & Err.Description, vbCritical
420       Resume Cleanup
End Sub

'-----------------------------------------------------------------------------
' STEP 3: Generate 3 or 5 (tables containing Long Text fields) Data Macros for all tables
'-----------------------------------------------------------------------------
Public Sub Three_GenerateAllAuditDataMacros()
100       On Error GoTo ErrorHandler
          
      Dim db As DAO.Database
      Dim rs As DAO.Recordset
      Dim dictTables As Object
      Dim tableName As String
      Dim FieldName As String
      Dim fieldDataType As Long
      Dim fieldIsPK As Boolean
      Dim fieldList As Collection
      Dim fieldInfo As Variant
      Dim tempPath As String
      Dim tableCount As Long
      Dim currentTable As Variant
          
110       Set db = CurrentDb
120       Set dictTables = CreateObject("Scripting.Dictionary")
          
          ' Read configuration and group by table
130       Set rs = db.OpenRecordset("SELECT TableName, FieldName, DataType, IsPrimaryKey FROM tblDataMacroConfig ORDER BY TableName, FieldName", dbOpenSnapshot)
          
140       Do While Not rs.EOF
150           tableName = Nz(rs!tableName, "")
160           FieldName = Nz(rs!FieldName, "")
170           fieldDataType = Nz(rs!DataType, 0)
180           fieldIsPK = Nz(rs!IsPrimaryKey, False)
              
190           If tableName <> "" And FieldName <> "" Then
200               If Not dictTables.Exists(tableName) Then
210                   Set fieldList = New Collection
220                   dictTables.Add tableName, fieldList
230               Else
240                   Set fieldList = dictTables(tableName)
250               End If
                  
                  ' Store field info as array: (FieldName, DataType, IsPrimaryKey)
260               fieldInfo = Array(FieldName, fieldDataType, fieldIsPK)
270               fieldList.Add fieldInfo
280           End If
              
290           rs.MoveNext
300       Loop
310       rs.Close
          
320       tempPath = Environ("TEMP") & "\"
          
          ' Process each table
330       tableCount = 0
340       For Each currentTable In dictTables.Keys
350           tableName = CStr(currentTable)
360           Set fieldList = dictTables(tableName)
              
370   Debug.Print "Processing: " & tableName & " (" & fieldList.count & " fields)"
              
380           Call CreateAllDataMacros(tableName, fieldList, tempPath)
              
390           tableCount = tableCount + 1
400       Next currentTable
          
410       Set rs = Nothing
420       Set db = Nothing
430       Set dictTables = Nothing
          
440       MsgBox "Successfully generated audit data macros for " & tableCount & " tables!", vbInformation
450   Debug.Print "Completed: " & tableCount & " tables processed"

Cleanup:
460       Exit Sub
          
ErrorHandler: 'If desired, replace with your own Error Handling
470       MsgBox "Error: " & Err.Number & " - " & Err.Description, vbCritical
480       Resume Cleanup
End Sub

'-----------------------------------------------------------------------------
' Generate 3 or 5 Data Macros in a single XML file
' Called by Three_GenerateAllAuditDataMacros()
' Generate Data Macros for a table (3 After macros + 2 Before macros if Long Text fields exist)
'-----------------------------------------------------------------------------
Private Sub CreateAllDataMacros(tableName As String, fieldList As Collection, tempPath As String)
100       On Error GoTo ErrorHandler
          
      Dim xmlContent As String
      Dim fso As Object
      Dim txtFile As Object
      Dim filePath As String
      Dim primaryKeyField As String
      Dim fieldInfo As Variant
      Dim hasLongText As Boolean
          
          ' Get primary key field from config and check for Long Text fields
110       hasLongText = False
120       For Each fieldInfo In fieldList
130           If fieldInfo(2) = True Then ' IsPrimaryKey
140               primaryKeyField = fieldInfo(0) ' FieldName
150           End If
160           If fieldInfo(1) = dbMemo Then ' Check for Long Text
170               hasLongText = True
180           End If
190       Next fieldInfo
          
          ' ========== CREATE THE THREE OR FIVE MACROS IN ONE FILE ==========
          ' Start XML
200       xmlContent = "<?xml version=""1.0"" encoding=""UTF-16"" standalone=""no""?>"
210       xmlContent = xmlContent & "<DataMacros xmlns=""http://schemas.microsoft.com/office/accessservices/2010/12/application"">"
          
          ' AFTER INSERT MACRO
220       xmlContent = xmlContent & BuildAfterInsertMacro(tableName, fieldList, primaryKeyField)
          
          ' AFTER UPDATE MACRO
230       xmlContent = xmlContent & BuildAfterUpdateMacro(tableName, fieldList, primaryKeyField)
          
          ' AFTER DELETE MACRO
240       xmlContent = xmlContent & BuildAfterDeleteMacro(tableName, fieldList, primaryKeyField)
          
250       If hasLongText Then
              ' BEFORE CHANGE MACRO
260           xmlContent = xmlContent & BuildBeforeChangeMacro(tableName, fieldList, primaryKeyField)
          
              ' BEFORE DELETE MACRO
270           xmlContent = xmlContent & BuildBeforeDeleteMacro(tableName, fieldList, primaryKeyField)
280       End If
          
          ' Close root element
290       xmlContent = xmlContent & "</DataMacros>"
          
          ' Write to file and load
300       filePath = tempPath & tableName & "_DataMacros.xml"
310       Set fso = CreateObject("Scripting.FileSystemObject")
320       Set txtFile = fso.CreateTextFile(filePath, True, True)
330       txtFile.Write xmlContent
340       txtFile.Close
350       Set txtFile = Nothing
          
          ' Load the Data Macros
360       DoCmd.OpenTable tableName, acViewDesign, acHidden
370       Application.LoadFromText acTableDataMacro, tableName, filePath
380       DoCmd.Close acTable, tableName, acSaveYes
          
          ' Delete temp file
390       fso.DeleteFile filePath

Cleanup:
400       Exit Sub
          
ErrorHandler: 'If desired, replace with your own Error Handling
410       MsgBox "Error creating macros for " & tableName & ": " & Err.Number & " - " & Err.Description, vbCritical

420       Resume Cleanup
       
End Sub

'-----------------------------------------------------------------------------
' Build After Insert Macro XML
' Called by Three_GenerateAllAuditDataMacros()
'-----------------------------------------------------------------------------
Private Function BuildAfterInsertMacro(tableName As String, fieldList As Collection, primaryKeyField As String) As String
      Dim xml As String
      Dim fieldInfo As Variant
      Dim FieldName As String
          
100       xml = "<DataMacro Event=""AfterInsert""><Statements>"
          
110       For Each fieldInfo In fieldList
120           FieldName = fieldInfo(0) ' FieldName from array
              
130           xml = xml & "<CreateRecord>"
140           xml = xml & "<Data Alias=""NewAudit""><Reference>tblAuditLog</Reference></Data>"
150           xml = xml & "<Statements>"
              
              ' TableName
160           xml = xml & "<Action Name=""SetField"">"
170           xml = xml & "<Argument Name=""Field"">NewAudit.TableName</Argument>"
180           xml = xml & "<Argument Name=""Value"">""" & tableName & """</Argument>"
190           xml = xml & "</Action>"
              
              ' PrimaryKey
200           If primaryKeyField <> "" Then
210               xml = xml & "<Action Name=""SetField"">"
220               xml = xml & "<Argument Name=""Field"">NewAudit.PrimaryKey</Argument>"
230               xml = xml & "<Argument Name=""Value"">[" & tableName & "].[" & primaryKeyField & "]</Argument>"
240               xml = xml & "</Action>"
250           End If
              
              ' FieldName
260           xml = xml & "<Action Name=""SetField"">"
270           xml = xml & "<Argument Name=""Field"">NewAudit.FieldName</Argument>"
280           xml = xml & "<Argument Name=""Value"">""" & FieldName & """</Argument>"
290           xml = xml & "</Action>"
              
              ' OldValue
300           xml = xml & "<Action Name=""SetField"">"
310           xml = xml & "<Argument Name=""Field"">NewAudit.OldValue</Argument>"
320           xml = xml & "<Argument Name=""Value"">""[NEW RECORD]""</Argument>"
330           xml = xml & "</Action>"
              
              ' NewValue
340           xml = xml & "<Action Name=""SetField"">"
350           xml = xml & "<Argument Name=""Field"">NewAudit.NewValue</Argument>"
360           xml = xml & "<Argument Name=""Value"">[" & tableName & "].[" & FieldName & "]</Argument>"
370           xml = xml & "</Action>"
              
              ' DateChanged
380           xml = xml & "<Action Name=""SetField"">"
390           xml = xml & "<Argument Name=""Field"">NewAudit.DateChanged</Argument>"
400           xml = xml & "<Argument Name=""Value"">Now()</Argument>"
410           xml = xml & "</Action>"
              
              ' ChangedBy
420           xml = xml & "<Action Name=""SetField"">"
430           xml = xml & "<Argument Name=""Field"">NewAudit.ChangedBy</Argument>"
440           xml = xml & "<Argument Name=""Value"">CurrentUser()</Argument>"
450           xml = xml & "</Action>"
              
460           xml = xml & "</Statements></CreateRecord>"
470       Next fieldInfo 'fieldName
          
480       xml = xml & "</Statements></DataMacro>"
          
490       BuildAfterInsertMacro = xml
End Function

'-----------------------------------------------------------------------------
' Build After Update Macro XML (with LookupRecord for Long Text)
' Called by Three_GenerateAllAuditDataMacros()
'-----------------------------------------------------------------------------
Private Function BuildAfterUpdateMacro(tableName As String, fieldList As Collection, primaryKeyField As String) As String
      Dim xml As String
      Dim fieldInfo As Variant
      Dim FieldName As String 'Variant
      Dim fldType As Long
      Dim isLongText As Boolean
          
100       xml = "<DataMacro Event=""AfterUpdate""><Statements>"
          
110       For Each fieldInfo In fieldList
120           FieldName = fieldInfo(0) ' FieldName
130           fldType = fieldInfo(1)   ' DataType
140           isLongText = (fldType = dbMemo)
              
              ' Skip primary key
150           If FieldName <> primaryKeyField Then
160               xml = xml & "<ConditionalBlock><If>"
170               xml = xml & "<Condition>" & GetComparisonExpression(tableName, FieldName, fldType) & "</Condition>"
180               xml = xml & "<Statements>"
                  
                  ' If Long Text, use LookUpRecord to get old value
190               If isLongText Then
200                   xml = xml & "<LookUpRecord>"
210                   xml = xml & "<Data Alias=""BackupRec"">"
220                   xml = xml & "<Reference>tblLongTextBackup</Reference>"
230                   xml = xml & "<WhereCondition>"
240                   xml = xml & "[tblLongTextBackup].[TableName]=""" & tableName & """ And "
250                   xml = xml & "[tblLongTextBackup].[PrimaryKey]=[" & tableName & "].[" & primaryKeyField & "] And "
260                   xml = xml & "[tblLongTextBackup].[FieldName]=""" & FieldName & """"
270                   xml = xml & "</WhereCondition>"
280                   xml = xml & "</Data>"
290                   xml = xml & "<Statements>"
300               End If
                  
                  ' Create audit record
310               xml = xml & "<CreateRecord>"
320               If isLongText Then
330                   xml = xml & "<Data><Reference>tblAuditLog</Reference></Data>"
340               Else
350                   xml = xml & "<Data Alias=""NewAudit""><Reference>tblAuditLog</Reference></Data>"
360               End If
370               xml = xml & "<Statements>"
                  
                  ' TableName
380               xml = xml & "<Action Name=""SetField"">"
390               If isLongText Then
400                   xml = xml & "<Argument Name=""Field"">tblAuditLog.TableName</Argument>"
410               Else
420                   xml = xml & "<Argument Name=""Field"">NewAudit.TableName</Argument>"
430               End If
440               xml = xml & "<Argument Name=""Value"">""" & tableName & """</Argument>"
450               xml = xml & "</Action>"
                  
                  ' PrimaryKey
460               If primaryKeyField <> "" Then
470                   xml = xml & "<Action Name=""SetField"">"
480                   If isLongText Then
490                       xml = xml & "<Argument Name=""Field"">tblAuditLog.PrimaryKey</Argument>"
500                   Else
510                       xml = xml & "<Argument Name=""Field"">NewAudit.PrimaryKey</Argument>"
520                   End If
530                   xml = xml & "<Argument Name=""Value"">[" & tableName & "].[" & primaryKeyField & "]</Argument>"
540                   xml = xml & "</Action>"
550               End If
                  
                  ' FieldName
560               xml = xml & "<Action Name=""SetField"">"
570               If isLongText Then
580                   xml = xml & "<Argument Name=""Field"">tblAuditLog.FieldName</Argument>"
590               Else
600                   xml = xml & "<Argument Name=""Field"">NewAudit.FieldName</Argument>"
610               End If
620               xml = xml & "<Argument Name=""Value"">""" & FieldName & """</Argument>"
630               xml = xml & "</Action>"
                  
                  ' OldValue - from LookUpRecord if Long Text, else from [Old]
640               xml = xml & "<Action Name=""SetField"">"
650               If isLongText Then
660                   xml = xml & "<Argument Name=""Field"">tblAuditLog.OldValue</Argument>"
670                   xml = xml & "<Argument Name=""Value"">[BackupRec].[OldValue]</Argument>"
680               Else
690                   xml = xml & "<Argument Name=""Field"">NewAudit.OldValue</Argument>"
700                   xml = xml & "<Argument Name=""Value"">[Old].[" & FieldName & "]</Argument>"
710               End If
720               xml = xml & "</Action>"
                  
                  ' NewValue
730               xml = xml & "<Action Name=""SetField"">"
740               If isLongText Then
750                   xml = xml & "<Argument Name=""Field"">tblAuditLog.NewValue</Argument>"
760               Else
770                   xml = xml & "<Argument Name=""Field"">NewAudit.NewValue</Argument>"
780               End If
790               xml = xml & "<Argument Name=""Value"">[" & tableName & "].[" & FieldName & "]</Argument>"
800               xml = xml & "</Action>"
                  
                  ' DateChanged
810               xml = xml & "<Action Name=""SetField"">"
820               If isLongText Then
830                   xml = xml & "<Argument Name=""Field"">tblAuditLog.DateChanged</Argument>"
840               Else
850                   xml = xml & "<Argument Name=""Field"">NewAudit.DateChanged</Argument>"
860               End If
870               xml = xml & "<Argument Name=""Value"">Now()</Argument>"
880               xml = xml & "</Action>"
                  
                  ' ChangedBy
890               xml = xml & "<Action Name=""SetField"">"
900               If isLongText Then
910                   xml = xml & "<Argument Name=""Field"">tblAuditLog.ChangedBy</Argument>"
920               Else
930                   xml = xml & "<Argument Name=""Field"">NewAudit.ChangedBy</Argument>"
940               End If
950               xml = xml & "<Argument Name=""Value"">CurrentUser()</Argument>"
960               xml = xml & "</Action>"
                  
970               xml = xml & "</Statements></CreateRecord>"
                  
                  ' Close LookUpRecord if Long Text
980               If isLongText Then
990                   xml = xml & "</Statements></LookUpRecord>"
1000              End If
                  
1010              xml = xml & "</Statements></If></ConditionalBlock>"
1020          End If
1030      Next fieldInfo 'fieldName
          
1040      xml = xml & "</Statements></DataMacro>"
          
1050      BuildAfterUpdateMacro = xml
End Function

'-----------------------------------------------------------------------------
' Build After Delete Macro XML (with LookupRecord for Long Text)
' Called by Three_GenerateAllAuditDataMacros()
'-----------------------------------------------------------------------------
Private Function BuildAfterDeleteMacro(tableName As String, fieldList As Collection, primaryKeyField As String) As String
      Dim xml As String
      Dim fieldInfo As Variant
      Dim FieldName As String
      Dim fldType As Long
      Dim isLongText As Boolean
          
100       xml = "<DataMacro Event=""AfterDelete""><Statements>"
          
110       For Each fieldInfo In fieldList
120           FieldName = fieldInfo(0) ' FieldName
130           fldType = fieldInfo(1)   ' DataType
140           isLongText = (fldType = dbMemo)
          
              ' If Long Text, use LookUpRecord
150           If isLongText Then
160               xml = xml & "<LookUpRecord>"
170               xml = xml & "<Data Alias=""BackupRec"">"
180               xml = xml & "<Reference>tblLongTextBackup</Reference>"
190               xml = xml & "<WhereCondition>"
200               xml = xml & "[tblLongTextBackup].[TableName]=""" & tableName & """ And "
210               xml = xml & "[tblLongTextBackup].[PrimaryKey]=[Old].[" & primaryKeyField & "] And "
220               xml = xml & "[tblLongTextBackup].[FieldName]=""" & FieldName & """"
230               xml = xml & "</WhereCondition>"
240               xml = xml & "</Data>"
250               xml = xml & "<Statements>"
260           End If
              
270           xml = xml & "<CreateRecord>"
280           If isLongText Then
290               xml = xml & "<Data><Reference>tblAuditLog</Reference></Data>"
300           Else
310               xml = xml & "<Data Alias=""NewAudit""><Reference>tblAuditLog</Reference></Data>"
320           End If
330           xml = xml & "<Statements>"
              
              ' TableName
340           xml = xml & "<Action Name=""SetField"">"
350           If isLongText Then
360               xml = xml & "<Argument Name=""Field"">tblAuditLog.TableName</Argument>"
370           Else
380               xml = xml & "<Argument Name=""Field"">NewAudit.TableName</Argument>"
390           End If
400           xml = xml & "<Argument Name=""Value"">""" & tableName & """</Argument>"
410           xml = xml & "</Action>"
              
              ' PrimaryKey
420           If primaryKeyField <> "" Then
430               xml = xml & "<Action Name=""SetField"">"
440               If isLongText Then
450                   xml = xml & "<Argument Name=""Field"">tblAuditLog.PrimaryKey</Argument>"
460               Else
470                   xml = xml & "<Argument Name=""Field"">NewAudit.PrimaryKey</Argument>"
480               End If
490               xml = xml & "<Argument Name=""Value"">[Old].[" & primaryKeyField & "]</Argument>"
500               xml = xml & "</Action>"
510           End If
              
              ' FieldName
520           xml = xml & "<Action Name=""SetField"">"
530           If isLongText Then
540               xml = xml & "<Argument Name=""Field"">tblAuditLog.FieldName</Argument>"
550           Else
560               xml = xml & "<Argument Name=""Field"">NewAudit.FieldName</Argument>"
570           End If
580           xml = xml & "<Argument Name=""Value"">""" & FieldName & """</Argument>"
590           xml = xml & "</Action>"
              
              ' OldValue
600           xml = xml & "<Action Name=""SetField"">"
610           If isLongText Then
620               xml = xml & "<Argument Name=""Field"">tblAuditLog.OldValue</Argument>"
630               xml = xml & "<Argument Name=""Value"">[BackupRec].[OldValue]</Argument>"
640           Else
650               xml = xml & "<Argument Name=""Field"">NewAudit.OldValue</Argument>"
660               xml = xml & "<Argument Name=""Value"">[Old].[" & FieldName & "]</Argument>"
670           End If
680           xml = xml & "</Action>"
              
              ' NewValue
690           xml = xml & "<Action Name=""SetField"">"
700           If isLongText Then
710               xml = xml & "<Argument Name=""Field"">tblAuditLog.NewValue</Argument>"
720           Else
730               xml = xml & "<Argument Name=""Field"">NewAudit.NewValue</Argument>"
740           End If
750           xml = xml & "<Argument Name=""Value"">""[DELETED]""</Argument>"
760           xml = xml & "</Action>"
              
              ' DateChanged
770           xml = xml & "<Action Name=""SetField"">"
780           If isLongText Then
790               xml = xml & "<Argument Name=""Field"">tblAuditLog.DateChanged</Argument>"
800           Else
810               xml = xml & "<Argument Name=""Field"">NewAudit.DateChanged</Argument>"
820           End If
830           xml = xml & "<Argument Name=""Value"">Now()</Argument>"
840           xml = xml & "</Action>"
              
              ' ChangedBy
850           xml = xml & "<Action Name=""SetField"">"
860           If isLongText Then
870               xml = xml & "<Argument Name=""Field"">tblAuditLog.ChangedBy</Argument>"
880           Else
890               xml = xml & "<Argument Name=""Field"">NewAudit.ChangedBy</Argument>"
900           End If
910           xml = xml & "<Argument Name=""Value"">CurrentUser()</Argument>"
920           xml = xml & "</Action>"
              
930           xml = xml & "</Statements></CreateRecord>"
              
              ' Close LookUpRecord if Long Text
940           If isLongText Then
950               xml = xml & "</Statements></LookUpRecord>"
960           End If
970       Next fieldInfo
          
980       xml = xml & "</Statements></DataMacro>"
          
990       BuildAfterDeleteMacro = xml
End Function

'-----------------------------------------------------------------------------
' Build Before Change Macro XML (backup Long Text fields before updates)
' Called by CreateAllDataMacros() for tables with Long Text fields
'-----------------------------------------------------------------------------
Private Function BuildBeforeChangeMacro(tableName As String, fieldList As Collection, primaryKeyField As String) As String
      Dim xml As String
      Dim fieldInfo As Variant
      Dim FieldName As String
      Dim fldType As Long
      Dim hasLongText As Boolean
          
          ' Check if there are any Long Text fields
100       hasLongText = False
110       For Each fieldInfo In fieldList
120           fldType = fieldInfo(1)
130           If fldType = dbMemo Then
140               hasLongText = True
150               Exit For
160           End If
170       Next fieldInfo
          
          ' Only create macro if there are Long Text fields
180       If Not hasLongText Then
190           BuildBeforeChangeMacro = ""
200           Exit Function
210       End If
          
          '    xml = "<?xml version=""1.0"" encoding=""UTF-16"" standalone=""no""?>"
          '    xml = xml & "<DataMacros xmlns=""http://schemas.microsoft.com/office/accessservices/2009/11/application"">"
          '    xml = xml & "<DataMacro Event=""BeforeChange""><Statements>"
220       xml = "<DataMacro Event=""BeforeChange""><Statements>"
          
          ' Add conditional block to check if this is a new record (PK is null) or update (PK has value)
230       xml = xml & "<ConditionalBlock><If>"
240       xml = xml & "<Condition>IsNull([Old].[" & primaryKeyField & "])</Condition>"
250       xml = xml & "<Statements>"
          ' If new record, set lngPkValue to 0
260       xml = xml & "<Action Name=""SetLocalVar"">"
270       xml = xml & "<Argument Name=""Name"">lngPkValue</Argument>"
280       xml = xml & "<Argument Name=""Value"">0</Argument>"
290       xml = xml & "</Action>"
300       xml = xml & "</Statements></If>"
          
          ' Else block for updates
310       xml = xml & "<Else><Statements>"
          
          ' Set PK value
320       xml = xml & "<Action Name=""SetLocalVar"">"
330       xml = xml & "<Argument Name=""Name"">lngPKValue</Argument>"
340       xml = xml & "<Argument Name=""Value"">=[" & primaryKeyField & "]</Argument>"
350       xml = xml & "</Action>"
          
          ' Set table name
360       xml = xml & "<Action Name=""SetLocalVar"">"
370       xml = xml & "<Argument Name=""Name"">strtableName</Argument>"
380       xml = xml & "<Argument Name=""Value"">""" & tableName & """</Argument>"
390       xml = xml & "</Action>"
          
          ' Add SetLocalVar for each Long Text field
400       For Each fieldInfo In fieldList
410           FieldName = fieldInfo(0)
420           fldType = fieldInfo(1)
              
430           If fldType = dbMemo Then
440               xml = xml & "<Action Name=""SetLocalVar"">"
450               xml = xml & "<Argument Name=""Name"">varLongTextBackup</Argument>"
460               xml = xml & "<Argument Name=""Value"">BackupLongTextFieldsDM([strTableName],[lngPKValue],""" & FieldName & """)</Argument>"
470               xml = xml & "</Action>"
480           End If
490       Next fieldInfo
          
500       xml = xml & "</Statements></Else></ConditionalBlock>"
510       xml = xml & "</Statements></DataMacro>"
          '    xml = xml & "</DataMacros>"
          
520       BuildBeforeChangeMacro = xml
End Function

'-----------------------------------------------------------------------------
' Build Before Delete Macro XML (backup Long Text fields before deletes)
' Called by CreateAllDataMacros() for tables with Long Text fields
'-----------------------------------------------------------------------------
Private Function BuildBeforeDeleteMacro(tableName As String, fieldList As Collection, primaryKeyField As String) As String
      Dim xml As String
      Dim fieldInfo As Variant
      Dim FieldName As String
      Dim fldType As Long
      Dim hasLongText As Boolean
          
          ' Check if there are any Long Text fields
100       hasLongText = False
110       For Each fieldInfo In fieldList
120           fldType = fieldInfo(1)
130           If fldType = dbMemo Then
140               hasLongText = True
150               Exit For
160           End If
170       Next fieldInfo
          
          ' Only create macro if there are Long Text fields
180       If Not hasLongText Then
190           BuildBeforeDeleteMacro = ""
200           Exit Function
210       End If
          
          '    xml = "<?xml version=""1.0"" encoding=""UTF-16"" standalone=""no""?>"
          '    xml = xml & "<DataMacros xmlns=""http://schemas.microsoft.com/office/accessservices/2009/11/application"">"
          '    xml = xml & "<DataMacro Event=""BeforeDelete""><Statements>"
220       xml = "<DataMacro Event=""BeforeDelete""><Statements>"
          ' Set PK value
230       xml = xml & "<Action Name=""SetLocalVar"">"
240       xml = xml & "<Argument Name=""Name"">lngPKValue</Argument>"
250       xml = xml & "<Argument Name=""Value"">=[" & primaryKeyField & "]</Argument>"
260       xml = xml & "</Action>"
          
          ' Set table name
270       xml = xml & "<Action Name=""SetLocalVar"">"
280       xml = xml & "<Argument Name=""Name"">strtableName</Argument>"
290       xml = xml & "<Argument Name=""Value"">""" & tableName & """</Argument>"
300       xml = xml & "</Action>"
          
          ' Add SetLocalVar for each Long Text field
310       For Each fieldInfo In fieldList
320           FieldName = fieldInfo(0)
330           fldType = fieldInfo(1)
              
340           If fldType = dbMemo Then
350               xml = xml & "<Action Name=""SetLocalVar"">"
360               xml = xml & "<Argument Name=""Name"">varLongTextBackup</Argument>"
370               xml = xml & "<Argument Name=""Value"">BackupLongTextFieldsDM([strTableName],[lngPKValue],""" & FieldName & """)</Argument>"
380               xml = xml & "</Action>"
390           End If
400       Next fieldInfo
          
410       xml = xml & "</Statements></DataMacro>"
          '    xml = xml & "</DataMacros>"
          
420       BuildBeforeDeleteMacro = xml

End Function

'-----------------------------------------------------------------------------
' Helper: Get comparison expression based on field type
'-----------------------------------------------------------------------------
Private Function GetComparisonExpression(tableName As String, FieldName As String, fldType As Long) As String
100       Select Case fldType
              Case dbMemo
                  ' Long Text: always log (can't compare old value)
110               GetComparisonExpression = "True"
120           Case Else
                  ' All other types: standard comparison
130               GetComparisonExpression = "StrComp(NZ([" & tableName & "].[" & FieldName & "],""""),NZ([Old].[" & FieldName & "],""""),0)&lt;&gt;0"
140       End Select
End Function

'-----------------------------------------------------------------------------
' Helper: Get primary key field name for a table
'-----------------------------------------------------------------------------
Private Function GetPrimaryKeyField(tableName As String) As String
      Dim db As DAO.Database
      Dim tdf As DAO.TableDef
      Dim idx As DAO.index
      Dim fld As DAO.Field
          
100       Set db = CurrentDb
110       Set tdf = db.TableDefs(tableName)
          
          ' Find primary key index
120       For Each idx In tdf.Indexes
130           If idx.Primary Then
140               For Each fld In idx.Fields
150                   GetPrimaryKeyField = fld.Name
160                   Exit Function
170               Next fld
180           End If
190       Next idx
          
200       GetPrimaryKeyField = ""
End Function




