Attribute VB_Name = "modDataMacros"
Option Compare Database
Option Explicit

'=============================================================================
' COMPLETE AUDIT DATA MACRO SYSTEM
'
' This module provides a complete auditing solution using Access Data Macros
' with special handling for Long Text fields via LookupRecord.
'
' SETUP INSTRUCTIONS:
' 1. Run CreateAuditTables() to create required tables
' 2. Run PopulateConfigTable() to populate the config with your tables/fields
' 3. Run GenerateAllAuditDataMacros() to create all Data Macros
' 4. Implement Form VBA (see separate instructions) for Long Text field backup
'=============================================================================

'-----------------------------------------------------------------------------
' STEP 1: Create all required tables
'-----------------------------------------------------------------------------
Public Sub CreateAuditTables()

100       On Error GoTo errHandler

      Dim db As DAO.Database
      Dim tdf As DAO.TableDef
      Dim fld As DAO.Field
      Dim idx As DAO.index
          
110       Set db = CurrentDb
          
          ' ========== Create tblAuditLog ==========
120       On Error Resume Next
130       Set tdf = db.TableDefs("tblAuditLog")
140       If Not tdf Is Nothing Then
150   Debug.Print "tblAuditLog already exists"
160           GoTo CreateLongTextBackup
170       End If
180       On Error GoTo errHandler
          
190       Set tdf = db.CreateTableDef("tblAuditLog")
          
200       Set fld = tdf.CreateField("AuditLogID", dbLong)
210       fld.Attributes = dbAutoIncrField
220       tdf.Fields.Append fld
          
230       Set fld = tdf.CreateField("TableName", dbText, 50)
240       fld.Required = True
250       tdf.Fields.Append fld
          
260       Set fld = tdf.CreateField("PrimaryKey", dbText, 50)
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
490       Set tdf = Nothing
500       On Error Resume Next
510       Set tdf = db.TableDefs("tblLongTextBackup")
520       If Not tdf Is Nothing Then
530   Debug.Print "tblLongTextBackup already exists"
540           GoTo CreateConfig
550       End If
560       On Error GoTo errHandler
          
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
          ' ========== Create tblDataMacroConfig ==========
860       Set tdf = Nothing
870       On Error Resume Next
880       Set tdf = db.TableDefs("tblDataMacroConfig")
890       If Not tdf Is Nothing Then
900   Debug.Print "tblDataMacroConfig already exists"
910           GoTo Cleanup
920       End If
930       On Error GoTo errHandler
          
940       Set tdf = db.CreateTableDef("tblDataMacroConfig")
          
950       Set fld = tdf.CreateField("ConfigID", dbLong)
960       fld.Attributes = dbAutoIncrField
970       tdf.Fields.Append fld
          
980       Set fld = tdf.CreateField("TableName", dbText, 50)
990       fld.Required = True
1000      tdf.Fields.Append fld
          
1010      Set fld = tdf.CreateField("FieldName", dbText, 50)
1020      fld.Required = True
1030      tdf.Fields.Append fld
          
1040      Set fld = tdf.CreateField("DataType", dbLong)
1050      fld.Required = True
1060      tdf.Fields.Append fld
          
1070      Set fld = tdf.CreateField("IsPrimaryKey", dbBoolean)
1080      fld.Required = True
1090      fld.DefaultValue = "False"
1100      tdf.Fields.Append fld
          
1110      db.TableDefs.Append tdf
          
1120      Set idx = tdf.CreateIndex("PrimaryKey")
1130      idx.Primary = True
1140      idx.Required = True
1150      Set fld = idx.CreateField("ConfigID")
1160      idx.Fields.Append fld
1170      tdf.Indexes.Append idx
          
1180  Debug.Print "tblDataMacroConfig created"
          
Cleanup:

1190      On Error Resume Next
1200      Set fld = Nothing
1210      Set idx = Nothing
1220      Set tdf = Nothing
1230      Set db = Nothing
1240      MsgBox "Audit tables created successfully!", vbInformation
1250      Exit Sub

errHandler:

1260      Call GlblErrMsg( _
              sFrm:=Application.VBE.ActiveCodePane.CodeModule, _
              sCtl:="CreateAuditTables")
1270      Resume Cleanup

1280      Resume
End Sub

'-----------------------------------------------------------------------------
' STEP 2: Populate configuration table with your tables and fields
'-----------------------------------------------------------------------------
Public Sub PopulateConfigTable()
    On Error GoTo errHandler

Dim db As DAO.Database
Dim tdef As DAO.TableDef
Dim fld As DAO.Field
Dim idx As DAO.index
Dim pkField As DAO.Field
Dim strSQL As String
Dim isPK As Boolean
Dim pkFieldName As String
    
    Set db = CurrentDb
    
    ' Clear existing config
    db.Execute "DELETE * FROM tblDataMacroConfig"
    
    ' Loop through all tables
    For Each tdef In db.TableDefs
        ' Filter: Include only tables starting with "tbl"
        ' Exclude: System tables, audit tables, PUTT tables
        If Left(tdef.Name, 3) = "tbl" _
            And tdef.Name <> "tblAuditLog" _
            And tdef.Name <> "tblLongTextBackup" _
            And tdef.Name <> "tblDataMacroConfig" _
            And tdef.Name <> "tblLoadTime" _
            And tdef.Name <> "tblPublicationHistory" _
            And Left(tdef.Name, 7) <> "tblPUTT" Then
            
            ' Get primary key field name for this table
            pkFieldName = ""
            For Each idx In tdef.Indexes
                If idx.Primary Then
                    For Each pkField In idx.Fields
                        pkFieldName = pkField.Name
                        Exit For
                    Next pkField
                    Exit For
                End If
            Next idx
            
            ' Add each field (excluding specific field names)
            For Each fld In tdef.Fields
                If fld.Name <> "AccessTS" _
                    And fld.Name <> "SSMA_TimeStamp" _
                    And fld.Name <> "ValidFrom" _
                    And fld.Name <> "ValidTo" Then
                    
                    ' Check if this field is the primary key
                    isPK = (fld.Name = pkFieldName)
                    
                    strSQL = "INSERT INTO tblDataMacroConfig (TableName, FieldName, DataType, IsPrimaryKey) " & _
                        "VALUES ('" & tdef.Name & "', '" & fld.Name & "', " & fld.type & ", " & isPK & ")"
                    db.Execute strSQL
                End If
            Next fld
        End If
    Next tdef
    
    Set pkField = Nothing
    Set idx = Nothing
    Set fld = Nothing
    Set tdef = Nothing
    Set db = Nothing
    
    MsgBox "Configuration table populated successfully!", vbInformation

Cleanup:
    
    On Error Resume Next
    Exit Sub

errHandler:
    Call GlblErrMsg( _
        sFrm:=Application.VBE.ActiveCodePane.CodeModule, _
        sCtl:="PopulateConfigTable")
    Resume Cleanup

    Resume
End Sub

'-----------------------------------------------------------------------------
' STEP 3: Generate all Data Macros for all tables
'-----------------------------------------------------------------------------
Public Sub GenerateAllAuditDataMacros()
       
100       On Error GoTo errHandler

          
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
          
          
410       MsgBox "Successfully generated audit data macros for " & tableCount & " tables!", vbInformation
Cleanup:
          
420       On Error Resume Next
         
430       Set rs = Nothing
440       Set db = Nothing
450       Set dictTables = Nothing
460       Exit Sub

errHandler:
470       Call GlblErrMsg( _
              sFrm:=Application.VBE.ActiveCodePane.CodeModule, _
              sCtl:="GenerateAllAuditDataMacros")
480       Resume Cleanup

490       Resume
End Sub

'-----------------------------------------------------------------------------
' Generate all three Data Macros in a single XML file
'-----------------------------------------------------------------------------
Private Sub CreateAllDataMacros(tableName As String, fieldList As Collection, tempPath As String)
          
100       On Error GoTo errHandler

      Dim xmlContent As String
      Dim fso As Object
      Dim txtFile As Object
      Dim filePath As String
      Dim primaryKeyField As String
      Dim fieldInfo As Variant
          
          ' Get primary key field from config
110       For Each fieldInfo In fieldList
120           If fieldInfo(2) = True Then ' IsPrimaryKey
130               primaryKeyField = fieldInfo(0) ' FieldName
140               Exit For
150           End If
160       Next fieldInfo
          
          ' Start XML
170       xmlContent = "<?xml version=""1.0"" encoding=""UTF-16"" standalone=""no""?>"
180       xmlContent = xmlContent & "<DataMacros xmlns=""http://schemas.microsoft.com/office/accessservices/2010/12/application"">"
          
          ' ========== AFTER INSERT MACRO ==========
190       xmlContent = xmlContent & BuildAfterInsertMacro(tableName, fieldList, primaryKeyField)
          
          ' ========== AFTER UPDATE MACRO ==========
200       xmlContent = xmlContent & BuildAfterUpdateMacro(tableName, fieldList, primaryKeyField)
          
          ' ========== AFTER DELETE MACRO ==========
210       xmlContent = xmlContent & BuildAfterDeleteMacro(tableName, fieldList, primaryKeyField)
          
          ' Close root element
220       xmlContent = xmlContent & "</DataMacros>"
          
          ' Write to file and load
230       filePath = tempPath & tableName & "_DataMacros.xml"
240       Set fso = CreateObject("Scripting.FileSystemObject")
250       Set txtFile = fso.CreateTextFile(filePath, True, True)
260       txtFile.Write xmlContent
270       txtFile.Close
          
280       Application.LoadFromText acTableDataMacro, tableName, filePath
290       fso.DeleteFile filePath
          
300       Set txtFile = Nothing
310       Set fso = Nothing
          
320   Debug.Print "  - All data macros created (Insert, Update, Delete)"

Cleanup:
          
330       On Error Resume Next
340       Exit Sub

errHandler:
350       Call GlblErrMsg( _
              sFrm:=Application.VBE.ActiveCodePane.CodeModule, _
              sCtl:="CreateAllDataMacros")
360       Resume Cleanup

370       Resume
End Sub

'-----------------------------------------------------------------------------
' Build After Insert Macro XML
'-----------------------------------------------------------------------------
Private Function BuildAfterInsertMacro(tableName As String, fieldList As Collection, primaryKeyField As String) As String
    On Error GoTo errHandler

Dim xml As String
Dim fieldInfo As Variant
Dim FieldName As String
    
    xml = "<DataMacro Event=""AfterInsert""><Statements>"
    
    For Each fieldInfo In fieldList
        FieldName = fieldInfo(0) ' FieldName from array
        
        xml = xml & "<CreateRecord>"
        xml = xml & "<Data Alias=""NewAudit""><Reference>tblAuditLog</Reference></Data>"
        xml = xml & "<Statements>"
        
        ' TableName
        xml = xml & "<Action Name=""SetField"">"
        xml = xml & "<Argument Name=""Field"">NewAudit.TableName</Argument>"
        xml = xml & "<Argument Name=""Value"">""" & tableName & """</Argument>"
        xml = xml & "</Action>"
        
        ' PrimaryKey
        If primaryKeyField <> "" Then
            xml = xml & "<Action Name=""SetField"">"
            xml = xml & "<Argument Name=""Field"">NewAudit.PrimaryKey</Argument>"
            xml = xml & "<Argument Name=""Value"">[" & tableName & "].[" & primaryKeyField & "]</Argument>"
            xml = xml & "</Action>"
        End If
        
        ' FieldName
        xml = xml & "<Action Name=""SetField"">"
        xml = xml & "<Argument Name=""Field"">NewAudit.FieldName</Argument>"
        xml = xml & "<Argument Name=""Value"">""" & FieldName & """</Argument>"
        xml = xml & "</Action>"
        
        ' OldValue
        xml = xml & "<Action Name=""SetField"">"
        xml = xml & "<Argument Name=""Field"">NewAudit.OldValue</Argument>"
        xml = xml & "<Argument Name=""Value"">""[NEW RECORD]""</Argument>"
        xml = xml & "</Action>"
        
        ' NewValue
        xml = xml & "<Action Name=""SetField"">"
        xml = xml & "<Argument Name=""Field"">NewAudit.NewValue</Argument>"
        xml = xml & "<Argument Name=""Value"">[" & tableName & "].[" & FieldName & "]</Argument>"
        xml = xml & "</Action>"
        
        ' DateChanged
        xml = xml & "<Action Name=""SetField"">"
        xml = xml & "<Argument Name=""Field"">NewAudit.DateChanged</Argument>"
        xml = xml & "<Argument Name=""Value"">Now()</Argument>"
        xml = xml & "</Action>"
        
        ' ChangedBy
        xml = xml & "<Action Name=""SetField"">"
        xml = xml & "<Argument Name=""Field"">NewAudit.ChangedBy</Argument>"
        xml = xml & "<Argument Name=""Value"">CurrentUser()</Argument>"
        xml = xml & "</Action>"
        
        xml = xml & "</Statements></CreateRecord>"
    Next fieldInfo 'fieldName
    
    xml = xml & "</Statements></DataMacro>"
    
    BuildAfterInsertMacro = xml
    
Cleanup:
    
    On Error Resume Next
    Exit Function

errHandler:
    Call GlblErrMsg( _
        sFrm:=Application.VBE.ActiveCodePane.CodeModule, _
        sCtl:="BuildAfterInsertMacro")
    Resume Cleanup

    Resume
End Function

'-----------------------------------------------------------------------------
' Build After Update Macro XML (with LookupRecord for Long Text)
'-----------------------------------------------------------------------------
Private Function BuildAfterUpdateMacro(tableName As String, fieldList As Collection, primaryKeyField As String) As String
100       On Error GoTo errHandler

      Dim xml As String
      Dim fieldInfo As Variant
      Dim FieldName As String 'Variant
      Dim fldType As Long
      Dim isLongText As Boolean
          
110       xml = "<DataMacro Event=""AfterUpdate""><Statements>"
          
120       For Each fieldInfo In fieldList
130           FieldName = fieldInfo(0) ' FieldName
140           fldType = fieldInfo(1)   ' DataType
150           isLongText = (fldType = dbMemo)
              
              ' Skip primary key
160           If FieldName <> primaryKeyField Then
170               xml = xml & "<ConditionalBlock><If>"
180               xml = xml & "<Condition>" & GetComparisonExpression(tableName, FieldName, fldType) & "</Condition>"
190               xml = xml & "<Statements>"
                  
                  ' If Long Text, use LookUpRecord to get old value
200               If isLongText Then
210                   xml = xml & "<LookUpRecord>"
220                   xml = xml & "<Data Alias=""BackupRec"">"
230                   xml = xml & "<Reference>tblLongTextBackup</Reference>"
240                   xml = xml & "<WhereCondition>"
250                   xml = xml & "[tblLongTextBackup].[TableName]=""" & tableName & """ And "
260                   xml = xml & "[tblLongTextBackup].[PrimaryKey]=[" & tableName & "].[" & primaryKeyField & "] And "
270                   xml = xml & "[tblLongTextBackup].[FieldName]=""" & FieldName & """"
280                   xml = xml & "</WhereCondition>"
290                   xml = xml & "</Data>"
300                   xml = xml & "<Statements>"
310               End If
                  
                  ' Create audit record
320               xml = xml & "<CreateRecord>"
330               If isLongText Then
340                   xml = xml & "<Data><Reference>tblAuditLog</Reference></Data>"
350               Else
360                   xml = xml & "<Data Alias=""NewAudit""><Reference>tblAuditLog</Reference></Data>"
370               End If
380               xml = xml & "<Statements>"
                  
                  ' TableName
390               xml = xml & "<Action Name=""SetField"">"
400               If isLongText Then
410                   xml = xml & "<Argument Name=""Field"">tblAuditLog.TableName</Argument>"
420               Else
430                   xml = xml & "<Argument Name=""Field"">NewAudit.TableName</Argument>"
440               End If
450               xml = xml & "<Argument Name=""Value"">""" & tableName & """</Argument>"
460               xml = xml & "</Action>"
                  
                  ' PrimaryKey
470               If primaryKeyField <> "" Then
480                   xml = xml & "<Action Name=""SetField"">"
490                   If isLongText Then
500                       xml = xml & "<Argument Name=""Field"">tblAuditLog.PrimaryKey</Argument>"
510                   Else
520                       xml = xml & "<Argument Name=""Field"">NewAudit.PrimaryKey</Argument>"
530                   End If
540                   xml = xml & "<Argument Name=""Value"">[" & tableName & "].[" & primaryKeyField & "]</Argument>"
550                   xml = xml & "</Action>"
560               End If
                  
                  ' FieldName
570               xml = xml & "<Action Name=""SetField"">"
580               If isLongText Then
590                   xml = xml & "<Argument Name=""Field"">tblAuditLog.FieldName</Argument>"
600               Else
610                   xml = xml & "<Argument Name=""Field"">NewAudit.FieldName</Argument>"
620               End If
630               xml = xml & "<Argument Name=""Value"">""" & FieldName & """</Argument>"
640               xml = xml & "</Action>"
                  
                  ' OldValue - from LookUpRecord if Long Text, else from [Old]
650               xml = xml & "<Action Name=""SetField"">"
660               If isLongText Then
670                   xml = xml & "<Argument Name=""Field"">tblAuditLog.OldValue</Argument>"
680                   xml = xml & "<Argument Name=""Value"">[BackupRec].[OldValue]</Argument>"
690               Else
700                   xml = xml & "<Argument Name=""Field"">NewAudit.OldValue</Argument>"
710                   xml = xml & "<Argument Name=""Value"">[Old].[" & FieldName & "]</Argument>"
720               End If
730               xml = xml & "</Action>"
                  
                  ' NewValue
740               xml = xml & "<Action Name=""SetField"">"
750               If isLongText Then
760                   xml = xml & "<Argument Name=""Field"">tblAuditLog.NewValue</Argument>"
770               Else
780                   xml = xml & "<Argument Name=""Field"">NewAudit.NewValue</Argument>"
790               End If
800               xml = xml & "<Argument Name=""Value"">[" & tableName & "].[" & FieldName & "]</Argument>"
810               xml = xml & "</Action>"
                  
                  ' DateChanged
820               xml = xml & "<Action Name=""SetField"">"
830               If isLongText Then
840                   xml = xml & "<Argument Name=""Field"">tblAuditLog.DateChanged</Argument>"
850               Else
860                   xml = xml & "<Argument Name=""Field"">NewAudit.DateChanged</Argument>"
870               End If
880               xml = xml & "<Argument Name=""Value"">Now()</Argument>"
890               xml = xml & "</Action>"
                  
                  ' ChangedBy
900               xml = xml & "<Action Name=""SetField"">"
910               If isLongText Then
920                   xml = xml & "<Argument Name=""Field"">tblAuditLog.ChangedBy</Argument>"
930               Else
940                   xml = xml & "<Argument Name=""Field"">NewAudit.ChangedBy</Argument>"
950               End If
960               xml = xml & "<Argument Name=""Value"">CurrentUser()</Argument>"
970               xml = xml & "</Action>"
                  
980               xml = xml & "</Statements></CreateRecord>"
                  
                  ' Close LookUpRecord if Long Text
990               If isLongText Then
1000                  xml = xml & "</Statements></LookUpRecord>"
1010              End If
                  
1020              xml = xml & "</Statements></If></ConditionalBlock>"
1030          End If
1040      Next fieldInfo 'fieldName
          
1050      xml = xml & "</Statements></DataMacro>"
          
1060      BuildAfterUpdateMacro = xml
          
Cleanup:
          
1070      On Error Resume Next
1080      Exit Function

errHandler:
1090      Call GlblErrMsg( _
              sFrm:=Application.VBE.ActiveCodePane.CodeModule, _
              sCtl:="BuildAfterUpdateMacro")
1100      Resume Cleanup

1110      Resume
End Function

'-----------------------------------------------------------------------------
' Build After Delete Macro XML (with LookupRecord for Long Text)
'-----------------------------------------------------------------------------
Private Function BuildAfterDeleteMacro(tableName As String, fieldList As Collection, primaryKeyField As String) As String
100       On Error GoTo errHandler

      Dim xml As String
      Dim fieldInfo As Variant
      Dim FieldName As String
      Dim fldType As Long
      Dim isLongText As Boolean
          
110       xml = "<DataMacro Event=""AfterDelete""><Statements>"
          
120       For Each fieldInfo In fieldList
130           FieldName = fieldInfo(0) ' FieldName
140           fldType = fieldInfo(1)   ' DataType
150           isLongText = (fldType = dbMemo)
          
160           If isLongText Then
170               xml = xml & "<LookUpRecord>"
180               xml = xml & "<Data Alias=""BackupRec"">"
190               xml = xml & "<Reference>tblLongTextBackup</Reference>"
200               xml = xml & "<WhereCondition>"
210               xml = xml & "[tblLongTextBackup].[TableName]=""" & tableName & """ And "
220               xml = xml & "[tblLongTextBackup].[PrimaryKey]=[Old].[" & primaryKeyField & "] And "
230               xml = xml & "[tblLongTextBackup].[FieldName]=""" & FieldName & """"
240               xml = xml & "</WhereCondition>"
250               xml = xml & "</Data>"
260               xml = xml & "<Statements>"
270           End If
              
280           xml = xml & "<CreateRecord>"
290           If isLongText Then
300               xml = xml & "<Data><Reference>tblAuditLog</Reference></Data>"
310           Else
320               xml = xml & "<Data Alias=""NewAudit""><Reference>tblAuditLog</Reference></Data>"
330           End If
340           xml = xml & "<Statements>"
              
350           xml = xml & "<Action Name=""SetField"">"
360           If isLongText Then
370               xml = xml & "<Argument Name=""Field"">tblAuditLog.TableName</Argument>"
380           Else
390               xml = xml & "<Argument Name=""Field"">NewAudit.TableName</Argument>"
400           End If
410           xml = xml & "<Argument Name=""Value"">""" & tableName & """</Argument>"
420           xml = xml & "</Action>"
              
430           If primaryKeyField <> "" Then
440               xml = xml & "<Action Name=""SetField"">"
450               If isLongText Then
460                   xml = xml & "<Argument Name=""Field"">tblAuditLog.PrimaryKey</Argument>"
470               Else
480                   xml = xml & "<Argument Name=""Field"">NewAudit.PrimaryKey</Argument>"
490               End If
500               xml = xml & "<Argument Name=""Value"">[Old].[" & primaryKeyField & "]</Argument>"
510               xml = xml & "</Action>"
520           End If
              
530           xml = xml & "<Action Name=""SetField"">"
540           If isLongText Then
550               xml = xml & "<Argument Name=""Field"">tblAuditLog.FieldName</Argument>"
560           Else
570               xml = xml & "<Argument Name=""Field"">NewAudit.FieldName</Argument>"
580           End If
590           xml = xml & "<Argument Name=""Value"">""" & FieldName & """</Argument>"
600           xml = xml & "</Action>"
              
610           xml = xml & "<Action Name=""SetField"">"
620           If isLongText Then
630               xml = xml & "<Argument Name=""Field"">tblAuditLog.OldValue</Argument>"
640               xml = xml & "<Argument Name=""Value"">[BackupRec].[OldValue]</Argument>"
650           Else
660               xml = xml & "<Argument Name=""Field"">NewAudit.OldValue</Argument>"
670               xml = xml & "<Argument Name=""Value"">[Old].[" & FieldName & "]</Argument>"
680           End If
690           xml = xml & "</Action>"
              
700           xml = xml & "<Action Name=""SetField"">"
710           If isLongText Then
720               xml = xml & "<Argument Name=""Field"">tblAuditLog.NewValue</Argument>"
730           Else
740               xml = xml & "<Argument Name=""Field"">NewAudit.NewValue</Argument>"
750           End If
760           xml = xml & "<Argument Name=""Value"">""[DELETED]""</Argument>"
770           xml = xml & "</Action>"
              
780           xml = xml & "<Action Name=""SetField"">"
790           If isLongText Then
800               xml = xml & "<Argument Name=""Field"">tblAuditLog.DateChanged</Argument>"
810           Else
820               xml = xml & "<Argument Name=""Field"">NewAudit.DateChanged</Argument>"
830           End If
840           xml = xml & "<Argument Name=""Value"">Now()</Argument>"
850           xml = xml & "</Action>"
              
860           xml = xml & "<Action Name=""SetField"">"
870           If isLongText Then
880               xml = xml & "<Argument Name=""Field"">tblAuditLog.ChangedBy</Argument>"
890           Else
900               xml = xml & "<Argument Name=""Field"">NewAudit.ChangedBy</Argument>"
910           End If
920           xml = xml & "<Argument Name=""Value"">CurrentUser()</Argument>"
930           xml = xml & "</Action>"
              
940           xml = xml & "</Statements></CreateRecord>"
              
950           If isLongText Then
960               xml = xml & "</Statements></LookUpRecord>"
970           End If
980       Next fieldInfo
          
990       xml = xml & "</Statements></DataMacro>"
          
1000      BuildAfterDeleteMacro = xml
          
Cleanup:
          
1010      On Error Resume Next
1020      Exit Function

errHandler:
1030      Call GlblErrMsg( _
              sFrm:=Application.VBE.ActiveCodePane.CodeModule, _
              sCtl:="BuildAfterDeleteMacro")
1040      Resume Cleanup

1050      Resume
End Function

'-----------------------------------------------------------------------------
' Helper: Get comparison expression based on field type
'-----------------------------------------------------------------------------
Private Function GetComparisonExpression(tableName As String, FieldName As String, fldType As Long) As String
    Select Case fldType
        Case dbMemo
            ' Long Text: always log (can't compare old value)
            GetComparisonExpression = "True"
        Case Else
            ' All other types: standard comparison
            GetComparisonExpression = "StrComp(NZ([" & tableName & "].[" & FieldName & "],""""),NZ([Old].[" & FieldName & "],""""),0)&lt;&gt;0"
    End Select
End Function

'-----------------------------------------------------------------------------
' Helper: Get primary key field name for a table
'-----------------------------------------------------------------------------
Private Function GetPrimaryKeyField(tableName As String) As String

Dim db As DAO.Database
Dim tdf As DAO.TableDef
Dim idx As DAO.index
Dim fld As DAO.Field
    
    Set db = CurrentDb
    Set tdf = db.TableDefs(tableName)
    
    For Each idx In tdf.Indexes
        If idx.Primary Then
            For Each fld In idx.Fields
                GetPrimaryKeyField = fld.Name
                Exit Function
            Next fld
        End If
    Next idx
    
    GetPrimaryKeyField = ""
End Function

