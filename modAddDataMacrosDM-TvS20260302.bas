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
    On Error GoTo ErrorHandler

    Dim db                  As DAO.Database
    Dim tdf                 As DAO.TableDef
    Dim fld                 As DAO.Field
    Dim idx                 As DAO.Index

    Set db = CurrentDb

    ' ========== Create tblAuditLog ==========
    On Error Resume Next
    Set tdf = db.TableDefs("tblAuditLog")
    If Not tdf Is Nothing Then
        Debug.Print "tblAuditLog already exists"
        GoTo CreateLongTextBackup
    End If
    On Error GoTo ErrorHandler

    Set tdf = db.CreateTableDef("tblAuditLog")

    Set fld = tdf.CreateField("AuditLogID", dbLong)
    fld.Attributes = dbAutoIncrField
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("TableName", dbText, 50)
    fld.Required = True
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("PrimaryKey", dbLong)
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("FieldName", dbText, 50)
    fld.Required = True
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("OldValue", dbMemo)
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("NewValue", dbMemo)
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("DateChanged", dbDate)
    fld.Required = True
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("ChangedBy", dbText, 50)
    fld.Required = True
    tdf.Fields.Append fld

    db.TableDefs.Append tdf

    Set idx = tdf.CreateIndex("PrimaryKey")
    idx.Primary = True
    idx.Required = True
    Set fld = idx.CreateField("AuditLogID")
    idx.Fields.Append fld
    tdf.Indexes.Append idx

    Debug.Print "tblAuditLog created"

CreateLongTextBackup:
    ' ========== Create tblLongTextBackup ==========
    'Optional, useful only when Long Text field data is <= ~2034 characters long and can be edited by VBA

    Set tdf = Nothing
    On Error Resume Next
    Set tdf = db.TableDefs("tblLongTextBackup")
    If Not tdf Is Nothing Then
        Debug.Print "tblLongTextBackup already exists"
        GoTo CreateConfig
    End If
    On Error GoTo ErrorHandler

    Set tdf = db.CreateTableDef("tblLongTextBackup")

    Set fld = tdf.CreateField("BackupID", dbLong)
    fld.Attributes = dbAutoIncrField
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("TableName", dbText, 50)
    fld.Required = True
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("PrimaryKey", dbLong)
    fld.Required = True
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("FieldName", dbText, 50)
    fld.Required = True
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("OldValue", dbMemo)
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("DateChanged", dbDate)
    fld.Required = True
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("ChangedBy", dbText, 50)
    fld.Required = True
    tdf.Fields.Append fld

    db.TableDefs.Append tdf

    Set idx = tdf.CreateIndex("PrimaryKey")
    idx.Primary = True
    idx.Required = True
    Set fld = idx.CreateField("BackupID")
    idx.Fields.Append fld
    tdf.Indexes.Append idx

    Debug.Print "tblLongTextBackup created"

CreateConfig:
    ' ========== Create tblAuditLogConfig ==========
    Set tdf = Nothing
    On Error Resume Next
    Set tdf = db.TableDefs("tblAuditLogConfig")
    If Not tdf Is Nothing Then
        Debug.Print "tblAuditLogConfig already exists"
        GoTo Cleanup
    End If
    On Error GoTo ErrorHandler

    Set tdf = db.CreateTableDef("tblAuditLogConfig")

    Set fld = tdf.CreateField("ConfigID", dbLong)
    fld.Attributes = dbAutoIncrField
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("TableName", dbText, 50)
    fld.Required = True
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("FieldName", dbText, 50)
    fld.Required = True
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("FieldPosition", dbLong)
    fld.Required = True
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("DataType", dbLong)
    fld.Required = True
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("IsPrimaryKey", dbBoolean)
    fld.Required = True
    fld.DefaultValue = "False"
    tdf.Fields.Append fld

    Set fld = tdf.CreateField("IsAuditable", dbBoolean)
    fld.Required = True
    fld.DefaultValue = "True"
    tdf.Fields.Append fld
    db.TableDefs.Append tdf

    Set idx = tdf.CreateIndex("PrimaryKey")
    idx.Primary = True
    idx.Required = True
    Set fld = idx.CreateField("ConfigID")
    idx.Fields.Append fld
    tdf.Indexes.Append idx

    Debug.Print "tblAuditLogConfig created"

Cleanup:
    Set fld = Nothing
    Set idx = Nothing
    Set tdf = Nothing
    Set db = Nothing

    MsgBox "Audit tables created successfully!", vbInformation
    Exit Sub

ErrorHandler: 'If desired, replace with your own Error Handling
    MsgBox "Error creating tables: " & Err.Number & " - " & Err.Description, vbCritical
    Resume Cleanup
End Sub


'-----------------------------------------------------------------------------
' STEP 2: Populate configuration table with your tables and fields
'         Customize by including/excluding specific tables and fields in your database.
'-----------------------------------------------------------------------------
Public Sub Two_PopulateConfigTable()
    On Error GoTo ErrorHandler

    Dim db                  As DAO.Database
    Dim tdef                As DAO.TableDef
    Dim fld                 As DAO.Field
    Dim idx                 As DAO.Index
    Dim pkField             As DAO.Field
    Dim strSQL              As String
    Dim isPK                As Boolean
    Dim pkFieldName         As String

    Set db = CurrentDb

    ' Clear existing config
    db.Execute "DELETE * FROM tblAuditLogConfig", dbFailOnError

    ' Loop through all tables
    For Each tdef In db.TableDefs
        ' Filter: Include only tables starting with "tbl"
        ' Exclude: System tables, audit tables, other tables you don't want to audit
        'This should be implemented as a table driven solution, with all auditable tables and auditable fields flagged or unflagged as appropriate.
        'The audit log tables could also be prefixed "USys" to avoid hard-coding them here or flagging them in the config table.
        'If Left(tdef.Name, 3) = "tbl" Then
        If tdef.Name = "tblInventory" Then

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

            ' Add each field to be audited (excluding specific field names you don't want to audit)
            ' Instead of hard-coding excluded fields, this step should be modified to add all fields.
            ' Then, an additional "IsAudited" field in the config table can be used to
            '  review and manually flag excluded fields.
            For Each fld In tdef.Fields
                ' If fld.Name <> "AccessTS" _
                  '     And fld.Name <> "SSMA_TimeStamp" _
                  '     And fld.Name <> "ValidFrom" _
                  '     And fld.Name <> "ValidTo" Then

                ' Check if this field is the primary key
                isPK = (fld.Name = pkFieldName)

                strSQL = "INSERT INTO tblAuditLogConfig (TableName, FieldName, FieldPosition, DataType, IsPrimaryKey, IsAuditable) " & _
                         "VALUES ('" & tdef.Name & "', '" & fld.Name & "', " & fld.OrdinalPosition & "," & fld.Type & ", " & isPK & ", " & -1 & ")"     'default to IsAuditable = true
                db.Execute strSQL, dbFailOnError

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
    Exit Sub

ErrorHandler: 'If desired, replace with your own Error Handling
    MsgBox "Error populating config: " & Err.Number & " - " & Err.Description, vbCritical
    Resume Cleanup
    Resume
End Sub

'-----------------------------------------------------------------------------
' STEP 3: Generate 3 or 5 (tables containing Long Text fields) Data Macros for all tables
'-----------------------------------------------------------------------------
Public Sub Three_GenerateAllAuditDataMacros()
    On Error GoTo ErrorHandler

    Dim db                  As DAO.Database
    Dim rs                  As DAO.Recordset
    Dim dictTables          As Object
    Dim tableName           As String
    Dim FieldName           As String
    Dim fieldDataType       As Long
    Dim fieldIsPK           As Boolean
    Dim fieldList           As Collection
    Dim fieldInfo           As Variant
    Dim tempPath            As String
    Dim tableCount          As Long
    Dim currentTable        As Variant

    Set db = CurrentDb
    Set dictTables = CreateObject("Scripting.Dictionary")

    ' Read configuration and group by table
    Set rs = db.OpenRecordset("SELECT TableName, FieldName, DataType, IsPrimaryKey FROM tblAuditLogConfig ORDER BY TableName, FieldName", dbOpenSnapshot)

    Do While Not rs.EOF
        tableName = Nz(rs!tableName, "")
        FieldName = Nz(rs!FieldName, "")
        fieldDataType = Nz(rs!DataType, 0)
        fieldIsPK = Nz(rs!IsPrimaryKey, False)

        If tableName <> "" And FieldName <> "" Then
            If Not dictTables.Exists(tableName) Then
                Set fieldList = New Collection
                dictTables.Add tableName, fieldList
            Else
                Set fieldList = dictTables(tableName)
            End If

            ' Store field info as array: (FieldName, DataType, IsPrimaryKey)
            fieldInfo = Array(FieldName, fieldDataType, fieldIsPK)
            fieldList.Add fieldInfo
        End If

        rs.MoveNext
    Loop
    rs.Close

    tempPath = Environ("TEMP") & "\"

    ' Process each table
    tableCount = 0
    For Each currentTable In dictTables.Keys
        tableName = CStr(currentTable)
        Set fieldList = dictTables(tableName)

        Debug.Print "Processing: " & tableName & " (" & fieldList.Count & " fields)"

        Call CreateAllDataMacros(tableName, fieldList, tempPath)

        tableCount = tableCount + 1
    Next currentTable

    Set rs = Nothing
    Set db = Nothing
    Set dictTables = Nothing

    MsgBox "Successfully generated audit data macros for " & tableCount & " tables!", vbInformation
    Debug.Print "Completed: " & tableCount & " tables processed"

Cleanup:
    Exit Sub

ErrorHandler: 'If desired, replace with your own Error Handling
    MsgBox "Error: " & Err.Number & " - " & Err.Description, vbCritical
    Resume Cleanup
End Sub

'-----------------------------------------------------------------------------
' Generate 3 or 5 Data Macros in a single XML file
' Called by Three_GenerateAllAuditDataMacros()
' Generate Data Macros for a table (3 After macros + 2 Before macros if Long Text fields exist)
'-----------------------------------------------------------------------------
Private Sub CreateAllDataMacros(tableName As String, fieldList As Collection, tempPath As String)
    On Error GoTo ErrorHandler

    Dim xmlContent          As String
    Dim fso                 As Object
    Dim txtFile             As Object
    Dim filePath            As String
    Dim primaryKeyField     As String
    Dim fieldInfo           As Variant
    Dim hasLongText         As Boolean

    ' Get primary key field from config and check for Long Text fields
    hasLongText = False
    For Each fieldInfo In fieldList
        If fieldInfo(2) = True Then    ' IsPrimaryKey
            primaryKeyField = fieldInfo(0)    ' FieldName
        End If
        If fieldInfo(1) = dbMemo Then    ' Check for Long Text
            hasLongText = True
        End If
    Next fieldInfo

    ' ========== CREATE THE THREE OR FIVE MACROS IN ONE FILE ==========
    ' Start XML
    xmlContent = "<?xml version=""1.0"" encoding=""UTF-16"" standalone=""no""?>"
    xmlContent = xmlContent & "<DataMacros xmlns=""http://schemas.microsoft.com/office/accessservices/2010/12/application"">"

    ' AFTER INSERT MACRO
    xmlContent = xmlContent & BuildAfterInsertMacro(tableName, fieldList, primaryKeyField)

    ' AFTER UPDATE MACRO
    xmlContent = xmlContent & BuildAfterUpdateMacro(tableName, fieldList, primaryKeyField)

    ' AFTER DELETE MACRO
    xmlContent = xmlContent & BuildAfterDeleteMacro(tableName, fieldList, primaryKeyField)

    If hasLongText Then
        ' BEFORE CHANGE MACRO
        xmlContent = xmlContent & BuildBeforeChangeMacro(tableName, fieldList, primaryKeyField)

        ' BEFORE DELETE MACRO
        xmlContent = xmlContent & BuildBeforeDeleteMacro(tableName, fieldList, primaryKeyField)
    End If

    ' Close root element
    xmlContent = xmlContent & "</DataMacros>"

    ' Write to file and load
    filePath = tempPath & tableName & "_DataMacros.xml"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set txtFile = fso.CreateTextFile(filePath, True, True)
    txtFile.Write xmlContent
    txtFile.Close
    Set txtFile = Nothing

    ' Load the Data Macros
    DoCmd.OpenTable tableName, acViewDesign, acHidden
    Application.LoadFromText acTableDataMacro, tableName, filePath
    DoCmd.Close acTable, tableName, acSaveYes

    ' Delete temp file
    fso.DeleteFile filePath

Cleanup:
    Exit Sub

ErrorHandler: 'If desired, replace with your own Error Handling
    MsgBox "Error creating macros for " & tableName & ": " & Err.Number & " - " & Err.Description, vbCritical

    Resume Cleanup

End Sub

'-----------------------------------------------------------------------------
' Build After Insert Macro XML
' Called by Three_GenerateAllAuditDataMacros()
'-----------------------------------------------------------------------------
Private Function BuildAfterInsertMacro(tableName As String, fieldList As Collection, primaryKeyField As String) As String
    Dim xml                 As String
    Dim fieldInfo           As Variant
    Dim FieldName           As String

    xml = "<DataMacro Event=""AfterInsert""><Statements>"

    For Each fieldInfo In fieldList
        FieldName = fieldInfo(0)    ' FieldName from array

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
        xml = xml & "<Argument Name=""Value"">GetAuditFieldsUserName()</Argument>"
        xml = xml & "</Action>"

        xml = xml & "</Statements></CreateRecord>"
    Next fieldInfo    'fieldName

    xml = xml & "</Statements></DataMacro>"

    BuildAfterInsertMacro = xml
End Function

'-----------------------------------------------------------------------------
' Build After Update Macro XML (with LookupRecord for Long Text)
' Called by Three_GenerateAllAuditDataMacros()
'-----------------------------------------------------------------------------
Private Function BuildAfterUpdateMacro(tableName As String, fieldList As Collection, primaryKeyField As String) As String
    Dim xml                 As String
    Dim fieldInfo           As Variant
    Dim FieldName           As String    'Variant
    Dim fldType             As Long
    Dim isLongText          As Boolean

    xml = "<DataMacro Event=""AfterUpdate""><Statements>"

    For Each fieldInfo In fieldList
        FieldName = fieldInfo(0)    ' FieldName
        fldType = fieldInfo(1)   ' DataType
        isLongText = (fldType = dbMemo)

        ' Skip primary key
        If FieldName <> primaryKeyField Then
            xml = xml & "<ConditionalBlock><If>"
            xml = xml & "<Condition>" & GetComparisonExpression(tableName, FieldName, fldType) & "</Condition>"
            xml = xml & "<Statements>"

            ' If Long Text, use LookUpRecord to get old value
            If isLongText Then
                xml = xml & "<LookUpRecord>"
                xml = xml & "<Data Alias=""BackupRec"">"
                xml = xml & "<Reference>tblLongTextBackup</Reference>"
                xml = xml & "<WhereCondition>"
                xml = xml & "[tblLongTextBackup].[TableName]=""" & tableName & """ And "
                xml = xml & "[tblLongTextBackup].[PrimaryKey]=[" & tableName & "].[" & primaryKeyField & "] And "
                xml = xml & "[tblLongTextBackup].[FieldName]=""" & FieldName & """"
                xml = xml & "</WhereCondition>"
                xml = xml & "</Data>"
                xml = xml & "<Statements>"
            End If

            ' Create audit record
            xml = xml & "<CreateRecord>"
            If isLongText Then
                xml = xml & "<Data><Reference>tblAuditLog</Reference></Data>"
            Else
                xml = xml & "<Data Alias=""NewAudit""><Reference>tblAuditLog</Reference></Data>"
            End If
            xml = xml & "<Statements>"

            ' TableName
            xml = xml & "<Action Name=""SetField"">"
            If isLongText Then
                xml = xml & "<Argument Name=""Field"">tblAuditLog.TableName</Argument>"
            Else
                xml = xml & "<Argument Name=""Field"">NewAudit.TableName</Argument>"
            End If
            xml = xml & "<Argument Name=""Value"">""" & tableName & """</Argument>"
            xml = xml & "</Action>"

            ' PrimaryKey
            If primaryKeyField <> "" Then
                xml = xml & "<Action Name=""SetField"">"
                If isLongText Then
                    xml = xml & "<Argument Name=""Field"">tblAuditLog.PrimaryKey</Argument>"
                Else
                    xml = xml & "<Argument Name=""Field"">NewAudit.PrimaryKey</Argument>"
                End If
                xml = xml & "<Argument Name=""Value"">[" & tableName & "].[" & primaryKeyField & "]</Argument>"
                xml = xml & "</Action>"
            End If

            ' FieldName
            xml = xml & "<Action Name=""SetField"">"
            If isLongText Then
                xml = xml & "<Argument Name=""Field"">tblAuditLog.FieldName</Argument>"
            Else
                xml = xml & "<Argument Name=""Field"">NewAudit.FieldName</Argument>"
            End If
            xml = xml & "<Argument Name=""Value"">""" & FieldName & """</Argument>"
            xml = xml & "</Action>"

            ' OldValue - from LookUpRecord if Long Text, else from [Old]
            xml = xml & "<Action Name=""SetField"">"
            If isLongText Then
                xml = xml & "<Argument Name=""Field"">tblAuditLog.OldValue</Argument>"
                xml = xml & "<Argument Name=""Value"">[BackupRec].[OldValue]</Argument>"
            Else
                xml = xml & "<Argument Name=""Field"">NewAudit.OldValue</Argument>"
                xml = xml & "<Argument Name=""Value"">[Old].[" & FieldName & "]</Argument>"
            End If
            xml = xml & "</Action>"

            ' NewValue
            xml = xml & "<Action Name=""SetField"">"
            If isLongText Then
                xml = xml & "<Argument Name=""Field"">tblAuditLog.NewValue</Argument>"
            Else
                xml = xml & "<Argument Name=""Field"">NewAudit.NewValue</Argument>"
            End If
            xml = xml & "<Argument Name=""Value"">[" & tableName & "].[" & FieldName & "]</Argument>"
            xml = xml & "</Action>"

            ' DateChanged
            xml = xml & "<Action Name=""SetField"">"
            If isLongText Then
                xml = xml & "<Argument Name=""Field"">tblAuditLog.DateChanged</Argument>"
            Else
                xml = xml & "<Argument Name=""Field"">NewAudit.DateChanged</Argument>"
            End If
            xml = xml & "<Argument Name=""Value"">Now()</Argument>"
            xml = xml & "</Action>"

            ' ChangedBy
            xml = xml & "<Action Name=""SetField"">"
            If isLongText Then
                xml = xml & "<Argument Name=""Field"">tblAuditLog.ChangedBy</Argument>"
            Else
                xml = xml & "<Argument Name=""Field"">NewAudit.ChangedBy</Argument>"
            End If
            xml = xml & "<Argument Name=""Value"">GetAuditFieldsUserName()</Argument>"
            xml = xml & "</Action>"

            xml = xml & "</Statements></CreateRecord>"

            ' Close LookUpRecord if Long Text
            If isLongText Then
                xml = xml & "</Statements></LookUpRecord>"
            End If

            xml = xml & "</Statements></If></ConditionalBlock>"
        End If
    Next fieldInfo    'fieldName

    xml = xml & "</Statements></DataMacro>"

    BuildAfterUpdateMacro = xml
End Function

'-----------------------------------------------------------------------------
' Build After Delete Macro XML (with LookupRecord for Long Text)
' Called by Three_GenerateAllAuditDataMacros()
'-----------------------------------------------------------------------------
Private Function BuildAfterDeleteMacro(tableName As String, fieldList As Collection, primaryKeyField As String) As String
    Dim xml                 As String
    Dim fieldInfo           As Variant
    Dim FieldName           As String
    Dim fldType             As Long
    Dim isLongText          As Boolean

    xml = "<DataMacro Event=""AfterDelete""><Statements>"

    For Each fieldInfo In fieldList
        FieldName = fieldInfo(0)    ' FieldName
        fldType = fieldInfo(1)   ' DataType
        isLongText = (fldType = dbMemo)

        ' If Long Text, use LookUpRecord
        If isLongText Then
            xml = xml & "<LookUpRecord>"
            xml = xml & "<Data Alias=""BackupRec"">"
            xml = xml & "<Reference>tblLongTextBackup</Reference>"
            xml = xml & "<WhereCondition>"
            xml = xml & "[tblLongTextBackup].[TableName]=""" & tableName & """ And "
            xml = xml & "[tblLongTextBackup].[PrimaryKey]=[Old].[" & primaryKeyField & "] And "
            xml = xml & "[tblLongTextBackup].[FieldName]=""" & FieldName & """"
            xml = xml & "</WhereCondition>"
            xml = xml & "</Data>"
            xml = xml & "<Statements>"
        End If

        xml = xml & "<CreateRecord>"
        If isLongText Then
            xml = xml & "<Data><Reference>tblAuditLog</Reference></Data>"
        Else
            xml = xml & "<Data Alias=""NewAudit""><Reference>tblAuditLog</Reference></Data>"
        End If
        xml = xml & "<Statements>"

        ' TableName
        xml = xml & "<Action Name=""SetField"">"
        If isLongText Then
            xml = xml & "<Argument Name=""Field"">tblAuditLog.TableName</Argument>"
        Else
            xml = xml & "<Argument Name=""Field"">NewAudit.TableName</Argument>"
        End If
        xml = xml & "<Argument Name=""Value"">""" & tableName & """</Argument>"
        xml = xml & "</Action>"

        ' PrimaryKey
        If primaryKeyField <> "" Then
            xml = xml & "<Action Name=""SetField"">"
            If isLongText Then
                xml = xml & "<Argument Name=""Field"">tblAuditLog.PrimaryKey</Argument>"
            Else
                xml = xml & "<Argument Name=""Field"">NewAudit.PrimaryKey</Argument>"
            End If
            xml = xml & "<Argument Name=""Value"">[Old].[" & primaryKeyField & "]</Argument>"
            xml = xml & "</Action>"
        End If

        ' FieldName
        xml = xml & "<Action Name=""SetField"">"
        If isLongText Then
            xml = xml & "<Argument Name=""Field"">tblAuditLog.FieldName</Argument>"
        Else
            xml = xml & "<Argument Name=""Field"">NewAudit.FieldName</Argument>"
        End If
        xml = xml & "<Argument Name=""Value"">""" & FieldName & """</Argument>"
        xml = xml & "</Action>"

        ' OldValue
        xml = xml & "<Action Name=""SetField"">"
        If isLongText Then
            xml = xml & "<Argument Name=""Field"">tblAuditLog.OldValue</Argument>"
            xml = xml & "<Argument Name=""Value"">[BackupRec].[OldValue]</Argument>"
        Else
            xml = xml & "<Argument Name=""Field"">NewAudit.OldValue</Argument>"
            xml = xml & "<Argument Name=""Value"">[Old].[" & FieldName & "]</Argument>"
        End If
        xml = xml & "</Action>"

        ' NewValue
        xml = xml & "<Action Name=""SetField"">"
        If isLongText Then
            xml = xml & "<Argument Name=""Field"">tblAuditLog.NewValue</Argument>"
        Else
            xml = xml & "<Argument Name=""Field"">NewAudit.NewValue</Argument>"
        End If
        xml = xml & "<Argument Name=""Value"">""[DELETED]""</Argument>"
        xml = xml & "</Action>"

        ' DateChanged
        xml = xml & "<Action Name=""SetField"">"
        If isLongText Then
            xml = xml & "<Argument Name=""Field"">tblAuditLog.DateChanged</Argument>"
        Else
            xml = xml & "<Argument Name=""Field"">NewAudit.DateChanged</Argument>"
        End If
        xml = xml & "<Argument Name=""Value"">Now()</Argument>"
        xml = xml & "</Action>"

        ' ChangedBy
        xml = xml & "<Action Name=""SetField"">"
        If isLongText Then
            xml = xml & "<Argument Name=""Field"">tblAuditLog.ChangedBy</Argument>"
        Else
            xml = xml & "<Argument Name=""Field"">NewAudit.ChangedBy</Argument>"
        End If
        xml = xml & "<Argument Name=""Value"">GetAuditFieldsUserName()</Argument>"
        xml = xml & "</Action>"

        xml = xml & "</Statements></CreateRecord>"

        ' Close LookUpRecord if Long Text
        If isLongText Then
            xml = xml & "</Statements></LookUpRecord>"
        End If
    Next fieldInfo

    xml = xml & "</Statements></DataMacro>"

    BuildAfterDeleteMacro = xml
End Function

'-----------------------------------------------------------------------------
' Build Before Change Macro XML (backup Long Text fields before updates)
' Called by CreateAllDataMacros() for tables with Long Text fields
'-----------------------------------------------------------------------------
Private Function BuildBeforeChangeMacro(tableName As String, fieldList As Collection, primaryKeyField As String) As String
    Dim xml                 As String
    Dim fieldInfo           As Variant
    Dim FieldName           As String
    Dim fldType             As Long
    Dim hasLongText         As Boolean

    ' Check if there are any Long Text fields
    hasLongText = False
    For Each fieldInfo In fieldList
        fldType = fieldInfo(1)
        If fldType = dbMemo Then
            hasLongText = True
            Exit For
        End If
    Next fieldInfo

    ' Only create macro if there are Long Text fields
    If Not hasLongText Then
        BuildBeforeChangeMacro = ""
        Exit Function
    End If

    '    xml = "<?xml version=""1.0"" encoding=""UTF-16"" standalone=""no""?>"
    '    xml = xml & "<DataMacros xmlns=""http://schemas.microsoft.com/office/accessservices/2009/11/application"">"
    '    xml = xml & "<DataMacro Event=""BeforeChange""><Statements>"
    xml = "<DataMacro Event=""BeforeChange""><Statements>"

    ' Add conditional block to check if this is a new record (PK is null) or update (PK has value)
    xml = xml & "<ConditionalBlock><If>"
    xml = xml & "<Condition>IsNull([Old].[" & primaryKeyField & "])</Condition>"
    xml = xml & "<Statements>"
    ' If new record, set lngPkValue to 0
    xml = xml & "<Action Name=""SetLocalVar"">"
    xml = xml & "<Argument Name=""Name"">lngPkValue</Argument>"
    xml = xml & "<Argument Name=""Value"">0</Argument>"
    xml = xml & "</Action>"
    xml = xml & "</Statements></If>"

    ' Else block for updates
    xml = xml & "<Else><Statements>"

    ' Set PK value
    xml = xml & "<Action Name=""SetLocalVar"">"
    xml = xml & "<Argument Name=""Name"">lngPKValue</Argument>"
    xml = xml & "<Argument Name=""Value"">=[" & primaryKeyField & "]</Argument>"
    xml = xml & "</Action>"

    ' Set table name
    xml = xml & "<Action Name=""SetLocalVar"">"
    xml = xml & "<Argument Name=""Name"">strtableName</Argument>"
    xml = xml & "<Argument Name=""Value"">""" & tableName & """</Argument>"
    xml = xml & "</Action>"

    ' Add SetLocalVar for each Long Text field
    For Each fieldInfo In fieldList
        FieldName = fieldInfo(0)
        fldType = fieldInfo(1)

        If fldType = dbMemo Then
            xml = xml & "<Action Name=""SetLocalVar"">"
            xml = xml & "<Argument Name=""Name"">varLongTextBackup</Argument>"
            xml = xml & "<Argument Name=""Value"">BackupLongTextFieldsDM([strTableName],[lngPKValue],""" & FieldName & """)</Argument>"
            xml = xml & "</Action>"
        End If
    Next fieldInfo

    xml = xml & "</Statements></Else></ConditionalBlock>"
    xml = xml & "</Statements></DataMacro>"
    '    xml = xml & "</DataMacros>"

    BuildBeforeChangeMacro = xml
End Function

'-----------------------------------------------------------------------------
' Build Before Delete Macro XML (backup Long Text fields before deletes)
' Called by CreateAllDataMacros() for tables with Long Text fields
'-----------------------------------------------------------------------------
Private Function BuildBeforeDeleteMacro(tableName As String, fieldList As Collection, primaryKeyField As String) As String
    Dim xml                 As String
    Dim fieldInfo           As Variant
    Dim FieldName           As String
    Dim fldType             As Long
    Dim hasLongText         As Boolean

    ' Check if there are any Long Text fields
    hasLongText = False
    For Each fieldInfo In fieldList
        fldType = fieldInfo(1)
        If fldType = dbMemo Then
            hasLongText = True
            Exit For
        End If
    Next fieldInfo

    ' Only create macro if there are Long Text fields
    If Not hasLongText Then
        BuildBeforeDeleteMacro = ""
        Exit Function
    End If

    '    xml = "<?xml version=""1.0"" encoding=""UTF-16"" standalone=""no""?>"
    '    xml = xml & "<DataMacros xmlns=""http://schemas.microsoft.com/office/accessservices/2009/11/application"">"
    '    xml = xml & "<DataMacro Event=""BeforeDelete""><Statements>"
    xml = "<DataMacro Event=""BeforeDelete""><Statements>"
    ' Set PK value
    xml = xml & "<Action Name=""SetLocalVar"">"
    xml = xml & "<Argument Name=""Name"">lngPKValue</Argument>"
    xml = xml & "<Argument Name=""Value"">=[" & primaryKeyField & "]</Argument>"
    xml = xml & "</Action>"

    ' Set table name
    xml = xml & "<Action Name=""SetLocalVar"">"
    xml = xml & "<Argument Name=""Name"">strtableName</Argument>"
    xml = xml & "<Argument Name=""Value"">""" & tableName & """</Argument>"
    xml = xml & "</Action>"

    ' Add SetLocalVar for each Long Text field
    For Each fieldInfo In fieldList
        FieldName = fieldInfo(0)
        fldType = fieldInfo(1)

        If fldType = dbMemo Then
            xml = xml & "<Action Name=""SetLocalVar"">"
            xml = xml & "<Argument Name=""Name"">varLongTextBackup</Argument>"
            xml = xml & "<Argument Name=""Value"">BackupLongTextFieldsDM([strTableName],[lngPKValue],""" & FieldName & """)</Argument>"
            xml = xml & "</Action>"
        End If
    Next fieldInfo

    xml = xml & "</Statements></DataMacro>"
    '    xml = xml & "</DataMacros>"

    BuildBeforeDeleteMacro = xml

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
    Dim db                  As DAO.Database
    Dim tdf                 As DAO.TableDef
    Dim idx                 As DAO.Index
    Dim fld                 As DAO.Field

    Set db = CurrentDb
    Set tdf = db.TableDefs(tableName)

    ' Find primary key index
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



