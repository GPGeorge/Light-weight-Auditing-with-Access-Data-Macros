# BeforeChange and BeforeDelete Data Macros - Quick Reference

## Purpose
These macros back up Long Text field values BEFORE changes occur, allowing the AfterUpdate and AfterDelete macros to retrieve the old values from tblLongTextBackup.

## BeforeChange Data Macro

### When It Fires
- Before any INSERT or UPDATE operation on the table

### What It Does
1. **For INSERT operations** (new records):
   - Sets `lngPKValue = 0` (skips backup since there's no old value)

2. **For UPDATE operations** (existing records):
   - Captures the Primary Key value
   - Calls `BackupLongTextFieldsDM()` for each Long Text field
   - Stores old Long Text values in `tblLongTextBackup`

### Sample XML Structure
```xml
<?xml version="1.0" encoding="UTF-16" standalone="no"?>
<DataMacros xmlns="http://schemas.microsoft.com/office/accessservices/2009/11/application">
  <DataMacro Event="BeforeChange">
    <Statements>
      <ConditionalBlock>
        <If>
          <Condition>IsNull([Old].[PublicationID])</Condition>
          <Statements>
            <!-- New Record - Do Nothing -->
            <Action Name="SetLocalVar">
              <Argument Name="Name">lngPkValue</Argument>
              <Argument Name="Value">0</Argument>
            </Action>
          </Statements>
        </If>
        <Else>
          <Statements>
            <!-- Existing Record - Backup Long Text Fields -->
            <Action Name="SetLocalVar">
              <Argument Name="Name">lngPKValue</Argument>
              <Argument Name="Value">=[PublicationID]</Argument>
            </Action>
            <Action Name="SetLocalVar">
              <Argument Name="Name">strtableName</Argument>
              <Argument Name="Value">"tblPublication"</Argument>
            </Action>
            <!-- Repeat for each Long Text field -->
            <Action Name="SetLocalVar">
              <Argument Name="Name">varLongTextBackup</Argument>
              <Argument Name="Value">BackupLongTextFieldsDM([strTableName],[lngPKValue],"PublicationTitle")</Argument>
            </Action>
            <Action Name="SetLocalVar">
              <Argument Name="Name">varLongTextBackup</Argument>
              <Argument Name="Value">BackupLongTextFieldsDM([strTableName],[lngPKValue],"Comments")</Argument>
            </Action>
          </Statements>
        </Else>
      </ConditionalBlock>
    </Statements>
  </DataMacro>
</DataMacros>
```

## BeforeDelete Data Macro

### When It Fires
- Before any DELETE operation on the table

### What It Does
1. Captures the Primary Key value of the record being deleted
2. Calls `BackupLongTextFieldsDM()` for each Long Text field
3. Stores the Long Text values in `tblLongTextBackup` before deletion

### Sample XML Structure
```xml
<?xml version="1.0" encoding="UTF-16" standalone="no"?>
<DataMacros xmlns="http://schemas.microsoft.com/office/accessservices/2009/11/application">
  <DataMacro Event="BeforeDelete">
    <Statements>
      <Action Name="SetLocalVar">
        <Argument Name="Name">lngPKValue</Argument>
        <Argument Name="Value">=[PublicationID]</Argument>
      </Action>
      <Action Name="SetLocalVar">
        <Argument Name="Name">strtableName</Argument>
        <Argument Name="Value">"tblPublication"</Argument>
      </Action>
      <!-- Repeat for each Long Text field -->
      <Action Name="SetLocalVar">
        <Argument Name="Name">varLongTextBackup</Argument>
        <Argument Name="Value">BackupLongTextFieldsDM([strTableName],[lngPKValue],"PublicationTitle")</Argument>
      </Action>
      <Action Name="SetLocalVar">
        <Argument Name="Name">varLongTextBackup</Argument>
        <Argument Name="Value">BackupLongTextFieldsDM([strTableName],[lngPKValue],"Comments")</Argument>
      </Action>
    </Statements>
  </DataMacro>
</DataMacros>
```

## Key Points

### Why Separate Files?
BeforeChange and BeforeDelete use the 2009 namespace, while AfterInsert/AfterUpdate/AfterDelete use the 2010 namespace. They cannot be combined in a single XML file.

### Local Variables
- `lngPKValue` - Primary key of the record
- `strTableName` - Name of the table being modified
- `varLongTextBackup` - Return value from BackupLongTextFieldsDM (not used, but required for function call)

### BackupLongTextFieldsDM Function
This VBA function (in modFormAuditLongText):
1. Reads the current Long Text value from the table
2. Deletes any existing backup for this table/field/record
3. Inserts new backup record with:
   - TableName
   - PrimaryKey
   - FieldName
   - OldValue (the Long Text content)
   - DateChanged
   - ChangedBy

### After Macros Retrieve Backup
The AfterUpdate and AfterDelete macros use LookupRecord to retrieve these backed-up values from tblLongTextBackup when creating audit log entries.

## Complete Flow Example

### UPDATE Operation
1. User changes a record with Long Text fields
2. **BeforeChange** fires → Backs up old Long Text values to tblLongTextBackup
3. Record is updated
4. **AfterUpdate** fires → Looks up old values from tblLongTextBackup → Creates audit log

### DELETE Operation
1. User deletes a record with Long Text fields
2. **BeforeDelete** fires → Backs up Long Text values to tblLongTextBackup
3. Record is deleted
4. **AfterDelete** fires → Looks up old values from tblLongTextBackup → Creates audit log

This ensures complete audit trail for Long Text fields despite Access Data Macro limitations!
