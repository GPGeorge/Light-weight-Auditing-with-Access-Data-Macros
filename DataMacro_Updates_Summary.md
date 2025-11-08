# Data Macro System Updates

## Overview
Updated the `modAddDataMacros` module to automatically generate **BeforeChange** and **BeforeDelete** Data Macros for tables containing Long Text fields, in addition to the existing **AfterInsert**, **AfterUpdate**, and **AfterDelete** macros.

## What Changed

### New Functions Added

#### 1. BuildBeforeChangeMacro()
- Generates XML for the BeforeChange Data Macro
- Checks if this is a new record (`IsNull([Old].[PrimaryKey])`)
- For existing records (updates), backs up Long Text field values by calling `BackupLongTextFieldsDM()` for each Long Text field
- Sets local variables: `lngPKValue`, `strTableName`, and `varLongTextBackup`

#### 2. BuildBeforeDeleteMacro()
- Generates XML for the BeforeDelete Data Macro
- Backs up Long Text field values before deletion by calling `BackupLongTextFieldsDM()` for each Long Text field
- Sets local variables: `lngPKValue`, `strTableName`, and `varLongTextBackup`

### Modified Functions

#### CreateAllDataMacros()
Enhanced to:
1. Check if any field in the table is a Long Text field (DataType = dbMemo)
2. If Long Text fields are present:
   - Generate and load BeforeChange macro (separate XML file)
   - Generate and load BeforeDelete macro (separate XML file)
   - Report: "All 5 data macros created"
3. If no Long Text fields:
   - Generate only the 3 After macros as before
   - Report: "All 3 data macros created"

## How It Works

### For Tables WITHOUT Long Text Fields
The system creates 3 Data Macros:
- AfterInsert
- AfterUpdate
- AfterDelete

### For Tables WITH Long Text Fields
The system creates 5 Data Macros:
- **BeforeChange** (new) - Backs up Long Text values before updates
- **BeforeDelete** (new) - Backs up Long Text values before deletes
- AfterInsert
- AfterUpdate (retrieves old Long Text values from tblLongTextBackup)
- AfterDelete (retrieves old Long Text values from tblLongTextBackup)

## Technical Details

### XML Namespaces
- BeforeChange/BeforeDelete use: `http://schemas.microsoft.com/office/accessservices/2009/11/application`
- After macros use: `http://schemas.microsoft.com/office/accessservices/2010/12/application`

### Separate File Loading
BeforeChange and BeforeDelete macros are loaded as separate XML files because they use a different namespace and cannot be combined with the After macros in a single file.

### VBA Function Integration
The BeforeChange and BeforeDelete macros call the `BackupLongTextFieldsDM()` function from `modFormAuditLongText`, which:
1. Deletes any existing backup for that table/field/record
2. Reads the current Long Text value from the table
3. Stores it in `tblLongTextBackup`
4. Timestamps and tags with current user

## Benefits
1. **Automatic Detection** - No manual configuration needed for Long Text fields
2. **Comprehensive Coverage** - All 5 macros are created automatically when needed
3. **Efficient** - Only adds BeforeChange/BeforeDelete when necessary
4. **Maintainable** - Single function call still creates entire audit system

## Usage
No changes to how you run the setup:
```vba
1. One_CreateAuditTables()
2. Two_PopulateConfigTable()
3. Three_GenerateAllAuditDataMacros()
```

The system now automatically detects Long Text fields and creates all required macros accordingly.
