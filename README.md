# Light-weight-Auditing-with-Access-Data-Macros
Oct. 12, 2025

Creating a light-weight Auditing function for Access databases is a 4 step process. The Subs and Procedures in these two code modules do all the work of setting it up. Customization is available to adapt the process to each Acess database.

'-----------------------------------------------------------------------------
' STEP 1: Create 3 required tables
'-----------------------------------------------------------------------------
         
Run Public Sub CreateAuditTables in module modDataMacros

'-----------------------------------------------------------------------------
' STEP 2: Populate configuration table with your tables and fields
'         Customize by including/excluding specific tables and fields in your database.
'-----------------------------------------------------------------------------
         
Run Public Sub PopulateConfigTable() in module modDataMacros

-----------------------------------------------------------------------------
' STEP 3: Generate all Data Macros for all tables and fields in the  tblDataMacroConfig table
'-----------------------------------------------------------------------------

Run Public Sub GenerateAllAuditDataMacros() in module modDataMacros

-----------------------------------------------------------------------------
' STEP 4: Copy 2 functions into each form bound to a table with one or more Long Text fields
          Long Text fields are not auditable with Data Macros, so VBA helper functions are required
'-----------------------------------------------------------------------------
Copy and paste Public Sub BackupLongTextFields(frm As Form)
Copy and paste Public Sub BackupLongTextFieldsBeforeDelete(frm As Form)

You can import the module modFormAuditLongText, which contains these two form level subs along with helper functions required
Private Function ControlExists(frm As Form, controlName As String) As Boolean
Private Function FindControlBoundToField(frm As Form, FieldName As String) As String
Private Function GetPKFieldName(frm As Form) As String
Private Function GetTableFromQuery(queryName As String) As String
Private Function GetTableFromQuery(queryName As String) As String

NOTE: Any improvements or enhancements you would like to suggest are welcome.
