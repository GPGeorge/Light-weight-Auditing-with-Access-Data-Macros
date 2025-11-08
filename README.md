# Light-weight-Auditing-with-Access-Data-Macros
Oct. 12, 2025

Creating a light-weight Auditing function for Access databases is a 5 step process. The Subs and Procedures in these code modules do all the work of setting it up. 
Customization is available to adapt the process to each Acess database.

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
' STEP 4: 

Import the Function in themodule modAuditLongText.bas into both the FE and BE accdbs.
Open the FE and link the 3 new tables created in the BE for audit loggin.
-----------------------------------------------------------------------------
NOTE: Any improvements or enhancements you would like to suggest are welcome.
