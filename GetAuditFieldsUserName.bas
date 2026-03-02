
'NOTE:
'   If you split the database into a Front End and Back End
'   this module needs to be where the tables are (the Back End)
'   because these function(s) are called by data macros in the tables.
'PURPOSE:
'   Provide name to be used when saving to the Audit Trail fields.
'   This function is called from the data macros in most tables.
Public Function GetAuditFieldsUserName() As String
    'NO ERRORIZE        'This code will be in Access BE where the error handler code does not exist.

    Dim strUserName         As String

    strUserName = WindowsUserName()
    GetAuditFieldsUserName = strUserName 
End Function

