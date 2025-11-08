Attribute VB_Name = "modAuditLongText"
Option Compare Database
Option Explicit


'================================================================================================================
'  VBA FOR LONG TEXT FIELD AUDITING
' November 8, 2025

' This function backs up Long Text field values to tblLongTextBackup
' before updates and deletes, so the After Update and After Delete Data Macros can retrieve them via LookupRecord.
'
' The After Update and After Delete data macros can retrieve the values from the backup table and insert them in the audit table.
' Compensates for inability to do String Comparisons on .Old and new values in Long Text fields to identify changes
' Because the Before Update and Before Delete Macros can execute VBA functions, they call this one as needed,
' eliminating the need for VBA in forms.
' This function needs to be in the BE for for creating the Data Macros on tables. The DMs won't recognize the code properly otherwise.
' This function needs to be in the FE for runtime execution by the Before Update and Before Delete Macros.
'================================================================================================================

Public Function BackupLongTextFieldsDM(strTableName As String, lngPKValue As Long, strFieldName As String)
100       On Error GoTo errHandler
          
      Dim db As DAO.Database
      Dim rs As DAO.Recordset
      Dim rsOldValue As DAO.Recordset
      Dim strPKField As String
      Dim strOldValue As Variant
          
110       Set db = CurrentDb
          
120       strTableName = strTableName
130       If strTableName = "" Then Exit Function
140       db.Execute "DELETE FROM tblLongTextBackup WHERE TableName='" & strTableName & "' AND  FieldName ='" & strFieldName & "' AND PrimaryKey=" & lngPKValue, dbFailOnError

150       strPKField = DLookup("FieldName", "tblDataMacroConfig", "TableName ='" & strTableName & "' AND IsPrimaryKey =" & True)
160       If lngPKValue > 0 Then ' updating an existing record or deleting only, don't run for new records
170           Set rsOldValue = db.OpenRecordset("SELECT " & strFieldName & " FROM " & strTableName & " WHERE " & strPKField & "= " & lngPKValue)
180           strOldValue = rsOldValue.Fields(strFieldName).Value
190           rsOldValue.Close
200           Set rs = db.OpenRecordset("tblLongTextBackup", dbOpenDynaset)
210           rs.AddNew
220           rs!tableName = strTableName
230           rs!PrimaryKey = lngPKValue
240           rs!FieldName = strFieldName
250           rs!OldValue = strOldValue
260           rs!DateChanged = Now()
270           rs!ChangedBy = CurrentUser()
280           rs.Update
290           rs.Close
300       End If

Cleanup:
          
310       On Error Resume Next
320       Set rsOldValue = Nothing
330       Set rs = Nothing
340       Set db = Nothing
350       Exit Function

errHandler:
          'uncomment in FE where the global error handler is available
          '500       Call GlblErrMsg( _
          '              sFrm:=Application.VBE.ActiveCodePane.CodeModule, _
          '              sCtl:="BackupLongTextFields")
360       Resume Cleanup

          '370       Resume
End Function


