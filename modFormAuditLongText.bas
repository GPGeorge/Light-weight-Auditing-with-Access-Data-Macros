Attribute VB_Name = "modFormAuditLongText"
Option Compare Database
Option Explicit


'=============================================================================
' FORM VBA FOR LONG TEXT FIELD AUDITING
'
' These functions backup Long Text field values to tblLongTextBackup
' before updates and deletes, so Data Macros can retrieve them via LookupRecord.
'
' USAGE IN YOUR FORMS BOUND TO AUDITED TABLES:
'
' In Form_BeforeUpdate event:
'     Call BackupLongTextFields(Me)
'
' In Form_BeforeDelConfirm event:
'     Call BackupLongTextFieldsBeforeDelete(Me)

'--------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Call this from Form_BeforeUpdate event
' Backs up all Long Text fields before the record is updated
' The After Update and After Delete data macros can retrieve the values from the backup table and insert them in the audit table.
' Compensates for inability to do String Comparisons on .Old and new values in Long Text fields to identify changes
'   Simply logs all values to Audit Table.
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------

'=============================================================================

Public Sub BackupLongTextFields(frm As Form)
100       On Error GoTo errHandler
          
      Dim db As DAO.Database
      Dim rs As DAO.Recordset
      Dim rsConfig As DAO.Recordset
      Dim strTableName As String
      Dim strPKField As String
      Dim strPKControl As String
      Dim strControlName As String
      Dim varPKValue As Variant
          
110       If frm.NewRecord Then Exit Sub
          
120       Set db = CurrentDb
          
130       strTableName = GetTableNameFromForm(frm)
140       If strTableName = "" Then Exit Sub
          
150       Set rsConfig = db.OpenRecordset("SELECT FieldName FROM tblDataMacroConfig WHERE TableName='" & strTableName & "' AND IsPrimaryKey=True", dbOpenSnapshot)
160       If rsConfig.EOF Then
170           rsConfig.Close
180           Exit Sub
190       End If
200       strPKField = rsConfig!FieldName
210       rsConfig.Close
220       strPKControl = FindControlBoundToField(frm, strPKField)
230       varPKValue = frm.Controls(strPKControl).OldValue
         
240       db.Execute "DELETE FROM tblLongTextBackup WHERE TableName='" & strTableName & "' AND PrimaryKey=" & varPKValue, dbFailOnError
          
250       Set rsConfig = db.OpenRecordset("SELECT FieldName FROM tblDataMacroConfig WHERE TableName='" & strTableName & "' AND DataType=" & dbMemo, dbOpenSnapshot)
          
260       Do While Not rsConfig.EOF
270           strControlName = FindControlBoundToField(frm, rsConfig!FieldName)
280           If Len(strControlName) > 0 Then ' if it's a ZLS, we already know it doesn't exist on the form, so don't bother with it
290               If ControlExists(frm, strControlName) Then
300                   Set rs = db.OpenRecordset("tblLongTextBackup", dbOpenDynaset)
310                   rs.AddNew
320                   rs!tableName = strTableName
330                   rs!PrimaryKey = varPKValue
340                   rs!FieldName = rsConfig!FieldName
350                   rs!OldValue = IIf(Len(frm.Controls(strControlName).OldValue & "") = 0, "N/A", frm.Controls(strControlName).OldValue)
360                   rs!DateChanged = Now()
370                   rs!ChangedBy = CurrentUser()
380                   rs.Update
390                   rs.Close
400               End If
410           End If
420           rsConfig.MoveNext
430       Loop
          
440       rsConfig.Close

          
CleanUp:
          
450       On Error Resume Next
460       Set rsConfig = Nothing
470       Set rs = Nothing
480       Set db = Nothing
490       Exit Sub

errHandler:
500       Call GlblErrMsg( _
              sFrm:=Application.VBE.ActiveCodePane.CodeModule, _
              sCtl:="BackupLongTextFields")
510       Resume CleanUp

520       Resume
End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
' Call this from Form_Delete event (NOT BeforeDelConfirm or Delete)
' Backs up all Long Text fields before the record is deleted
' The After Update and After Delete data macros can retrieve the values from the backup table and insert them in the audit table.
' Compensates for inability to do String Comparisons on .Old and new values in Long Text fields to identify deletions
'   Simply logs all values to Audit Table.
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Sub BackupLongTextFieldsBeforeDelete(frm As Form)
          
100       On Error GoTo errHandler

      Dim db As DAO.Database
      Dim rs As DAO.Recordset
      Dim rsConfig As DAO.Recordset
      Dim strTableName As String
      Dim strPKField As String
      Dim strPKControl As String
      Dim strControlName As String
      Dim varPKValue As Variant
          
110       Set db = CurrentDb
          
          ' Get table name from form's RecordSource
120       strTableName = GetTableNameFromForm(frm)
130       If strTableName = "" Then Exit Sub
          
          ' Get primary key field from config table
140       Set rsConfig = db.OpenRecordset("SELECT FieldName FROM tblDataMacroConfig WHERE TableName='" & strTableName & "' AND IsPrimaryKey=True", dbOpenSnapshot)
150       If rsConfig.EOF Then
160           rsConfig.Close
170           Exit Sub
180       End If
190       strPKField = rsConfig!FieldName
200       rsConfig.Close
          ' Get old value from the control bound to the PK (still available in Form_Delete)
210       strPKControl = FindControlBoundToField(frm, strPKField)
220       varPKValue = frm.Controls(strPKControl).Value
          
          ' Delete any existing backup records for this record
230       db.Execute "DELETE FROM tblLongTextBackup WHERE TableName='" & strTableName & "' AND PrimaryKey=" & varPKValue, dbFailOnError
          
          ' Get Long Text fields from config table
240       Set rsConfig = db.OpenRecordset("SELECT FieldName FROM tblDataMacroConfig WHERE TableName='" & strTableName & "' AND DataType=" & dbMemo, dbOpenSnapshot)
          
250       Do While Not rsConfig.EOF
260           strControlName = FindControlBoundToField(frm, rsConfig!FieldName)
270           If Len(strControlName) > 0 Then ' if it's a ZLS, we already know it doesn't exist on the form, so don't bother with it
280               If ControlExists(frm, strControlName) Then
                      ' Insert backup record using current Value (not OldValue)
290                   Set rs = db.OpenRecordset("tblLongTextBackup", dbOpenDynaset)
300                   rs.AddNew
310                   rs!tableName = strTableName
320                   rs!PrimaryKey = varPKValue
330                   rs!FieldName = rsConfig!FieldName
340                   rs!OldValue = Nz(frm.Controls(rsConfig!FieldName).Value, "N/A")
350                   rs!DateChanged = Now()
360                   rs!ChangedBy = CurrentUser()
370                   rs.Update
380                   rs.Close
390               End If
400           End If
410           rsConfig.MoveNext
420       Loop
430       rsConfig.Close
          
440       Set rsConfig = Nothing
450       Set rs = Nothing
460       Set db = Nothing
470       Exit Sub
          

          
CleanUp:
          
480       On Error Resume Next
490       Exit Sub

errHandler:
500       Call GlblErrMsg( _
              sFrm:=Application.VBE.ActiveCodePane.CodeModule, _
              sCtl:="BackupLongTextFieldsBeforeDelete")
510       Resume CleanUp

520       Resume
End Sub
'


'-----------------------------------------------------------------------------
' Helper: Extract table name from form's RecordSource
'-----------------------------------------------------------------------------
Private Function GetTableNameFromForm(frm As Form) As String
Dim strRecordSource As String
    
    strRecordSource = frm.RecordSource
    
    If Left(strRecordSource, 3) = "tbl" Then ' Weak assumption is that all auditable tables have the "tbl" prefix
        GetTableNameFromForm = strRecordSource
    Else
        GetTableNameFromForm = GetTableFromQuery(strRecordSource)
    End If
End Function


'-----------------------------------------------------------------------------
' Helper: Get table name from a query
'         Assumes the query uses one table.
'-----------------------------------------------------------------------------
Private Function GetTableFromQuery(queryName As String) As String
    On Error Resume Next
Dim db As DAO.Database
Dim qdf As DAO.QueryDef
Dim strSQL As String
Dim lngPos As Long
    
    Set db = CurrentDb
    Set qdf = db.QueryDefs(queryName)
    
    If Not qdf Is Nothing Then
        strSQL = qdf.SQL
        
        lngPos = InStr(1, UCase(strSQL), "FROM ")
        If lngPos > 0 Then
            strSQL = Mid(strSQL, lngPos + 5)
            strSQL = Trim(Split(strSQL, " ")(0))
            strSQL = Replace(strSQL, "[", "")
            strSQL = Replace(strSQL, "]", "")
            strSQL = Replace(strSQL, ";", "")
            GetTableFromQuery = strSQL
        End If
    End If
    
    Set qdf = Nothing
    Set db = Nothing
End Function


'-----------------------------------------------------------------------------
' Helper: Check if a control exists on the form
'-----------------------------------------------------------------------------
Private Function ControlExists(frm As Form, controlName As String) As Boolean
    On Error Resume Next
Dim ctl As control
    Set ctl = frm.Controls(controlName)
    ControlExists = (Err.Number = 0)
    On Error GoTo 0
End Function
'---------------------------------------------------------------------------------------------
' Helper: Find the control to which the Primary Key is bound and return the control name
'---------------------------------------------------------------------------------------------
Private Function FindControlBoundToField(frm As Form, FieldName As String) As String
Dim ctl As control

    For Each ctl In frm.Controls

        If ctl.ControlType = acTextBox _
            Or ctl.ControlType = acComboBox _
            Or ctl.ControlType = acListBox Then

            If Len(Trim(ctl.ControlSource & "")) > 0 Then
                If StrComp(ctl.ControlSource, FieldName, vbTextCompare) = 0 Then
                    FindControlBoundToField = ctl.Name
                    Exit Function
                End If
            End If
        End If
    Next ctl

    FindControlBoundToField = ""
End Function
Private Function GetPKFieldName(frm As Form) As String
Dim db As DAO.Database
Dim rs As DAO.Recordset
Dim strTableName As String

    Set db = CurrentDb
    strTableName = frm.RecordSource

    Set rs = db.OpenRecordset("SELECT FieldName FROM tblDataMacroConfig WHERE TableName='" & strTableName & "' AND IsPrimaryKey=True")
    If Not rs.EOF Then
        GetPKFieldName = rs!FieldName
    End If
    rs.Close
End Function

'=============================================================================
' EXAMPLE USAGE IN A FORM MODULE
'=============================================================================
'
' Copy this code into your form's code module:
'
' Private Sub Form_BeforeUpdate(Cancel As Integer)
'     ' Backup Long Text fields before saving changes
'     Call BackupLongTextFields(Me)
' End Sub
'
' Private Sub Form_Delete(Cancel As Integer)
'     ' Capture Primary Key before it's lost
'     ' Replace YourPKFieldName with your actual PK field name
'     TempVars.Add Name:="varPKValue", Value:=Me.YourPKFieldName.Value
' End Sub
'
' Private Sub Form_BeforeDelConfirm(Cancel As Integer, Response As Integer)
'     ' Backup Long Text fields before deleting record
'     Call BackupLongTextFieldsBeforeDelete(Me)
' End Sub
'
' OPTIONAL: To automatically get PK field name instead of hardcoding:
' Add this helper function to get PK from config:
'
'Private Sub Form_Delete(Cancel As Integer)
'    ' Automatically get PK field from config
'    Dim db As DAO.Database
'    Dim rs As DAO.Recordset
'    Dim strTableName As String
'    Dim strPKField As String
'
'    Set db = CurrentDb
'    strTableName = Me.RecordSource ' Or GetTableNameFromForm(Me)
'
'    Set rs = db.OpenRecordset("SELECT FieldName FROM tblDataMacroConfig WHERE TableName='" & strTableName & "' AND IsPrimaryKey=True")
'    If Not rs.EOF Then
'        strPKField = rs!fieldName
'        TempVars.Add Name:="varPKValue", Value:=Me.Controls(strPKField).Value
'    End If
'    rs.Close
'End Sub
' Then in Form_Delete:
'     TempVars.Add Name:="varPKValue", Value:=Me.Controls(GetPKFieldName()).Value
'
'============================
 






