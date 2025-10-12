Attribute VB_Name = "modFormAuditLongText"
Option Compare Database
Option Explicit


'=============================================================================
' FORM VBA FOR LONG TEXT FIELD AUDITING
'
' These functions backup Long Text field values to tblLongTextBackup
' before updates and deletes, so Data Macros can retrieve them via LookupRecord.
'
' USAGE IN YOUR FORMS:
'
' In Form_BeforeUpdate event:
'     Call BackupLongTextFields(Me)
'
' In Form_BeforeDelConfirm event:
'     Call BackupLongTextFieldsBeforeDelete(Me)
'=============================================================================

'-----------------------------------------------------------------------------
' Call this from Form_BeforeUpdate event
' Backs up all Long Text fields before the record is updated
'-----------------------------------------------------------------------------
Public Sub BackupLongTextFields(frm As Form)
    On Error GoTo ErrHandler
    
Dim db As DAO.Database
Dim rs As DAO.Recordset
Dim rsConfig As DAO.Recordset
Dim strTableName As String
Dim strPKField As String
Dim strPKControl As String
Dim strControlName As String
Dim varPKValue As Variant
    
    ' Exit if new record (nothing to backup)
    If frm.NewRecord Then Exit Sub
    
    Set db = CurrentDb
    
    ' Get table name from form's RecordSource
    strTableName = GetTableNameFromForm(frm)
    If strTableName = "" Then Exit Sub
    
    ' Get primary key field from config table
    Set rsConfig = db.OpenRecordset("SELECT FieldName FROM tblDataMacroConfig WHERE TableName='" & strTableName & "' AND IsPrimaryKey=True", dbOpenSnapshot)
    If rsConfig.EOF Then
        rsConfig.Close
        Exit Sub
    End If
    strPKField = rsConfig!FieldName
    rsConfig.Close
    strPKControl = FindControlBoundToField(frm, strPKField)
    varPKValue = frm.Controls(strPKControl).OldValue
   
    
    ' Delete any existing backup records for this record
    db.Execute "DELETE FROM tblLongTextBackup WHERE TableName='" & strTableName & "' AND PrimaryKey=" & varPKValue, dbFailOnError
    
    ' Get Long Text fields from config table
    Set rsConfig = db.OpenRecordset("SELECT FieldName FROM tblDataMacroConfig WHERE TableName='" & strTableName & "' AND DataType=" & dbMemo, dbOpenSnapshot)
    
    Do While Not rsConfig.EOF
        strControlName = FindControlBoundToField(frm, rsConfig!FieldName)
        If Len(strControlName) > 0 Then ' if it's a ZLS, we already know it doesn't exist on the form, so don't bother with it
            If ControlExists(frm, strControlName) Then
                ' Insert backup record
                Set rs = db.OpenRecordset("tblLongTextBackup", dbOpenDynaset)
                rs.AddNew
                rs!tableName = strTableName
                rs!PrimaryKey = varPKValue
                rs!FieldName = rsConfig!FieldName
                rs!OldValue = IIf(Len(frm.Controls(strControlName).OldValue & "") = 0, "N/A", frm.Controls(strControlName).OldValue)
                rs!DateChanged = Now()
                rs!ChangedBy = CurrentUser()
                rs.Update
                rs.Close
            End If
        End If
        rsConfig.MoveNext
    Loop
    rsConfig.Close
    
    Set rsConfig = Nothing
    Set rs = Nothing
    Set db = Nothing
    Exit Sub
    
CleanUp:
    
    On Error Resume Next
    Exit Sub

ErrHandler:
    Call GlblErrMsg( _
        sFrm:=Application.vbe.ActiveCodePane.CodeModule, _
        sCtl:="BackupLongTextFields")
    Resume CleanUp

    Resume
End Sub
'-----------------------------------------------------------------------------
' Call this from Form_Delete event (NOT BeforeDelConfirm)
' Backs up all Long Text fields before the record is deleted
'-----------------------------------------------------------------------------
Public Sub BackupLongTextFieldsBeforeDelete(frm As Form)
          
100       On Error GoTo ErrHandler

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

ErrHandler:
500       Call GlblErrMsg( _
              sFrm:=Application.vbe.ActiveCodePane.CodeModule, _
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
    
    ' If RecordSource is a table name, return it
    ' If it's a query, you may need additional logic
    If Left(strRecordSource, 3) = "tbl" Then
        GetTableNameFromForm = strRecordSource
    Else
        ' Handle query-based forms - get table from query
        GetTableNameFromForm = GetTableFromQuery(strRecordSource)
    End If
End Function


'-----------------------------------------------------------------------------
' Helper: Get table name from a query
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
        
        ' Simple parser to extract table name after FROM
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

Private Function FindControlBoundToField(frm As Form, FieldName As String) As String
Dim ctl As control

    For Each ctl In frm.Controls
        ' Only check controls that can have a ControlSource
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
 
Private Function GetPKFieldName(frm As Form) As String
Dim db As DAO.Database
Dim rs As DAO.Recordset
Dim strTableName As String

    Set db = CurrentDb
    strTableName = frm.RecordSource  ' Or use GetTableNameFromForm

    Set rs = db.OpenRecordset("SELECT FieldName FROM tblDataMacroConfig WHERE TableName='" & strTableName & "' AND IsPrimaryKey=True")
    If Not rs.EOF Then
        GetPKFieldName = rs!FieldName
    End If
    rs.Close
End Function




