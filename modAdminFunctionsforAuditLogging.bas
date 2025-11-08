Attribute VB_Name = "modAdminFunctionsforAuditLogging"
Public Function BackupAndRemoveAllDataMacros(Optional strBackupPath As String = "") As Boolean
100       On Error GoTo ErrorHandler

      Dim db As DAO.Database
      Dim rst As DAO.Recordset
      Dim strSQL As String
      Dim strTempFile As String
      Dim strBackupFile As String
      Dim intFileNum As Integer
      Dim intMacrosRemoved As Integer

110       Set db = CurrentDb
120       intMacrosRemoved = 0

          ' Set backup path
130       If strBackupPath = "" Then
140           strBackupPath = CurrentProject.Path & "\DataMacroBackups\"
150       End If

          ' Create backup folder if it doesn't exist
160       If Dir(strBackupPath, vbDirectory) = "" Then
170           MkDir strBackupPath
180       End If

          ' Create a temporary blank data macro XML file
190       strTempFile = Environ("TEMP") & "\BlankDataMacro.xml"

200       intFileNum = FreeFile
210       Open strTempFile For Output As intFileNum
220       Print #intFileNum, "<?xml version=""1.0"" encoding=""UTF-16""?>"
230       Print #intFileNum, "<DataMacros xmlns=""http://schemas.microsoft.com/office/accessservices/2009/04/application"">"
240       Print #intFileNum, "</DataMacros>"
250       Close #intFileNum

          ' Find tables with data macros
260       strSQL = "SELECT [Name] FROM MSysObjects " & _
              "WHERE Not IsNull(LvExtra) AND Type = 1 " & _
              "ORDER BY [Name]"

270       Set rst = db.OpenRecordset(strSQL, dbOpenSnapshot)

280       Do While Not rst.EOF
290   Debug.Print "Processing: " & rst!Name

              ' Backup the data macro first
300           strBackupFile = strBackupPath & rst!Name & "_DataMacro_" & _
                  Format(Now(), "yyyymmdd_hhnnss") & ".xml"
310           Application.SaveAsText acTableDataMacro, rst!Name, strBackupFile
320   Debug.Print "  Backed up to: " & strBackupFile

              ' Remove data macros by loading blank XML to replace the existing data macros, if any
330           Application.LoadFromText acTableDataMacro, rst!Name, strTempFile
340   Debug.Print "  Data macros removed"

350           intMacrosRemoved = intMacrosRemoved + 1
360           rst.MoveNext
370       Loop

380       rst.Close
390       Set rst = Nothing
400       Kill strTempFile

410       MsgBox "Successfully backed up and removed data macros from " & intMacrosRemoved & " tables." & vbCrLf & _
              "Backups saved to: " & strBackupPath, vbInformation, "Data Macros Removed"

420       BackupAndRemoveAllDataMacros = True
Cleanup:
430       Exit Function

ErrorHandler: 'If desired, replace with your own Error Handling
440       If Err <> 2950 Then
450   Debug.Print Err & " Error: " & Err.Description
460           MsgBox Err & " Error: " & Err.Description, vbExclamation
470           On Error Resume Next
480           If Not rst Is Nothing Then rst.Close
490           If Dir(strTempFile) <> "" Then Kill strTempFile
500           BackupAndRemoveAllDataMacros = False
510       Else
520   Debug.Print rst!Name & "  No Data Macros to process "
530           Resume Cleanup
540           Resume Next

550       End If
End Function
'
Public Sub ListAllTableProperties()
      Dim db As DAO.Database
      Dim tdf As DAO.TableDef
      Dim prp As DAO.Property
          
100       Set db = CurrentDb
          
110   Debug.Print "=== ALL Table Properties ==="
          
          ' Check the tables you showed in the screenshot
      Dim arrTables As Variant
120       arrTables = Array("tblApprovedKeyWords", "tblBannedKeyWord", "tblBookcase")
          
      Dim i As Integer
130       For i = LBound(arrTables) To UBound(arrTables)
140           On Error Resume Next
150           Set tdf = db.TableDefs(arrTables(i))
160           If Err.Number = 0 Then
170   Debug.Print vbCrLf & "Table: " & tdf.Name
180   Debug.Print String(50, "-")
190               For Each prp In tdf.Properties
200   Debug.Print "  Property: " & prp.Name & " | Type: " & prp.type
210               Next prp
220           End If
230           On Error GoTo 0
240       Next i
          
250   Debug.Print vbCrLf & "=== End of List ==="
End Sub
'=====================================================================================================
' Clean up tables by removing invalid fields from tables.
' Example: SQL Server tables imported into the Access accdb can have fields only relevant to the SQL Server environment
'                   such as TimeStamp, or Rowversion, fields.
'=====================================================================================================
Public Function RemoveInvalidFields(fldName As String) As Boolean
100       On Error GoTo ErrorHandler
          
      Dim db As DAO.Database
      Dim tdf As DAO.TableDef
      Dim fld As DAO.Field
      Dim strTableName As String
      Dim intTablesProcessed As Integer
      Dim intFieldsRemoved As Integer
      Dim strRemovedFields As String
          
110       Set db = CurrentDb
120       intTablesProcessed = 0
130       intFieldsRemoved = 0
          
          ' Loop through all tables
140       For Each tdf In db.TableDefs
              ' Skip system and temporary tables
150           If Left(tdf.Name, 4) <> "MSys" And Left(tdf.Name, 1) <> "~" Then
160               strTableName = tdf.Name
170               strRemovedFields = ""
                  
                  ' Try to remove ValidFrom field
180               On Error Resume Next
190               Set fld = tdf.Fields(fldName)
200               If Err.Number = 0 Then
210                   tdf.Fields.Delete fldName
220                   If Err.Number = 0 Then
230                       intFieldsRemoved = intFieldsRemoved + 1
240                       strRemovedFields = fldName
250   Debug.Print strTableName & ": Removed " & fldName
260                   End If
270               End If
280               Err.Clear
                  
400               On Error GoTo ErrorHandler
                  
410               intTablesProcessed = intTablesProcessed + 1
420           End If
430       Next tdf
          
          ' Display results
440   Debug.Print vbCrLf & "Processed " & intTablesProcessed & " tables"
450   Debug.Print "Removed " & intFieldsRemoved & " fields total"
          
460       MsgBox "Successfully processed " & intTablesProcessed & " tables." & vbCrLf & _
              "Removed " & intFieldsRemoved & " field(s) total.", _
              vbInformation, "Fields Removed"
          
470       RemoveInvalidFields = True
480       Exit Function
          
ErrorHandler:
490   Debug.Print "Error in RemoveInvalidFields: " & Err.Description & " (Table: " & strTableName & ")"
500       MsgBox "Error removing fields: " & Err.Description & vbCrLf & _
              "Table: " & strTableName, vbExclamation
510       RemoveInvalidFields = False
End Function

Public Function FindModernChartControls() As Boolean
100       On Error GoTo ErrorHandler
          
      Dim db As DAO.Database
      Dim doc As Object
      Dim frm As Access.Form
      Dim rpt As Access.Report
      Dim ctl As Access.control
      Dim strObjectName As String
      Dim intChartsFound As Integer
      Dim strResults As String
          
110       Set db = CurrentDb
120       intChartsFound = 0
130       strResults = "Modern Chart Controls Found:" & vbCrLf & vbCrLf
          
          ' Check all Forms
140   Debug.Print "=== Checking Forms ==="
150       For Each doc In CurrentProject.AllForms
160           strObjectName = doc.Name
              
              ' Open form in design view (hidden)
170           DoCmd.OpenForm strObjectName, acDesign, , , , acHidden
180           Set frm = Forms(strObjectName)
              
              ' Check each control on the form
190           For Each ctl In frm.Controls
                  ' Modern Chart controls are typically Web Browser controls (type 128)
                  ' or have ControlType of acWebBrowser
200               If ctl.ControlType = acWebBrowser Then
                      ' Check if it might be a chart by looking at properties
210                   On Error Resume Next
220                   If InStr(1, ctl.ControlSource, "chart", vbTextCompare) > 0 Or _
                          InStr(1, ctl.Name, "chart", vbTextCompare) > 0 Then
230   Debug.Print "FOUND: Form [" & strObjectName & "] - Control: " & ctl.Name
240                       strResults = strResults & "Form: " & strObjectName & " - Control: " & ctl.Name & vbCrLf
250                       intChartsFound = intChartsFound + 1
260                   End If
270                   On Error GoTo ErrorHandler
                      ' Also check for Modern Chart specific control type (143)
280               ElseIf ctl.ControlType = 143 Then
290   Debug.Print "FOUND: Form [" & strObjectName & "] - Control: " & ctl.Name & " (Modern Chart)"
300                   strResults = strResults & "Form: " & strObjectName & " - Control: " & ctl.Name & " (Modern Chart)" & vbCrLf
310                   intChartsFound = intChartsFound + 1
320               End If
330           Next ctl
              
              ' Close the form
340           DoCmd.Close acForm, strObjectName, acSaveNo
350       Next doc
          
          ' Check all Reports
360   Debug.Print vbCrLf & "=== Checking Reports ==="
370       For Each doc In CurrentProject.AllReports
380           strObjectName = doc.Name
              
              ' Open report in design view (hidden)
390           DoCmd.OpenReport strObjectName, acViewDesign, , , acHidden
400           Set rpt = Reports(strObjectName)
              
              ' Check each control on the report
410           For Each ctl In rpt.Controls
420               If ctl.ControlType = acWebBrowser Then
430                   On Error Resume Next
440                   If InStr(1, ctl.ControlSource, "chart", vbTextCompare) > 0 Or _
                          InStr(1, ctl.Name, "chart", vbTextCompare) > 0 Then
450   Debug.Print "FOUND: Report [" & strObjectName & "] - Control: " & ctl.Name
460                       strResults = strResults & "Report: " & strObjectName & " - Control: " & ctl.Name & vbCrLf
470                       intChartsFound = intChartsFound + 1
480                   End If
490                   On Error GoTo ErrorHandler
500               ElseIf ctl.ControlType = 143 Then
510   Debug.Print "FOUND: Report [" & strObjectName & "] - Control: " & ctl.Name & " (Modern Chart)"
520                   strResults = strResults & "Report: " & strObjectName & " - Control: " & ctl.Name & " (Modern Chart)" & vbCrLf
530                   intChartsFound = intChartsFound + 1
540               End If
550           Next ctl
              
              ' Close the report
560           DoCmd.Close acReport, strObjectName, acSaveNo
570       Next doc
          
          ' Display results
580   Debug.Print vbCrLf & "=== Summary ==="
590   Debug.Print "Total Modern Chart controls found: " & intChartsFound
          
600       If intChartsFound > 0 Then
610           strResults = strResults & vbCrLf & "Total: " & intChartsFound & " Modern Chart control(s) found."
620           MsgBox strResults, vbExclamation, "Modern Chart Controls Found"
630       Else
640           MsgBox "No Modern Chart controls found in forms or reports.", vbInformation, "Search Complete"
650       End If
          
660       FindModernChartControls = True
670       Exit Function
          
ErrorHandler:
680   Debug.Print "Error: " & Err.Description & " (Object: " & strObjectName & ")"
690       On Error Resume Next
700       DoCmd.Close acForm, strObjectName, acSaveNo
710       DoCmd.Close acReport, strObjectName, acSaveNo
720       MsgBox "Error searching for chart controls: " & Err.Description, vbExclamation
730       FindModernChartControls = False
End Function

Sub FindCalculatedFields()
      Dim db As DAO.Database
      Dim tdf As DAO.TableDef
      Dim fld As DAO.Field
      Dim intCount As Integer
          
100       Set db = CurrentDb
110       intCount = 0
          
120   Debug.Print "=== CALCULATED FIELDS REPORT ==="
130   Debug.Print "Generated: " & Now()
140   Debug.Print String(50, "=")
          
          ' Loop through all tables
150       For Each tdf In db.TableDefs
              ' Skip system and temporary tables
160           If Left(tdf.Name, 4) <> "MSys" And Left(tdf.Name, 1) <> "~" Then
                  
                  ' Loop through all fields in the table
170               For Each fld In tdf.Fields
                      ' Check if field is calculated (Type = 130 or dbCalculatedField)
180                   If fld.type = 130 Then
190                       intCount = intCount + 1
200   Debug.Print ""
210   Debug.Print "Table: " & tdf.Name
220   Debug.Print "  Field: " & fld.Name
230   Debug.Print "  Expression: " & fld.Properties("Expression").Value
240   Debug.Print "  Result Type: " & GetFieldTypeName(fld.Properties("ResultType").Value)
250                   End If
260               Next fld
                  
270           End If
280       Next tdf
          
290   Debug.Print ""
300   Debug.Print String(50, "=")
310   Debug.Print "Total Calculated Fields Found: " & intCount
320   Debug.Print String(50, "=")
          
330       If intCount = 0 Then
340           MsgBox "No calculated fields found in this database.", vbInformation
350       Else
360           MsgBox "Found " & intCount & " calculated field(s). See Immediate Window (Ctrl+G) for details.", vbInformation
370       End If
          
380       Set fld = Nothing
390       Set tdf = Nothing
400       Set db = Nothing
End Sub

Function GetFieldTypeName(intType As Integer) As String
    ' Returns friendly name for field result type
    Select Case intType
        Case 1: GetFieldTypeName = "Boolean"
        Case 2: GetFieldTypeName = "Byte"
        Case 3: GetFieldTypeName = "Integer"
        Case 4: GetFieldTypeName = "Long"
        Case 5: GetFieldTypeName = "Currency"
        Case 6: GetFieldTypeName = "Single"
        Case 7: GetFieldTypeName = "Double"
        Case 8: GetFieldTypeName = "Date/Time"
        Case 10: GetFieldTypeName = "Short Text"
        Case 12: GetFieldTypeName = "Long Text/Memo"
        Case Else: GetFieldTypeName = "Type " & intType
    End Select
End Function
