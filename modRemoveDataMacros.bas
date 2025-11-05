Attribute VB_Name = "modRemoveDataMacros"
'Public Function BackupAndRemoveAllDataMacros(Optional strBackupPath As String = "") As Boolean
'    On Error GoTo ErrorHandler
'
'    Dim db As DAO.Database
'    Dim rst As DAO.Recordset
'    Dim strSQL As String
'    Dim strTempFile As String
'    Dim strBackupFile As String
'    Dim intFileNum As Integer
'    Dim intMacrosRemoved As Integer
'
'    Set db = CurrentDb
'    intMacrosRemoved = 0
'
'    ' Set backup path
'    If strBackupPath = "" Then
'        strBackupPath = CurrentProject.Path & "\DataMacroBackups\"
'    End If
'
'    ' Create backup folder if it doesn't exist
'    If Dir(strBackupPath, vbDirectory) = "" Then
'        MkDir strBackupPath
'    End If
'
'    ' Create a temporary blank data macro XML file
'    strTempFile = Environ("TEMP") & "\BlankDataMacro.xml"
'
'    intFileNum = FreeFile
'    Open strTempFile For Output As intFileNum
'    Print #intFileNum, "<?xml version=""1.0"" encoding=""UTF-16""?>"
'    Print #intFileNum, "<DataMacros xmlns=""http://schemas.microsoft.com/office/accessservices/2009/04/application"">"
'    Print #intFileNum, "</DataMacros>"
'    Close #intFileNum
'
'    ' Find tables with data macros
'    strSQL = "SELECT [Name] FROM MSysObjects " & _
'             "WHERE Not IsNull(LvExtra) AND Type = 1 " & _
'             "ORDER BY [Name]"
'
'    Set rst = db.OpenRecordset(strSQL, dbOpenSnapshot)
'
'    Do While Not rst.EOF
'        Debug.Print "Processing: " & rst!Name
'
'        ' Backup the data macro first
'        strBackupFile = strBackupPath & rst!Name & "_DataMacro_" & _
'                        Format(Now(), "yyyymmdd_hhnnss") & ".xml"
'        Application.SaveAsText acTableDataMacro, rst!Name, strBackupFile
'        Debug.Print "  Backed up to: " & strBackupFile
'
'        ' Remove data macros by loading blank XML
'        Application.LoadFromText acTableDataMacro, rst!Name, strTempFile
'        Debug.Print "  Data macros removed"
'
'        intMacrosRemoved = intMacrosRemoved + 1
'        rst.MoveNext
'    Loop
'
'    rst.Close
'    Set rst = Nothing
'    Kill strTempFile
'
'    MsgBox "Successfully backed up and removed data macros from " & intMacrosRemoved & " tables." & vbCrLf & _
'           "Backups saved to: " & strBackupPath, vbInformation, "Data Macros Removed"
'
'    BackupAndRemoveAllDataMacros = True
'    Exit Function
'
'ErrorHandler:
'    Debug.Print "Error: " & Err.Description
'    MsgBox "Error: " & Err.Description, vbExclamation
'
'    On Error Resume Next
'    If Not rst Is Nothing Then rst.Close
'    If Dir(strTempFile) <> "" Then Kill strTempFile
'
'    BackupAndRemoveAllDataMacros = False
'End Function
'
Public Sub ListAllTableProperties()
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim prp As DAO.Property
    
    Set db = CurrentDb
    
    Debug.Print "=== ALL Table Properties ==="
    
    ' Check the tables you showed in the screenshot
    Dim arrTables As Variant
    arrTables = Array("Table1", "Table2", "Tables" [, "other tables"])
    
    Dim i As Integer
    For i = LBound(arrTables) To UBound(arrTables)
        On Error Resume Next
        Set tdf = db.TableDefs(arrTables(i))
        If Err.Number = 0 Then
            Debug.Print vbCrLf & "Table: " & tdf.Name
            Debug.Print String(50, "-")
            For Each prp In tdf.Properties
                Debug.Print "  Property: " & prp.Name & " | Type: " & prp.type
            Next prp
        End If
        On Error GoTo 0
    Next i
    
    Debug.Print vbCrLf & "=== End of List ==="
End Sub

Public Function FindModernChartControls() As Boolean
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim doc As Object
    Dim frm As Access.Form
    Dim rpt As Access.Report
    Dim ctl As Access.control
    Dim strObjectName As String
    Dim intChartsFound As Integer
    Dim strResults As String
    
    Set db = CurrentDb
    intChartsFound = 0
    strResults = "Modern Chart Controls Found:" & vbCrLf & vbCrLf
    
    ' Check all Forms
    Debug.Print "=== Checking Forms ==="
    For Each doc In CurrentProject.AllForms
        strObjectName = doc.Name
        
        ' Open form in design view (hidden)
        DoCmd.OpenForm strObjectName, acDesign, , , , acHidden
        Set frm = Forms(strObjectName)
        
        ' Check each control on the form
        For Each ctl In frm.Controls
            ' Modern Chart controls are typically Web Browser controls (type 128)
            ' or have ControlType of acWebBrowser
            If ctl.ControlType = acWebBrowser Then
                ' Check if it might be a chart by looking at properties
                On Error Resume Next
                If InStr(1, ctl.ControlSource, "chart", vbTextCompare) > 0 Or _
                   InStr(1, ctl.Name, "chart", vbTextCompare) > 0 Then
                    Debug.Print "FOUND: Form [" & strObjectName & "] - Control: " & ctl.Name
                    strResults = strResults & "Form: " & strObjectName & " - Control: " & ctl.Name & vbCrLf
                    intChartsFound = intChartsFound + 1
                End If
                On Error GoTo ErrorHandler
            ' Also check for Modern Chart specific control type (143)
            ElseIf ctl.ControlType = 143 Then
                Debug.Print "FOUND: Form [" & strObjectName & "] - Control: " & ctl.Name & " (Modern Chart)"
                strResults = strResults & "Form: " & strObjectName & " - Control: " & ctl.Name & " (Modern Chart)" & vbCrLf
                intChartsFound = intChartsFound + 1
            End If
        Next ctl
        
        ' Close the form
        DoCmd.Close acForm, strObjectName, acSaveNo
    Next doc
    
    ' Check all Reports
    Debug.Print vbCrLf & "=== Checking Reports ==="
    For Each doc In CurrentProject.AllReports
        strObjectName = doc.Name
        
        ' Open report in design view (hidden)
        DoCmd.OpenReport strObjectName, acViewDesign, , , acHidden
        Set rpt = Reports(strObjectName)
        
        ' Check each control on the report
        For Each ctl In rpt.Controls
            If ctl.ControlType = acWebBrowser Then
                On Error Resume Next
                If InStr(1, ctl.ControlSource, "chart", vbTextCompare) > 0 Or _
                   InStr(1, ctl.Name, "chart", vbTextCompare) > 0 Then
                    Debug.Print "FOUND: Report [" & strObjectName & "] - Control: " & ctl.Name
                    strResults = strResults & "Report: " & strObjectName & " - Control: " & ctl.Name & vbCrLf
                    intChartsFound = intChartsFound + 1
                End If
                On Error GoTo ErrorHandler
            ElseIf ctl.ControlType = 143 Then
                Debug.Print "FOUND: Report [" & strObjectName & "] - Control: " & ctl.Name & " (Modern Chart)"
                strResults = strResults & "Report: " & strObjectName & " - Control: " & ctl.Name & " (Modern Chart)" & vbCrLf
                intChartsFound = intChartsFound + 1
            End If
        Next ctl
        
        ' Close the report
        DoCmd.Close acReport, strObjectName, acSaveNo
    Next doc
    
    ' Display results
    Debug.Print vbCrLf & "=== Summary ==="
    Debug.Print "Total Modern Chart controls found: " & intChartsFound
    
    If intChartsFound > 0 Then
        strResults = strResults & vbCrLf & "Total: " & intChartsFound & " Modern Chart control(s) found."
        MsgBox strResults, vbExclamation, "Modern Chart Controls Found"
    Else
        MsgBox "No Modern Chart controls found in forms or reports.", vbInformation, "Search Complete"
    End If
    
    FindModernChartControls = True
    Exit Function
    
ErrorHandler:
    Debug.Print "Error: " & Err.Description & " (Object: " & strObjectName & ")"
    On Error Resume Next
    DoCmd.Close acForm, strObjectName, acSaveNo
    DoCmd.Close acReport, strObjectName, acSaveNo
    MsgBox "Error searching for chart controls: " & Err.Description, vbExclamation
    FindModernChartControls = False
End Function

Sub FindCalculatedFields()
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim strOutput As String
    Dim intCount As Integer
    
    Set db = CurrentDb
    intCount = 0
    
    Debug.Print "=== CALCULATED FIELDS REPORT ==="
    Debug.Print "Generated: " & Now()
    Debug.Print String(50, "=")
    
    ' Loop through all tables
    For Each tdf In db.TableDefs
        ' Skip system and temporary tables
        If Left(tdf.Name, 4) <> "MSys" And Left(tdf.Name, 1) <> "~" Then
            
            ' Loop through all fields in the table
            For Each fld In tdf.Fields
                ' Check if field is calculated (Type = 130 or dbCalculatedField)
                If fld.type = 130 Then
                    intCount = intCount + 1
                    Debug.Print ""
                    Debug.Print "Table: " & tdf.Name
                    Debug.Print "  Field: " & fld.Name
                    Debug.Print "  Expression: " & fld.Properties("Expression").Value
                    Debug.Print "  Result Type: " & GetFieldTypeName(fld.Properties("ResultType").Value)
                End If
            Next fld
            
        End If
    Next tdf
    
    Debug.Print ""
    Debug.Print String(50, "=")
    Debug.Print "Total Calculated Fields Found: " & intCount
    Debug.Print String(50, "=")
    
    If intCount = 0 Then
        MsgBox "No calculated fields found in this database.", vbInformation
    Else
        MsgBox "Found " & intCount & " calculated field(s). See Immediate Window (Ctrl+G) for details.", vbInformation
    End If
    
    Set fld = Nothing
    Set tdf = Nothing
    Set db = Nothing
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
        Case 10: GetFieldTypeName = "Text"
        Case 12: GetFieldTypeName = "Memo"
        Case Else: GetFieldTypeName = "Type " & intType
    End Select
End Function

