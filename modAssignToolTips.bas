Attribute VB_Name = "modAssignToolTips"
Option Compare Database
Option Explicit


Public Sub AssignToolTips(strToolTipBackup As String)

100       On Error GoTo errHandler

      Dim db As DAO.Database
      Dim rsToolTips As DAO.Recordset
      Dim strCurrentForm As String
      Dim strPreviousForm As String
      Dim strSQL As String
      Dim frm As Form
      Dim ctl As control
      Dim strControlName As String
      Dim strTipText As String
          
110       Set db = CurrentDb
       
120       strSQL = "SELECT FormName, ControlName, TipText FROM " & strToolTipBackup & " " & _
              " WHERE TipText Is Not Null AND TipText <> '' " & _
              " ORDER BY FormName, ControlName"
          
130       Set rsToolTips = db.OpenRecordset(strSQL, dbOpenSnapshot)
          
140       If rsToolTips.EOF Then
150           MsgBox "No tooltip data found to process.", vbInformation
160           GoTo Cleanup
170       End If
          
180       strPreviousForm = ""
          
190       Do While Not rsToolTips.EOF
200           strCurrentForm = rsToolTips!formName
       
210           If strCurrentForm <> strPreviousForm Then
220               If strPreviousForm <> "" Then
230                   DoCmd.Close acForm, strPreviousForm, acSaveYes
240   Debug.Print "Completed processing form: " & strPreviousForm
250               End If

260               DoCmd.OpenForm strCurrentForm, acDesign
270               Set frm = Forms(strCurrentForm)
280   Debug.Print "Processing form: " & strCurrentForm
                  
290               strPreviousForm = strCurrentForm
300           End If
              
310           strControlName = rsToolTips!controlName
320           strTipText = Nz(rsToolTips!tipText, "")
       
330           If Len(strTipText) > 0 Then
340               On Error Resume Next
350               Set ctl = frm.Controls(strControlName)
                  
360               If Err.Number = 0 Then
370                   ctl.ControlTipText = strTipText
380   Debug.Print "  - Set tooltip for: " & strControlName
390               Else

400   Debug.Print "  - WARNING: Control not found: " & strControlName
410                   Err.Clear
420               End If
430               On Error GoTo errHandler
440           End If
              
450           rsToolTips.MoveNext
460       Loop
          
470       If strPreviousForm <> "" Then
480           DoCmd.Close acForm, strPreviousForm, acSaveYes
490   Debug.Print "Completed processing form: " & strPreviousForm
500       End If
          
510       MsgBox "Tooltip assignment completed successfully!", vbInformation
          
Cleanup:
520       On Error Resume Next
530       If Not rsToolTips Is Nothing Then
540           rsToolTips.Close
550           Set rsToolTips = Nothing
560       End If
570       Set db = Nothing
580       Exit Sub
          

errHandler:
590       Call GlblErrMsg( _
              sFrm:=Application.VBE.ActiveCodePane.CodeModule, _
              sCtl:="AssignToolTips")
600       Resume Cleanup

610       Resume
End Sub


