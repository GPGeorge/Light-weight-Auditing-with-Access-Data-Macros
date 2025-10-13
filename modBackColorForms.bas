Attribute VB_Name = "modBackColorForms"
Public Function SetFormSectionsAndLabels(lngNewColor As Long)
Dim obj As AccessObject
    
    For Each obj In CurrentProject.AllForms
        UpdateFormSectionsAndLabels obj.Name, lngNewColor
    Next obj
    
    MsgBox "Form sections and labels updated.", vbInformation
End Function

Private Sub UpdateFormSectionsAndLabels(strFormName As String, lngNewColor As Long)
    On Error GoTo errHandler
    
Dim frm As Form
Dim ctl As control
Dim s As Section
    
    ' Open form in design view (hidden)
    DoCmd.OpenForm strFormName, acDesign, , , , acHidden
    Set frm = Forms(strFormName)
    
    ' Update all sections
    For Each s In frm.Sections
        On Error Resume Next
        s.backColor = lngNewColor
        On Error GoTo errHandler
    Next s
    
    ' Update labels
    For Each ctl In frm.Controls
        If ctl.ControlType = acLabel Then
            On Error Resume Next
            ctl.BackStyle = 1      ' Opaque so color shows
            ctl.backColor = lngNewColor
            On Error GoTo errHandler
        End If
        
        ' Recurse into subforms
        If ctl.ControlType = acSubform Then
            If Len(ctl.SourceObject) > 0 And Left(ctl.SourceObject, 5) = "Form." Then
                UpdateFormSectionsAndLabels Mid(ctl.SourceObject, 6), lngNewColor
            End If
        End If
    Next ctl
    
    DoCmd.Close acForm, strFormName, acSaveYes
    Exit Sub

errHandler:
    Resume Next
End Sub

Public Function SetCommandButtons(lngButtonColor As Long, Optional lngTextColor As Long = -1)
Dim obj As AccessObject
    
    For Each obj In CurrentProject.AllForms
        UpdateCommandButtons obj.Name, lngButtonColor, lngTextColor
    Next obj
    
    MsgBox "Command buttons updated.", vbInformation
End Function

Private Sub UpdateCommandButtons(strFormName As String, lngButtonColor As Long, lngTextColor As Long)
    On Error GoTo errHandler
    
Dim frm As Form
Dim ctl As control
    
    ' Open form in design view (hidden)
    DoCmd.OpenForm strFormName, acDesign, , , , acHidden
    Set frm = Forms(strFormName)
    
    ' Update command buttons
    For Each ctl In frm.Controls
        If ctl.ControlType = acCommandButton Then
            On Error Resume Next
Debug.Print acSpecialEffectRaised
            ctl.SpecialEffect = acSpecialEffectRaised
            ctl.UseTheme = True           ' allow custom color
            ctl.backColor = lngButtonColor
            If lngTextColor <> -1 Then
                ctl.ForeColor = lngTextColor
            End If

            On Error GoTo errHandler
        End If
        
        ' Recurse into subforms
        If ctl.ControlType = acSubform Then
            If Len(ctl.SourceObject) > 0 And Left(ctl.SourceObject, 5) = "Form." Then
                UpdateCommandButtons Mid(ctl.SourceObject, 6), lngButtonColor, lngTextColor
            End If
        End If
    Next ctl
    
    DoCmd.Close acForm, strFormName, acSaveYes
    Exit Sub

errHandler:
    Resume Next
End Sub

