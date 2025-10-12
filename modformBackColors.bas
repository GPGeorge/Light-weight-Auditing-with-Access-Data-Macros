Attribute VB_Name = "modformBackColors"
Option Compare Database
Option Explicit
Public varBackColor As Variant
 

Public Function DevBackColor(ByRef frm As Form, ByVal lngBackColor As Long)

    On Error Resume Next
    
    frm.Detail.backColor = lngBackColor
    frm.Detail.AlternateBackColor = lngBackColor
    frm.FormHeader.backColor = lngBackColor
    frm.FormFooter.backColor = lngBackColor
 
End Function


