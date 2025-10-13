Attribute VB_Name = "modGlobalConst"
Option Compare Database
Option Explicit

'George Hepworth
'Public Constants

Public blnSaveRecord As Boolean
Public Const intCharacterMatch As Integer = 2
Public Const SW_HIDE As Integer = 0
Public Const SW_SHOWNORMAL As Integer = 1
Public Const SW_SHOWMINIMIZED As Integer = 2
Public Const SW_SHOWMAXIMIZED As Integer = 3
Public Const SM_CXSCREEN As Integer = 0
Public Const SM_CYSCREEN As Integer = 1
    
Public strGroupBy As String
Public strHaving As String
Public strInsert As String
Public strOrderBy As String
Public strSELECT As String
Public strSQL As String
Public strWhere As String
    
Public strCaption As String
Public strMBTitle As String
Public strMonth As String
Public strDay As String
Public strYear As String

Public lngMBBtn As Long
  
Public intStartGROUPBY As Integer
Public intStartHAVING As Integer
Public intStartINSERT As Integer
Public intStartORDERBY As Integer
Public intStartSELECT As Integer
Public intStartWHERE As Integer
Public intStartUnion As Integer
    
Public intLenINSERT As Integer
Public intLenSELECT As Integer
Public intLenWHERE As Integer
Public intLenORDERBY As Integer
Public intLenGROUPBY As Integer
Public intLenHAVING As Integer
Public intLenSQL As Integer
'

Public Enum appcolor
    ctlForeClrHI = 12101016 'vbWhite
    ctlBackClrHi = 5000786 '14211288
    ctlForeClrOn = vbWhite
    ctlBackClrOn = 13434828 ' Light Green
    lblForeclrOn = 13434828 ' Light Green
    LimeGreen = 326589
    BurntRed = 990033
    Gray = 12632256   'Darker Gray
    vbGray = 14211288 'Lighter Gray
    BlinkAmnt = 18
    BlueBack = 15709952 'Orange
End Enum

Public Enum AppErr
    AccessDenied = 5
    actioncancelled = 2501
    BadInputMask = 2279
    CancelledbyUser = 3059
    CantDelete = 2046
    CantFindTable = 3078
    CantDeleteFile = 70
    CantGotoRecord = 2105
    CantOpenFile = 2220
    CantPerformOperation = 3032
    CantSetProp = 3267
    ControlNotFound = 2465
    DBExists = 3204
    DupeIxorPK = 3022
    FileExists = 58
    FileNotFound = 53
    FrmMustbeActive = 2475
    FormNotFound = 2102
    IMExNotFound = 3625
    InvalidFldVal = 2113
    InvalidFormRef = 2452
    InvalidOperation = 3219
    ItemnotFound = 3265
    MissingMcrTlbr = 2485
    MissingObject = 2450
    MissingOptr = 3075
    NoConnection = 3151
    NoControl = 2474
    NoCurrentRecord = 3021
    NoObject = 2467
    NoParent = 2452
    NoRemoteTable = vbObjectError + 2000
    NotSupportedonObjectType = 3251
    NumberInvalid = 2200
    ODBCCalledFailed = 3146
    ObjNotSet = 91
    PermissionDenied = 70
    PropNotFound = 3270
    QueryExists = 3012
    QdefExists = 3012
    RecordIsLocked = 3260
    SrchKeyNotFound = 3709
    TableExists = 3010
    TableLnkExists = 3012
    Tablelocked = 3211
    TooFewParam = 3061
    TypeMismatch = 13
    Unknown = 3316
    Unsupported = 438
    UserCancel = vbObjectError + 1000
    ViolateRIAdd = 3201
    ViolateRIDelete = 3200
    ViolateRIFrm = 3314
 
End Enum

Public Enum AppPrint
    PlainPrint = 1
    EmailHTML = 2
    SnapShot = 3
    PDF = 4
    acPrinter = 0
    acPreview = 2
    acXLSExport = 98
    acPDF = 99
End Enum

Public Const strDoubleLine As Variant = vbNewLine & vbNewLine
Public Const strSingleLine As Variant = vbNewLine

#If VBA7 Then
Private Declare PtrSafe Sub sapiSleep Lib "kernel32" _
    Alias "Sleep" _
    (ByVal dwMilliseconds As Long)
#Else
Private Declare Sub sapiSleep Lib "kernel32" _
    Alias "Sleep" _
    (ByVal dwMilliseconds As Long)
#End If

Public Function AppString() As String

    On Error GoTo ErrHandler:

    AppString = "tblAppString"
    TempVars.Add Name:="AppString", Value:="tblAppString"
    CurrentDb.OpenRecordset "Select Top 1 * FROM " & TempVars("AppString"), dbOpenSnapshot, dbSeeChanges
    '    CurrentDb.OpenRecordset "Select Top 1 * FROM " & AppString, dbOpenSnapshot, dbSeeChanges
    Exit Function
    
ErrHandler:
 
    If Err = AppErr.CantFindTable Then
        TempVars.Add Name:="AppString", Value:="UsysAppString"
        AppString = "UsysAppString"
    End If

End Function

Public Function GlblErrMsg( _
    Optional ByVal iLn As Integer, _
    Optional ByVal sFrm As String, _
    Optional ByVal sCtl As String) As Boolean
    
Dim appc As Appconstants
Dim strErrMessage As String
Dim strSupport As String
Dim errNum As Long
Dim strErrDesc As String
Dim intLN As Long
    
    If Err.Number = AppErr.actioncancelled Then Exit Function
    errNum = Nz(Err.Number, 9999)
    strErrDesc = IIf(Len(Err.description & "") = 0, " Unidentified Error", Err.description)
    intLN = IIf(Len(Erl & "") = 0, 0, Erl)
    sCtl = IIf(Len(sCtl & "") = 0, " Unidentified Procedure", sCtl)
    sFrm = IIf(Len(sFrm & "") = 0, " Environment", sFrm)
    
    Set appc = New_AppConstants
    GlblErrMsg = False
    strSupport = Replace(appc.SupportEmail, ",", vbNewLine)
    strMBTitle = appc.MBErr
    lngMBBtn = appc.MBYNBtn
    
    strErrMessage = "Please report this error to Your Support Person: " & appc.SupportPerson & "." & strDoubleLine & _
        "The Error Number was: """ & errNum & """." & strSingleLine & _
        "The Error Description was: """ & Nz(strErrDesc, " No Description") & """." & strDoubleLine & _
        "The error occurred at Line Number: " & intLN & strSingleLine & _
        "In procedure: """ & sCtl & """." & strSingleLine & _
        "In module: """ & sFrm & """." & strDoubleLine & _
        "Please submit this error report " & _
        "and a brief description " & "of what you were doing when the error occurred." & strDoubleLine
    On Error Resume Next


    GlblErrMsg = True
                    
    If MsgBox(Prompt:=strErrMessage & vbCrLf & vbCrLf & "Send report?", buttons:=lngMBBtn, Title:=strMBTitle & " Send Report?") = vbYes Then
        DoCmd.SendObject objecttype:=acSendNoObject, To:=strSupport, Subject:="Error Occurred in  " & appc.AppName, messagetext:=strErrMessage, templatefile:=False
    Else
        If Nz(appc.booDebug) Then Stop
    End If
End Function

Public Function RemoveSPPublish() As Boolean

100       On Error GoTo ErrHandler

      Dim appc As Appconstants
      Dim db As DAO.Database

110       Set appc = New_AppConstants
120       If Nz(appc.booDebug, 0) Then Stop
130       RemoveSPPublish = False
140       db.Properties.Delete ("PublishURL")
150       RemoveSPPublish = True
          
CleanUp:
          
160       On Error Resume Next
170       Exit Function

ErrHandler:
180       If Err <> AppErr.ItemnotFound And Err <> AppErr.ObjNotSet Then
190           Call GlblErrMsg( _
                  sFrm:=Application.vbe.ActiveCodePane.CodeModule, _
                  sCtl:="RemoveSPPublish" _
                  )
200       End If
210       Resume CleanUp

220       Resume
End Function

Public Sub Sleep(ByVal lngMilliSec As Long)

    If lngMilliSec > 0 Then
        Call sapiSleep(lngMilliSec)
    End If
    
End Sub

Public Sub testit()
100       On Error GoTo ErrHandler

      Dim db As DAO.Database
      Dim rst As DAO.Recordset
          
110       Set db = CurrentDb
120       Set rst = db.OpenRecordset("Select * from USysAppString", dbOpenSnapshot)
130       rst.MoveLast
140       rst.MoveFirst
150       Debug.Print rst.RecordCount / 0

CleanUp:
          
160       On Error Resume Next
170       rst.Close
180       Set rst = Nothing
190       db.Close
200       Set db = Nothing
220       Exit Sub

ErrHandler:
230       Call GlblErrMsg( _
              sFrm:=Application.vbe.ActiveCodePane.CodeModule, _
              sCtl:="testit" _
              )
240       Resume CleanUp

250       Resume
End Sub








