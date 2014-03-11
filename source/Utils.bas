Attribute VB_Name = "modUtils"
Option Explicit

Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const ERROR_SUCCESS = 0&
Private Const ERROR_FILE_NOT_FOUND = 2&

Private Const KEY_CREATE_LINK = &H20
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2

Private Const KEY_ALL_ACCESS = KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_SUB_KEY Or KEY_CREATE_LINK Or KEY_SET_VALUE
    'Declares
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Type utErrObjectInfo
        Description As String
        Number As Long
        HelpContext As Long
        HelpFile As String
        Source As String
End Type

Private oSavedErr As utErrObjectInfo

Public Function GetConnectionString(ByRef rsServerName As String, ByRef rsDBName As String, ByRef rsLogin As String, ByRef rsPassword As String) As String

    Dim cDatabaseServerName As String
    Dim cDefaultDatabaseName As String
    Dim cUserId As String
    Dim cPassword As String
    
    'Get the stuff using API registry Manipulation
    'Gets the necessary info from the Windows Registry
    'Gets data for ADO Connections...

    If Not GetLamdaRegSettings("Lamda Server Name", cDatabaseServerName) Then
        Err.Raise 1000, , "Error retrieving registry setting for :Lamda Server Name"
    Else
        rsServerName = cDatabaseServerName
    End If
    
    If Not GetLamdaRegSettings("Lamda Database Name", cDefaultDatabaseName) Then
        Err.Raise 1000, , "Error retrieving registry setting for :Lamda Database Name"
    Else
        rsDBName = cDefaultDatabaseName
    End If
        
    If Not GetLamdaRegSettings("Lamda UserID", cUserId) Then
        Err.Raise 1000, , "Error retrieving registry setting for :Lamda UserID"
    Else
        rsLogin = cUserId
    End If
    
    If Not GetLamdaRegSettings("Lamda Password", cPassword) Then
        Err.Raise 1000, , "Error retrieving registry setting for :Lamda Password"
    Else
        rsPassword = cPassword
    End If
    
    GetConnectionString = MakeConnectString(rsServerName, rsDBName, rsLogin, rsPassword)

End Function


Public Function MakeConnectString(ByVal vsServerName As String, ByVal vsDBName As String, ByVal vsLogin As String, ByVal vsPassword As String) As String
    MakeConnectString = "driver={SQL Server};server=" & vsServerName & _
                        ";database=" & vsDBName & ";uid=" & vsLogin & ";pwd=" & vsPassword
End Function

Private Function GetLamdaRegSettings(ByVal cReqString As String, ByRef cGetString As String) As Boolean
Dim bSuccess As Boolean
Dim bfound As Boolean
Dim lOpenKeyResult As Long
Dim lType As Long
Dim cMyData As String
Dim lLength As Long
Dim cKey As String
Dim cErrorMessage As String
On Error GoTo ErrorTrap
cKey = "SOFTWARE\VB AND VBA PROGRAM SETTINGS\MARLBOROUGH STIRLING\LAMDA"
    bSuccess = OpenKey(HKEY_LOCAL_MACHINE, cKey, bfound, lOpenKeyResult)
    If bSuccess Then
        If bfound Then
        'Get Setting
            bSuccess = QueryValue(lOpenKeyResult, cReqString, lType, cMyData, lLength, bfound)
            If bSuccess Then
                If bfound Then
                    cGetString = Left(cMyData, lLength - 1)
                    GetLamdaRegSettings = True
                Else
                    cErrorMessage = "Error ; Unable to locate registry value " & cReqString
                    GoTo ErrorTrap
                End If
            Else
                cErrorMessage = "Error ; Unable to open registry value " & cReqString
                GoTo ErrorTrap
            End If
        Else
            cErrorMessage = "Error ; Unable to locate registry key " & cKey
            GoTo ErrorTrap
        End If
    Else
        cErrorMessage = "Error ; Unable to open registry key " & cKey
        GoTo ErrorTrap
    End If
    If Not CloseKey(lOpenKeyResult) Then
        cErrorMessage = "Error Closing Key"
        GoTo ErrorTrap
    End If
Exit Function
ErrorTrap:
    GetLamdaRegSettings = False
End Function

Private Function OpenKey(ByVal lKey As Long, ByVal cSubKey As String, ByRef bfound As Boolean, ByRef lOpenKey As Long) As Boolean

    On Error GoTo ErrorHandler
    
    Dim lReturn As Long
    Dim bSuccess As Boolean
    
    lReturn = RegOpenKeyEx(lKey, cSubKey, 0&, KEY_QUERY_VALUE, lOpenKey)
    Select Case lReturn
        Case ERROR_SUCCESS
            bfound = True
            bSuccess = True
        Case ERROR_FILE_NOT_FOUND
            bfound = False
            bSuccess = True
        Case Else
            bfound = False
            bSuccess = False
    End Select
    OpenKey = bSuccess

Exit Function

ErrorHandler:
    bfound = False
    bSuccess = False
End Function

Private Function QueryValue(ByVal lKey As Long, ByVal cValueName As String, ByRef lType As Long, ByRef cData As String, ByRef lDataLength As Long, ByRef bfound As Boolean) As Boolean

    On Error GoTo ErrorHandler
    
    Dim lReturn As Long
    Dim bSuccess As Boolean
    
    lDataLength = 255
    cData = String$(lDataLength, 0)
    lReturn = RegQueryValueEx(lKey, cValueName, 0&, lType, cData, lDataLength)
    Select Case lReturn
        Case 0
            bfound = True
            bSuccess = True
        Case ERROR_FILE_NOT_FOUND
            bfound = False
            bSuccess = True
        Case Else
            bfound = False
            bSuccess = False
    End Select
    QueryValue = bSuccess

Exit Function

ErrorHandler:
    bfound = False
    bSuccess = False
End Function

Private Function CloseKey(ByVal lKey As Long) As Boolean

    On Error GoTo ErrorHandler
    
    Dim lReturn As Long
    Dim bSuccess As Boolean
    
    lReturn = RegCloseKey(lKey)
    Select Case lReturn
        Case ERROR_SUCCESS
            bSuccess = True
        Case Else
            bSuccess = False
    End Select
    CloseKey = bSuccess

Exit Function
ErrorHandler:
    bSuccess = False
End Function

Public Sub SelectAllInTextBox(ByVal voTextVox As TextBox)
    voTextVox.SelStart = 0
    If Len(voTextVox.Text) <> 0 Then
        voTextVox.SelLength = Len(voTextVox.Text)
    End If
End Sub

Public Function ReplaceString(ByVal vsInString As String, ByVal vsFindString As String, ByVal vsReplaceString As String) As String
    Dim lCurPos As Long
    Dim lFindPos As String
    Dim sOutString As String
    Dim lReplaceStringLen As Long
    Dim lFindStringLen As Long
    
    lCurPos = 1
    sOutString = vsInString
    lReplaceStringLen = Len(sOutString)
    lFindStringLen = Len(vsFindString)
    
    Do While lCurPos <= Len(sOutString)
        lFindPos = InStr(lCurPos, sOutString, vsFindString)
        If lFindPos <> 0 Then
            If lFindPos <> 1 Then
                sOutString = Mid$(sOutString, 1, lFindPos - 1) & vsReplaceString & Mid$(sOutString, lFindPos + lFindStringLen)
            Else
                sOutString = vsReplaceString & Mid$(sOutString, lFindPos + lFindStringLen)
            End If
            lCurPos = lFindPos + lFindStringLen + 1
        Else
            lCurPos = Len(sOutString) + 1
        End If
    Loop
    
    ReplaceString = sOutString
    
End Function

Public Sub SaveErrorObj()
    With oSavedErr
        .Description = Err.Description
        .Number = Err.Number
        .HelpContext = Err.HelpContext
        .HelpFile = Err.HelpFile
        .Source = Err.Source
    End With
End Sub

Public Sub ReRaiseError(Optional ByVal vboolUseSavedError As Boolean = False)
    If Not vboolUseSavedError Then
        With Err
            Err.Raise .Number, .Source, .Description, .HelpFile, .HelpContext
        End With
    Else
        With oSavedErr
            Err.Raise .Number, .Source, .Description, .HelpFile, .HelpContext
        End With
    End If
End Sub

Public Function GetCurUserLoginName(Optional vsDefaultLogin As String = "<no_login>") As String
    Dim sUserName As String
    Dim lSize As Long
    sUserName = Space$(255)
    lSize = Len(sUserName)
    Call GetUserName(sUserName, lSize)
    sUserName = Left(sUserName, lSize - 1)
    If sUserName = vbNullString Then
        sUserName = vsDefaultLogin
    End If
    GetCurUserLoginName = sUserName
End Function

