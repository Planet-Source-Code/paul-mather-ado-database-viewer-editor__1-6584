Attribute VB_Name = "basRegistry"
Option Explicit

Const DCP_AUTHN_LEVEL_DEFAULT = 0
Const DCP_AUTHN_LEVEL_NONE = 1
Const DCP_AUTHN_LEVEL_CONNECT = 2
Const DCP_AUTHN_LEVEL_CALL = 3
Const DCP_AUTHN_LEVEL_PKT = 4
Const DCP_AUTHN_LEVEL_PKT_INTEGRITY = 5
Const DCP_AUTHN_LEVEL_PKT_PRIVACY = 6

Public Const REG_NONE = 0                       ' No value type
Public Const REG_SZ = 1                         ' Unicode nul terminated string
Public Const REG_EXPAND_SZ = 2                  ' Unicode nul terminated string
Public Const REG_BINARY = 3                     ' Free form binary
Public Const REG_DWORD = 4                      ' 32-bit number
Public Const REG_DWORD_LITTLE_ENDIAN = 4        ' 32-bit number (same as REG_DWORD)
Public Const REG_DWORD_BIG_ENDIAN = 5           ' 32-bit number
Public Const REG_LINK = 6                       ' Symbolic Link (unicode)
Public Const REG_MULTI_SZ = 7                   ' Multiple Unicode strings
Public Const REG_RESOURCE_LIST = 8              ' Resource list in the resource map
Public Const REG_FULL_RESOURCE_DESCRIPTOR = 9   ' Resource list in the hardware description
Public Const REG_RESOURCE_REQUIREMENTS_LIST = 10

Public Enum hKeyNames
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_CURRENT_CONFIG = &H80000005
End Enum

Public Const ERROR_SUCCESS = 0&
Public Const ERROR_NONE = 0
Public Const ERROR_BADDB = 1
Public Const ERROR_BADKEY = 2
Public Const ERROR_CANTOPEN = 3
Public Const ERROR_CANTREAD = 4
Public Const ERROR_CANTWRITE = 5
Public Const ERROR_OUTOFMEMORY = 6
Public Const ERROR_ARENA_TRASHED = 7
Public Const ERROR_ACCESS_DENIED = 8
Public Const ERROR_INVALID_PARAMETERS = 87
Public Const ERROR_NO_MORE_ITEMS = 259

Private Const SYNCHRONIZE = &H100000
Private Const STANDARD_RIGHTS_READ = &H20000
Private Const STANDARD_RIGHTS_WRITE = &H20000
Private Const STANDARD_RIGHTS_EXECUTE = &H20000
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SEDataValue = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))

Public Const KEY_ALL_ACCESS = &H3F

Public Const REG_OPTION_NON_VOLATILE = 0

'INI File Functions
Declare Function WritePrivateProfileString Lib "KERNEL32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString Lib "KERNEL32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileInt Lib "KERNEL32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long

'Registry Functions
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long

'Environment Functions
Declare Function SetEnvironmentVariable Lib "KERNEL32" Alias "SetEnvironmentVariableA" (ByVal lpName As String, ByVal lpValue As String) As Long
Declare Function GetEnvironmentVariable Lib "KERNEL32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long


'Registry Functions
Public Sub EnumRegKeys(ByRef returnName As Collection, Optional ByRef returnSubs As Collection, Optional hKeyName As String = "HKEY_CURRENT_USER", Optional KeyName As String = "SOFTWARE", Optional ByVal checkForSubs As Boolean = False)
    Dim lRetVal As Long      'result of the API functions
    Dim lngResult2 As Long      'result of the API functions
    Dim hKey2 As Long
    Dim hKey As Long         'handle of opened key
    Dim vValue As Variant    'setting of queried value
    Dim lngKeyHandle As Long
    Dim lngResult As Long
    Dim lngCurIdx As Long
    Dim strValue As String
    Dim lngValueLen As Long
    Dim lngData As Long
    Dim lngDataLen As Long
    Dim strResult As String
    Dim lKeyName As Long
    Dim SubLevel As Boolean

    Set returnName = New Collection
    Set returnSubs = New Collection
    
    KeyName = CompileKeyString(KeyName)
    
    lKeyName = resolveHkeyLong(hKeyName)
    
    Do
        lRetVal = RegOpenKeyEx(lKeyName, KeyName, 0, KEY_READ, hKey)
        lngValueLen = 2000
        strValue = String(lngValueLen, 0)
        lngDataLen = 2000
        lngResult = RegEnumKey(hKey, lngCurIdx, ByVal strValue, lngValueLen)
        lngCurIdx = lngCurIdx + 1
        RegCloseKey (hKey)
        
        If lngResult = ERROR_SUCCESS Then
            strResult = Left(strValue, lngValueLen)
            If InStr(1, strResult, Chr(0) & Chr(0) & Chr(0) & Chr(0), vbTextCompare) <> 0 Then
                strResult = Mid(strResult, 1, InStr(1, strResult, Chr(0) & Chr(0) & Chr(0) & Chr(0), vbTextCompare) - 1)
            Else
                strResult = strResult
            End If
            If checkForSubs = True Then
                If KeyName = "" Then
                    lngResult2 = RegOpenKeyEx(lKeyName, strResult, 0, KEY_READ, hKey2)
                Else
                    lngResult2 = RegOpenKeyEx(lKeyName, KeyName & "\" & strResult, 0, KEY_READ, hKey2)
                End If
                strValue = String(lngValueLen, 0)
                lngResult2 = RegEnumKey(hKey2, 0, ByVal strValue, lngValueLen)
                RegCloseKey (hKey2)
                If lngResult2 = ERROR_SUCCESS Then
                    SubLevel = True
                Else
                    SubLevel = False
                End If
                returnSubs.Add SubLevel
            End If
            returnName.Add strResult
        End If
    Loop While lngResult = ERROR_SUCCESS
    

End Sub

Public Sub EnumRegValues(ByRef returnName As Collection, Optional ByRef returnData As Collection, Optional ByRef returnType As Collection, Optional hKeyName As String = "HKEY_CURRENT_USER", Optional KeyName As String = "SOFTWARE")
    Dim lRetVal As Long      'result of the API functions
    Dim hKey As Long         'handle of opened key
    Dim hKey2 As Long
    Dim vValue As Variant    'setting of queried value
    Dim Count As Integer
    Dim lngKeyHandle As Long
    Dim lngResult As Long
    Dim lngCurIdx As Long
    Dim strValue As String
    Dim lngValueLen As Long
    Dim lngData As Long
    Dim lngDataLen As Long
    Dim strResult As String
    Dim lKeyName As Long
    Dim retData As String
    Dim retType As Long

    lKeyName = resolveHkeyLong(hKeyName)

    Set returnName = New Collection
    Set returnData = New Collection
    Set returnType = New Collection
    
    KeyName = CompileKeyString(KeyName)
    
    lRetVal = RegOpenKeyEx(lKeyName, KeyName, 0, KEY_READ, hKey)

    Do
        lngValueLen = 2000
        strValue = String(lngValueLen, 0)
        lngDataLen = 2000
        lngResult = RegEnumValue(hKey, lngCurIdx, ByVal strValue, lngValueLen, 0&, REG_DWORD, ByVal lngData, lngDataLen)
        lngCurIdx = lngCurIdx + 1
        If lngResult = ERROR_SUCCESS Then
            strResult = Left(strValue, lngValueLen)
            Call returnName.Add(strResult)
            Call RegOpenKeyEx(lKeyName, KeyName, 0, KEY_ALL_ACCESS, hKey2)
            Call QueryValueEx(hKey2, strResult, retData, retType)
            Call RegCloseKey(hKey2)
            Call returnData.Add(retData)
            Call returnType.Add(retType)
        End If
    Loop While lngResult = ERROR_SUCCESS

    RegCloseKey (hKey)

End Sub

Public Function GetSetting(appName As String, Section As String, Key As String, Optional Default As String, Optional hKeyName As hKeyNames = HKEY_CURRENT_USER, Optional AppNameHeader As String = "SOFTWARE")

Dim lRetVal As Long      'result of the API functions
Dim hKey As Long         'handle of opened key
Dim vValue As Variant    'setting of queried value
Dim keyString As String

    On Error GoTo e_Trap
    
    keyString = CompileKeyString(AppNameHeader, appName, Section)

    lRetVal = RegOpenKeyEx(hKeyName, keyString, 0, KEY_ALL_ACCESS, hKey)
    lRetVal = QueryValueEx(hKey, Key, vValue)
    If IsEmpty(vValue) Or vValue = "" Then
        vValue = Default
    End If
    GetSetting = vValue
    RegCloseKey (hKey)
    Exit Function
e_Trap:
    vValue = Default
    Exit Function
End Function
Public Function SaveSetting(appName As String, Section As String, Key As String, Setting As String, Optional hKeyName As hKeyNames = HKEY_CURRENT_USER, Optional AppNameHeader As String = "SOFTWARE") As Boolean

Dim lRetVal As Long       'result of the SetValueEx function
Dim hKey As Long          'handle of open key
Dim keyString As String

    On Error GoTo e_Trap
    
    keyString = CompileKeyString(AppNameHeader, appName, Section)

    lRetVal = RegCreateKeyEx(hKeyName, keyString, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, lRetVal)
    lRetVal = SetValueEx(hKey, Key, REG_SZ, Setting)
    RegCloseKey (hKey)
    SaveSetting = True
    Exit Function
e_Trap:
    SaveSetting = False
    Exit Function
End Function
Public Function DeleteSetting(appName As String, Optional Section As String, Optional Key As String, Optional hKeyName As hKeyNames = HKEY_CURRENT_USER, Optional AppNameHeader As String = "SOFTWARE", Optional recurseSubs As Boolean = True) As Boolean

Dim hNewKey As Long       'handle to the new key
Dim lRetVal As Long       'result of the SetValueEx function
Dim hKey As Long          'handle of open key
Dim keyString As String
Dim returnName As Collection
Dim returnSubs As Collection
Dim Count As Integer

    On Error GoTo e_Trap
    
    keyString = CompileKeyString(AppNameHeader, appName, Section)
    
    If Key <> "" Then
        lRetVal = RegCreateKeyEx(hKeyName, keyString, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, lRetVal)
        lRetVal = RegDeleteValue(hKey, Key)
        RegCloseKey (hKey)
    Else
        lRetVal = RegDeleteKey(hKeyName, keyString)
        If lRetVal = ERROR_CANTWRITE Then
            Call EnumRegKeys(returnName, returnSubs, resolveHkeyString(hKeyName), keyString)
            For Count = 1 To returnName.Count
                Call DeleteSetting(keyString & "\" & returnName(Count), "", "", hKeyName, "")
            Next Count
            lRetVal = RegDeleteKey(hKeyName, keyString)
        End If
    End If
    If lRetVal = ERROR_SUCCESS Then
        DeleteSetting = True
    Else
        DeleteSetting = False
    End If
    Exit Function
e_Trap:
    DeleteSetting = False
    Exit Function
End Function

Public Function AssociateFileType(extension As String, Optional useNotepadToEdit As Boolean = True, Optional appName As String, Optional filePath As String, Optional setDefault As Boolean = False) As Boolean
Dim lRetVal As Long       'result of the SetValueEx function
Dim hKey As Long          'handle of open key
Dim appPath As String
Dim appTitle As String
Dim commandString As String
Dim appKey As String

    On Error GoTo e_Trap
    
    If filePath = "" Then
        If Mid(App.Path, Len(App.Path) - 1, 1) = "\" Then
            appPath = App.Path & App.EXEName & ".exe"
        Else
            appPath = App.Path & "\" & App.EXEName & ".exe"
        End If
    Else
        appPath = filePath
    End If
    
    appPath = Replace(appPath, "\\", "\")
    
    If appName = "" Then
        appTitle = App.Title
    Else
        appTitle = appName
    End If
    
    If GetSetting("." & LCase(extension), "", "", appTitle, HKEY_CLASSES_ROOT, "") = appTitle Then
        setDefault = True
    End If
    
    If setDefault = True Then
        lRetVal = RegCreateKeyEx(HKEY_CLASSES_ROOT, appTitle, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, lRetVal)
        lRetVal = SetValueEx(hKey, "", REG_SZ, appTitle)
        RegCloseKey (hKey)
    
        lRetVal = RegCreateKeyEx(HKEY_CLASSES_ROOT, "." & LCase(extension), 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, lRetVal)
        lRetVal = SetValueEx(hKey, "", REG_SZ, appTitle)
        RegCloseKey (hKey)
    End If
    
    If setDefault = False Then
        If GetSetting("." & LCase(extension), "", "", "", HKEY_CLASSES_ROOT, "") <> "" Then
            appKey = GetSetting("." & LCase(extension), "", "", "", HKEY_CLASSES_ROOT, "")
        Else
            appKey = appTitle
        End If
        commandString = appKey & "\shell\Open2"
    Else
        appKey = appTitle
        commandString = appTitle & "\shell\Open"
    End If
    
    lRetVal = RegCreateKeyEx(HKEY_CLASSES_ROOT, commandString & "\command", 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, lRetVal)
    lRetVal = SetValueEx(hKey, "", REG_SZ, """" & appPath & """ %1")
    RegCloseKey (hKey)
    
    If appTitle <> "" Then
        lRetVal = RegCreateKeyEx(HKEY_CLASSES_ROOT, commandString, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, lRetVal)
        lRetVal = SetValueEx(hKey, "", REG_SZ, "Open with " & appTitle)
        RegCloseKey (hKey)
    End If
    
    If useNotepadToEdit = True Then
        lRetVal = RegCreateKeyEx(HKEY_CLASSES_ROOT, appKey & "\shell\Edit\command", 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, lRetVal)
        lRetVal = SetValueEx(hKey, "", REG_SZ, "notepad.exe %1")
        RegCloseKey (hKey)
    ElseIf GetSetting(appTitle & "\shell\Edit", "command", "", "", HKEY_CLASSES_ROOT, "") <> "" Then
        Call DeleteSetting(appTitle & "\shell", "Edit", "", HKEY_CLASSES_ROOT, "", True)
    End If
    
    If setDefault = True Then
        lRetVal = RegCreateKeyEx(HKEY_CLASSES_ROOT, appKey & "\DefaultIcon", 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, lRetVal)
        lRetVal = SetValueEx(hKey, "", REG_SZ, appPath)
        RegCloseKey (hKey)
    End If

    AssociateFileType = True
    Exit Function
e_Trap:
    AssociateFileType = False
    Exit Function

End Function

Public Sub CreateRunOnStartup(Optional ByVal appTitle As String, Optional ByVal appPath As String, Optional ByVal commandLine As String, Optional ByVal hKeyName As hKeyNames = HKEY_CURRENT_USER)
    If commandLine <> "" Then
        commandLine = " " & commandLine
    End If
    If appTitle = "" Then
        appTitle = App.Title
    End If
    If appPath = "" Then
        appPath = App.Path & "\" & App.EXEName & ".exe"
    End If
    Call SaveSetting("CurrentVersion", "Run", appTitle, appPath & commandLine, hKeyName, "Software\Microsoft\Windows")
End Sub
Public Sub DeleteRunOnStartup(Optional ByVal appTitle As String, Optional hKeyName As hKeyNames = HKEY_CURRENT_USER)
    Call DeleteSetting("CurrentVersion", "Run", appTitle, hKeyName, "Software\Microsoft\Windows")
End Sub

Public Sub SetDcomComputer(RemoteServerClassName As String, RemoteComputerName As String, Optional runLocal As Boolean = False, Optional UserName As String, Optional Password As String)
Dim defaultPath As String
Dim CLSID As String
'Dim dcomObj As Object

    CLSID = GetSetting(RemoteServerClassName, "Clsid", "", "", HKEY_CLASSES_ROOT, "")
    If CLSID <> "" Then
        If GetSetting(CLSID, "", "", "", HKEY_LOCAL_MACHINE, "SOFTWARE\Classes\CLSID") = RemoteServerClassName Then
            If runLocal = False Then
                If GetSetting(CLSID, "_LocalServer32", "", "", HKEY_LOCAL_MACHINE, "SOFTWARE\Classes\CLSID") = "" Then
                    defaultPath = GetSetting(CLSID, "LocalServer32", "", "", HKEY_LOCAL_MACHINE, "SOFTWARE\Classes\CLSID")
                End If
                Call SaveSetting("", CLSID, "RemoteServerName", RemoteComputerName, HKEY_LOCAL_MACHINE, "SOFTWARE\Classes\AppID")
                Call DeleteSetting(CLSID, "LocalServer32", "", HKEY_LOCAL_MACHINE, "SOFTWARE\Classes\CLSID")
                If defaultPath <> "" Then
                    Call SaveSetting(CLSID, "_LocalServer32", "", defaultPath, HKEY_LOCAL_MACHINE, "SOFTWARE\Classes\CLSID")
                End If
            Else
                If GetSetting(CLSID, "LocalServer32", "", "", HKEY_LOCAL_MACHINE, "SOFTWARE\Classes\CLSID") = "" Then
                    defaultPath = GetSetting(CLSID, "_LocalServer32", "", "", HKEY_LOCAL_MACHINE, "SOFTWARE\Classes\CLSID")
                End If
                Call SaveSetting(CLSID, "", "AppID", CLSID, HKEY_LOCAL_MACHINE, "SOFTWARE\Classes\CLSID")
                
'                Set dcomObj = CreateObject("DcomPerm")
'                If Trim(UserName) <> "" And OperatingSystemVersion = WindowsNT Then
'                    Call dcomObj.SetRunAs(RemoteServerClassName, Trim(UserName), Trim(password))
'                End If
'
'                Call dcomObj.SetAuthenticationLevel(RemoteServerClassName, DCP_AUTHN_LEVEL_NONE)
                
                Call DeleteSetting("AppID", CLSID, "RemoteServerName", HKEY_LOCAL_MACHINE, "SOFTWARE\Classes")
                Call DeleteSetting(CLSID, "_LocalServer32", "", HKEY_LOCAL_MACHINE, "SOFTWARE\Classes\CLSID")
                If defaultPath <> "" Then
                    Call SaveSetting(CLSID, "LocalServer32", "", defaultPath, HKEY_LOCAL_MACHINE, "SOFTWARE\Classes\CLSID")
                End If
            End If
        End If
    End If
            
End Sub
' INI Functions
Public Function GetIniInt(Section As String, Key As String, IniLocation As String, Optional Default As Long) As Long
    GetIniInt = GetPrivateProfileInt(Section, Key, Default, IniLocation)
End Function
Public Function GetIniString(Section As String, Key As String, IniLocation As String, Optional Default As String) As String
Dim ReturnValue As String * 128
Dim i, sLet
Dim iLen As Long
Dim Length As Long
        Length = GetPrivateProfileString(Section, Key, Default, ReturnValue, 128, IniLocation)
        i = InStr(1, Trim(ReturnValue), Chr(0))
        iLen = Len(Trim(ReturnValue))
        GetIniString = CStr(Left(Trim(ReturnValue), (i - 1)))
End Function
Public Function SaveIniString(Section As String, Key As String, Setting As String, IniLocation As String) As Long
    Setting = Replace(Setting, "[", "")
    Setting = Replace(Setting, "]", "")
    SaveIniString = WritePrivateProfileString(Section, Key, Setting, IniLocation)
End Function

Public Sub VerifyPath(pathString As String)
Dim CurrentPath As String

    pathString = Trim(pathString)
    If pathString = "" Then Exit Sub
    
    CurrentPath = Environ("PATH")
    If Mid(pathString, 1, 1) = ";" Then
        pathString = Mid(pathString, 2)
    End If
    If Mid(pathString, Len(pathString), 1) = ";" Then
        pathString = Mid(pathString, 1, Len(pathString) - 1)
    End If
    If InStr(1, UCase(CurrentPath), UCase(pathString), vbTextCompare) = 0 Then
        If Mid(CurrentPath, Len(CurrentPath), 1) = ";" Then
            Environ("PATH") = CurrentPath & pathString
        Else
            Environ("PATH") = CurrentPath & ";" & pathString
        End If
    End If
End Sub

Public Function resolveHkeyLong(hKeyName As String) As hKeyNames
    Select Case UCase(hKeyName)
        Case "HKEY_CURRENT_USER"
            resolveHkeyLong = HKEY_CURRENT_USER
        Case "HKEY_LOCAL_MACHINE"
            resolveHkeyLong = HKEY_LOCAL_MACHINE
        Case "HKEY_USERS"
            resolveHkeyLong = HKEY_USERS
        Case "HKEY_CLASSES_ROOT"
            resolveHkeyLong = HKEY_CLASSES_ROOT
        Case "HKEY_CURRENT_CONFIG"
            resolveHkeyLong = HKEY_CURRENT_CONFIG
    End Select
End Function
Public Function resolveHkeyString(hKeyName As hKeyNames) As String
    Select Case UCase(hKeyName)
        Case HKEY_CURRENT_USER
            resolveHkeyString = "HKEY_CURRENT_USER"
        Case HKEY_LOCAL_MACHINE
            resolveHkeyString = "HKEY_LOCAL_MACHINE"
        Case HKEY_USERS
            resolveHkeyString = "HKEY_USERS"
        Case HKEY_CLASSES_ROOT
            resolveHkeyString = "HKEY_CLASSES_ROOT"
        Case HKEY_CURRENT_CONFIG
            resolveHkeyString = "HKEY_CURRENT_CONFIG"
    End Select
End Function

' Private Functions
Private Function CompileKeyString(Optional AppNameHeader As String, Optional appName As String, Optional Section As String, Optional Key As String) As String
    If AppNameHeader <> "" Then
        CompileKeyString = AppNameHeader
    End If
    If appName <> "" Then
        If CompileKeyString <> "" Then
            CompileKeyString = CompileKeyString & "\"
        End If
        CompileKeyString = CompileKeyString & appName
    End If
    If Section <> "" Then
        If CompileKeyString <> "" Then
            CompileKeyString = CompileKeyString & "\"
        End If
        CompileKeyString = CompileKeyString & Section
    End If
    If Key <> "" Then
        If CompileKeyString <> "" Then
            CompileKeyString = CompileKeyString & "\"
        End If
        CompileKeyString = CompileKeyString & Key
    End If
    Do While InStr(1, CompileKeyString, "\\", vbTextCompare) <> 0
        If InStr(1, CompileKeyString, "\\", vbTextCompare) <> 0 Then
            CompileKeyString = Mid(CompileKeyString, 1, InStr(1, CompileKeyString, "\\", vbTextCompare) - 1) & Mid(CompileKeyString, InStr(1, CompileKeyString, "\\", vbTextCompare) + 1)
        End If
    Loop

    Do While InStr(1, CompileKeyString, "/", vbTextCompare) <> 0
        If InStr(1, CompileKeyString, "/", vbTextCompare) <> 0 Then
            CompileKeyString = Mid(CompileKeyString, 1, InStr(1, CompileKeyString, "/", vbTextCompare) - 1) & "\" & Mid(CompileKeyString, InStr(1, CompileKeyString, "/", vbTextCompare) + 1)
        End If
    Loop

End Function
Private Function SetValueEx(ByVal hKey As Long, sValueName As String, lType As Long, vValue As Variant) As Long
    Dim lValue As Long
    Dim sValue As String
    Select Case lType
        Case REG_SZ, REG_EXPAND_SZ
            sValue = vValue & Chr$(0)
            SetValueEx = RegSetValueExString(hKey, sValueName, 0&, lType, sValue, Len(sValue))
        Case REG_DWORD, REG_DWORD_BIG_ENDIAN
            lValue = vValue
            SetValueEx = RegSetValueExLong(hKey, sValueName, 0&, lType, lValue, 4)
        End Select
End Function

Private Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As String, vValue As Variant, Optional dataType As Long) As Long
    Dim cch As Long
    Dim lrc As Long
    Dim lType As Long
    Dim lValue As Long
    Dim sValue As String
    Dim Count As Integer
    Dim Holder As String
    Dim NewVal As String

    On Error GoTo QueryValueExError
    vValue = ""
    
    ' Determine the size and type of data to be read
    lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
    If lrc <> ERROR_NONE Then Error 5

    dataType = lType
    Select Case lType
        ' For strings
        Case REG_SZ, REG_EXPAND_SZ:
            sValue = String(cch, 0)
            lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)
            If lrc = ERROR_NONE Then
                vValue = Left$(sValue, cch - 1)
            Else
                vValue = Empty
            End If
        ' For DWORDS
        Case REG_DWORD, REG_DWORD_BIG_ENDIAN:
            lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, cch)
            If lrc = ERROR_NONE Then vValue = lValue
        Case REG_BINARY
            sValue = String(cch, 0)
            lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)
            If lrc = ERROR_NONE Then
                Holder = Left$(sValue, cch - 1)
                vValue = ""
                For Count = 1 To Len(Holder)
                    NewVal = Format(Hex(Asc(Mid(Holder, Count, 1))), "00")
                    If Len(NewVal) = 1 Then
                        NewVal = "0" & NewVal
                    End If
                    vValue = vValue & NewVal & " "
                Next Count
                vValue = Trim(vValue)
            Else
                vValue = Empty
            End If
            
        Case Else
            'all other data types not supported
            lrc = -1
    End Select

QueryValueExExit:
    QueryValueEx = lrc
    Exit Function
QueryValueExError:
    Resume QueryValueExExit
End Function


