Attribute VB_Name = "basGeneral"
Option Explicit

Public Const DEF_CUSTOM_SQL As String = "Custom SQL"
Public Const DEF_REGISTRY_CONNECTIONS As String = "Connections"
Public Const DEF_REGISTRY_SETTINGS As String = "Settings"

Public Const DEF_ORACLE_CLIENT As String = "Oracle (Needs Client)"
Public Const DEF_ORACLE_ODBC As String = "Oracle (ODBC)"
Public Const DEF_SQL_SERVER As String = "SQL Server"
Public Const DEF_DSN_FILE As String = "DSN File"
Public Const DEF_ACCESS As String = "Access"

Public Const DEF_ACCESS97_OLEDB As String = "3.51"
Public Const DEF_ACCESS2K_OLEDB As String = "4.0"

Public dbObj As ADODB.Connection
Public dbPath As String
Public dbConnectionString As String
Public dbType As e_DatabaseTypes
Public LastOpenedType As e_LastOpened

Public Enum e_LastOpened
    e_LastOpened_Access = 0
    e_LastOpened_Network
End Enum

Public Enum e_DatabaseTypes
    e_DatabaseTypes_Undefined = 0
    e_databaseTypes_OracleMSDA = 1
    e_databaseTypes_OracleODBC = 2
    e_databaseTypes_SQLserver = 3
    e_databaseTypes_MicrosoftJet = 4
    e_databaseTypes_MicrosoftAccess97File = 5
    e_databaseTypes_MicrosoftAccess2KFile = 6
    e_databaseTypes_DSNFile = 7
    e_databaseTypes_AccessFile = 99
End Enum

Public Declare Sub Sleep Lib "KERNEL32" (ByVal dwMilliseconds As Long)
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long
Public Sub LockWindow(ByVal hwnd As Long)
Dim lRet As Long
    lRet = LockWindowUpdate(hwnd)
End Sub
Public Sub ReleaseWindow()
Dim lRet As Long
    lRet = LockWindowUpdate(0)
End Sub

Public Function BuildConnectString(ByVal databaseType As e_DatabaseTypes, ByVal serverOrFilename As String, Optional ByVal databaseName As String, Optional ByVal UserName As String, Optional ByVal Password As String) As String
    Select Case databaseType
        Case e_databaseTypes_OracleMSDA
            BuildConnectString = "Provider=MSDAORA;Data Source=" & serverOrFilename & ";User ID=" & IIf(UserName <> "", UserName, "") & ";Password=" & IIf(Password <> "", Password, "") & ";" & IIf(databaseName <> "", "Initial Catalog=" & databaseName & ";", "")
        Case e_databaseTypes_OracleODBC
            BuildConnectString = "DRIVER={Microsoft ODBC for Oracle};SERVER=" & serverOrFilename & ";UID=" & UserName & ";PWD=" & Password & ";" & IIf(databaseName <> "", "Initial Catalog=" & databaseName & ";", "")
        Case e_databaseTypes_SQLserver
            BuildConnectString = "Provider=SQLOLEDB.1;Persist Security Info=False;Data Source=" & serverOrFilename & ";User ID=" & IIf(UserName <> "", UserName, "") & ";Password=" & IIf(Password <> "", Password, "") & ";" & IIf(databaseName <> "", "Initial Catalog=" & databaseName & ";", "")
        Case e_databaseTypes_DSNFile
            BuildConnectString = "Provider=MSDASQL;DSN=" & serverOrFilename & ";UID=" & IIf(UserName <> "", UserName, "") & ";PWD=" & IIf(Password <> "", Password & ";", "") & ";" & IIf(databaseName <> "", "Initial Catalog=" & databaseName & ";", "")
        Case e_databaseTypes_MicrosoftAccess2KFile, e_databaseTypes_MicrosoftAccess97File
            BuildConnectString = "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & serverOrFilename & ";DefaultDir=" & DetermineDirectory(serverOrFilename) & ";"
    End Select
End Function
Public Sub SaveDefaultTable()
    If frmMain.cmbTables.Text <> DEF_CUSTOM_SQL Then
        If LastOpenedType = e_LastOpened_Access Then
            Call SaveSetting(App.Title, DEF_REGISTRY_CONNECTIONS & "\" & DEF_ACCESS, "Default Table", frmMain.cmbTables.Text, HKEY_LOCAL_MACHINE, "SOFTWARE\Database")
        Else
            If dbType = e_databaseTypes_OracleMSDA Then
                Call SaveSetting(App.Title, DEF_REGISTRY_CONNECTIONS & "\" & DEF_ORACLE_CLIENT, "Default Table", frmMain.cmbTables.Text, HKEY_LOCAL_MACHINE, "SOFTWARE\Database")
            ElseIf dbType = e_databaseTypes_OracleODBC Then
                Call SaveSetting(App.Title, DEF_REGISTRY_CONNECTIONS & "\" & DEF_ORACLE_ODBC, "Default Table", frmMain.cmbTables.Text, HKEY_LOCAL_MACHINE, "SOFTWARE\Database")
            ElseIf dbType = e_databaseTypes_SQLserver Then
                Call SaveSetting(App.Title, DEF_REGISTRY_CONNECTIONS & "\" & DEF_SQL_SERVER, "Default Table", frmMain.cmbTables.Text, HKEY_LOCAL_MACHINE, "SOFTWARE\Database")
            ElseIf dbType = e_databaseTypes_DSNFile Then
                Call SaveSetting(App.Title, DEF_REGISTRY_CONNECTIONS & "\" & DEF_DSN_FILE, "Default Table", frmMain.cmbTables.Text, HKEY_LOCAL_MACHINE, "SOFTWARE\Database")
            End If
        End If
    End If
End Sub

Public Function ResolveTable(inputTable As String) As String
    ResolveTable = IIf(InStr(1, inputTable, " ") <> 0 Or IsNumeric(Left(inputTable, 1)), "[" & inputTable & "]", inputTable)
End Function
Public Function BuildDatabaseName(ByVal databaseType As e_DatabaseTypes, ByVal databasePath As String, ByVal lastDatabaseOpenType As e_LastOpened) As String
    If databaseType = e_databaseTypes_MicrosoftAccess2KFile Then
        BuildDatabaseName = "Access 2000: " & DetermineFilename(databasePath)
    ElseIf databaseType = e_databaseTypes_MicrosoftAccess97File Then
        BuildDatabaseName = "Access 97: " & DetermineFilename(databasePath)
    ElseIf databaseType = e_databaseTypes_OracleMSDA Or databaseType = e_databaseTypes_OracleODBC Then
        BuildDatabaseName = "Oracle: " & databasePath
    ElseIf databaseType = e_databaseTypes_SQLserver Then
        BuildDatabaseName = "SQL Server: " & databasePath
    ElseIf databaseType = e_databaseTypes_DSNFile Then
        BuildDatabaseName = "DSN Source: " & databasePath
    ElseIf databaseType = e_databaseTypes_AccessFile Then
        BuildDatabaseName = "Access Database: " & DetermineFilename(databasePath)
    End If
    If lastDatabaseOpenType = e_LastOpened_Network Then
        BuildDatabaseName = "Network - " & BuildDatabaseName
    End If
End Function

