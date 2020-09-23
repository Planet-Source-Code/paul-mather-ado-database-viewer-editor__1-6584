VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMain 
   Caption         =   "Database Editor"
   ClientHeight    =   5865
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7725
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   7725
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid gridData 
      Bindings        =   "frmMain.frx":0442
      Height          =   3975
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   7011
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraSqlStatement 
      Caption         =   "SQL Statement"
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   4560
      Width           =   7455
      Begin VB.CommandButton cmdExecute 
         Caption         =   "Execute"
         Height          =   495
         Left            =   6240
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtSqlStatement 
         Height          =   495
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   6015
      End
   End
   Begin VB.CheckBox chkEditMode 
      Caption         =   "&Edit Mode"
      Height          =   255
      Left            =   6480
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc adoData 
      Height          =   330
      Left            =   4800
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      ConnectMode     =   3
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "adoData"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.ComboBox cmbTables 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin VB.Image imgLoading 
      Height          =   240
      Left            =   7320
      Picture         =   "frmMain.frx":0458
      Top             =   5520
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      Caption         =   "Status"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   5520
      Width           =   450
   End
   Begin VB.Label lblDatabaseName 
      AutoSize        =   -1  'True
      Caption         =   "Database Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2880
      TabIndex        =   4
      Top             =   120
      Width           =   1710
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpenAccess 
         Caption         =   "Open &Access Database"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuOpenNetwork 
         Caption         =   "Open &Network Database"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewColumns 
         Caption         =   "&View Column Types"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuSep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditMode 
         Caption         =   "&Edit Mode"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuShowSQL 
         Caption         =   "&Show SQL Statements"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSep11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClearPassword 
         Caption         =   "&Clear Password"
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnuReload 
         Caption         =   "&Reload Table List"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuPurgeDate 
         Caption         =   "&Purge by Date"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuCompact 
         Caption         =   "Compact Database"
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddTable 
         Caption         =   "Add &Table"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuRenameTable 
         Caption         =   "Re&name Table"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuDropTable 
         Caption         =   "Drop Ta&ble"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddColumn 
         Caption         =   "Add &Column"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuRenameColumn 
         Caption         =   "Rename Column"
         Shortcut        =   ^W
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDeleteColumn 
         Caption         =   "De&lete Column"
         Shortcut        =   ^L
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SortOrder As Boolean
Dim LockTextBox As Boolean
Dim LastColumnSort As Integer

Private Sub cmdExecute_Click()

    If Trim(txtSqlStatement) = "" Then Exit Sub
    On Error GoTo e_Trap
    imgLoading.Visible = True
    Me.Refresh
    SortOrder = True
    adoData.RecordSource = txtSqlStatement
    adoData.Refresh
    Set gridData.DataSource = adoData.Recordset
    lblStatus.caption = "Record Count: " & adoData.Recordset.RecordCount
    imgLoading.Visible = False
    Exit Sub
e_Trap:
    lblStatus.caption = "Error: " & Err.Description & " (" & Err.Number & ")"
    imgLoading.Visible = False
End Sub


Private Sub Form_Load()
Dim commandLine As String
Dim serverType As Integer
Dim serverName As String
Dim databaseName As String
Dim Password As String
Dim UserName As String
Dim defaultTable As String
Dim registryString As String
    
    Call Hook(Me.hwnd, 7000, 3500)
    
    Set dbObj = New ADODB.Connection
        
    mnuEdit.Enabled = False
    lblDatabaseName.caption = ""
    
    Me.Width = GetSetting(App.Title, DEF_REGISTRY_SETTINGS, "Form Width", WorkAreaWidth, HKEY_LOCAL_MACHINE, "SOFTWARE\Database")
    Me.Height = GetSetting(App.Title, DEF_REGISTRY_SETTINGS, "Form Height", WorkAreaHeight / 2, HKEY_LOCAL_MACHINE, "SOFTWARE\Database")
    
    Me.Top = GetSetting(App.Title, DEF_REGISTRY_SETTINGS, "Form Top", WorkAreaBottom - Me.Height, HKEY_LOCAL_MACHINE, "SOFTWARE\Database")
    If Me.Top < WorkAreaTop Then
        Me.Top = WorkAreaTop
    ElseIf Me.Top > WorkAreaBottom - Me.Height Then
        Me.Top = WorkAreaBottom - Me.Height
    End If
    If Me.Height > WorkAreaHeight Then
        Me.Height = WorkAreaHeight
    End If
    
    Me.Left = GetSetting(App.Title, DEF_REGISTRY_SETTINGS, "Form Left", WorkAreaLeft, HKEY_LOCAL_MACHINE, "SOFTWARE\Database")
    If Me.Left < WorkAreaLeft Then
        Me.Left = WorkAreaLeft
    ElseIf Me.Left > WorkAreaRight - Me.Width Then
        Me.Left = WorkAreaRight - Me.Width
    End If
    If Me.Width > WorkAreaWidth Then
        Me.Width = WorkAreaWidth
    End If
    
    Me.WindowState = GetSetting(App.Title, DEF_REGISTRY_SETTINGS, "WindowState", vbNormal, HKEY_LOCAL_MACHINE, "SOFTWARE\Database")
    
    mnuShowSQL.Checked = GetSetting(App.Title, DEF_REGISTRY_SETTINGS, "Show SQL", True, HKEY_LOCAL_MACHINE, "SOFTWARE\Database")
    Call Form_Resize
    
    On Error Resume Next
    Call AssociateFileType("mdb", False, App.Title)

    LastOpenedType = GetSetting(App.Title, DEF_REGISTRY_SETTINGS, "Last Opened Type", 0, HKEY_LOCAL_MACHINE, "SOFTWARE\Database")

    commandLine = Command
    If commandLine <> "" Then
        If Dir(commandLine) <> "" Then
            dbPath = commandLine
            dbConnectionString = BuildConnectString(e_databaseTypes_MicrosoftAccess2KFile, dbPath)
            dbType = e_databaseTypes_AccessFile
            defaultTable = ""
            If LastOpenedType = e_LastOpened_Access Then
                If dbPath = GetSetting(App.Title, DEF_REGISTRY_CONNECTIONS & "\Access", "Database Path", "", HKEY_LOCAL_MACHINE, "SOFTWARE\Database") Then
                    defaultTable = GetSetting(App.Title, DEF_REGISTRY_CONNECTIONS & "\Access", "Default Table", "", HKEY_LOCAL_MACHINE, "SOFTWARE\Database")
                End If
            End If
            LastOpenedType = e_LastOpened_Access
            Call SaveSetting(App.Title, DEF_REGISTRY_CONNECTIONS & "\Access", "Database Path", dbPath, HKEY_LOCAL_MACHINE, "SOFTWARE\Database")
            Call SetupDatabase(defaultTable, True)
        End If
    Else
        If LastOpenedType = e_LastOpened_Access Then
            dbType = e_databaseTypes_AccessFile
            dbPath = GetSetting(App.Title, DEF_REGISTRY_CONNECTIONS & "\Access", "Database Path", "", HKEY_LOCAL_MACHINE, "SOFTWARE\Database")
            If dbPath <> "" Then
                If Dir(dbPath) <> "" Then
                    dbConnectionString = BuildConnectString(e_databaseTypes_MicrosoftAccess2KFile, dbPath)
                    defaultTable = GetSetting(App.Title, DEF_REGISTRY_CONNECTIONS & "\Access", "Default Table", "", HKEY_LOCAL_MACHINE, "SOFTWARE\Database")
                    Call SetupDatabase(defaultTable, True)
                End If
            End If
        ElseIf LastOpenedType = e_LastOpened_Network Then
            serverType = GetSetting(App.Title, DEF_REGISTRY_CONNECTIONS, "", 0, HKEY_LOCAL_MACHINE, "SOFTWARE\Database")
            
            If serverType = e_databaseTypes_OracleMSDA Then
                registryString = DEF_REGISTRY_CONNECTIONS & "\" & DEF_ORACLE_CLIENT
            ElseIf serverType = e_databaseTypes_OracleODBC Then
                registryString = DEF_REGISTRY_CONNECTIONS & "\" & DEF_ORACLE_ODBC
            ElseIf serverType = e_databaseTypes_SQLserver Then
                registryString = DEF_REGISTRY_CONNECTIONS & "\" & DEF_SQL_SERVER
            ElseIf serverType = e_databaseTypes_DSNFile Then
                registryString = DEF_REGISTRY_CONNECTIONS & "\" & DEF_DSN_FILE
            End If
            
            serverName = GetSetting(App.Title, registryString, "Server Name", "", HKEY_LOCAL_MACHINE, "SOFTWARE\Database")
            databaseName = GetSetting(App.Title, registryString, "Database Name", "", HKEY_LOCAL_MACHINE, "SOFTWARE\Database")
            UserName = GetSetting(App.Title, registryString, "Username", "", HKEY_LOCAL_MACHINE, "SOFTWARE\Database")
            Password = GetSetting(App.Title, registryString, "Password", "", HKEY_LOCAL_MACHINE, "SOFTWARE\Database")
            defaultTable = GetSetting(App.Title, registryString, "Default Table", "", HKEY_LOCAL_MACHINE, "SOFTWARE\Database")
             
            
            dbType = serverType
            dbPath = UCase(Mid(serverName, 1, 1)) & LCase(Mid(serverName, 2))
            If Trim(databaseName) <> "" Then
                dbPath = dbPath & " : " & databaseName
            End If
            
            dbConnectionString = BuildConnectString(serverType, serverName, databaseName, UserName, Password)
            Call SetupDatabase(defaultTable, True)
        End If
    End If
    
End Sub
Public Sub SetupDatabase(Optional defaultTable As String, Optional centerScreen As Boolean = False)
    If dbType = e_databaseTypes_AccessFile Or dbType = e_databaseTypes_MicrosoftAccess2KFile Or dbType = e_databaseTypes_MicrosoftAccess97File Then
        frmMain.mnuCompact.Visible = True
    Else
        frmMain.mnuCompact.Visible = False
    End If
    
    lblDatabaseName.caption = BuildDatabaseName(dbType, dbPath, LastOpenedType)
    Call frmConnecting.ShowConnecting("Connecting to " & lblDatabaseName.caption, IIf(centerScreen = False, Me, Nothing))
    Me.caption = App.Title & " (" & lblDatabaseName.caption & ")"
    Call GetTableList(defaultTable, centerScreen)
    frmConnecting.Hide
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim frmObj As Object
    If Me.WindowState <> vbMinimized Then
        Call SaveSetting(App.Title, DEF_REGISTRY_SETTINGS, "WindowState", Me.WindowState, HKEY_LOCAL_MACHINE, "SOFTWARE\Database")
    End If
    If Me.WindowState = vbNormal Then
        Call SaveSetting(App.Title, DEF_REGISTRY_SETTINGS, "Form Top", Me.Top, HKEY_LOCAL_MACHINE, "SOFTWARE\Database")
        Call SaveSetting(App.Title, DEF_REGISTRY_SETTINGS, "Form Left", Me.Left, HKEY_LOCAL_MACHINE, "SOFTWARE\Database")
        Call SaveSetting(App.Title, DEF_REGISTRY_SETTINGS, "Form Width", Me.Width, HKEY_LOCAL_MACHINE, "SOFTWARE\Database")
        Call SaveSetting(App.Title, DEF_REGISTRY_SETTINGS, "Form Height", Me.Height, HKEY_LOCAL_MACHINE, "SOFTWARE\Database")
    End If
    Call SaveSetting(App.Title, DEF_REGISTRY_SETTINGS, "Last Opened Type", CStr(LastOpenedType), HKEY_LOCAL_MACHINE, "SOFTWARE\Database")
    Call SaveSetting(App.Title, DEF_REGISTRY_SETTINGS, "Show SQL", CStr(mnuShowSQL.Checked), HKEY_LOCAL_MACHINE, "SOFTWARE\Database")
    
    Call SaveDefaultTable
    
    For Each frmObj In Forms
        Unload frmObj
    Next
    
    Set dbObj = Nothing
    Call Unhook
End Sub

Private Sub Form_Resize()

    If mnuShowSQL.Checked = True Then
        fraSqlStatement.Visible = True
    Else
        fraSqlStatement.Visible = False
    End If
    
    If Me.WindowState = vbMinimized Then Exit Sub
    
    'Height
    If mnuShowSQL.Checked = True Then
        gridData.Height = Me.Height - gridData.Top - 2000
    Else
        gridData.Height = Me.Height - gridData.Top - 1050
    End If
    
    'Width
    gridData.Width = Me.Width - 360
    fraSqlStatement.Width = gridData.Width
    txtSqlStatement.Width = fraSqlStatement.Width - txtSqlStatement.Left - cmdExecute.Width - 200
    
    'Top
    imgLoading.Top = Me.Height - 950
    lblStatus.Top = Me.Height - 950
    fraSqlStatement.Top = gridData.Top + gridData.Height + 100
    
    'Left
    imgLoading.Left = gridData.Left + gridData.Width - imgLoading.Width
    chkEditMode.Left = gridData.Left + gridData.Width - chkEditMode.Width
    cmdExecute.Left = txtSqlStatement.Left + txtSqlStatement.Width + 100
    
End Sub

Private Sub GetTableList(Optional ByVal defaultTable As String, Optional ByVal centerScreen As Boolean = False)
Dim rsSchema As ADODB.Recordset
Dim nCount As Integer
Dim newTableName As String
Dim Password As String

    On Error Resume Next
    LockTextBox = True
    frmMain.cmbTables.Clear
    frmPurgeDate.cmbTables.Clear
    frmRenameTable.cmbTables.Clear
    If dbObj.State = adStateOpen Then
        Set dbObj = New ADODB.Connection
    End If
    dbObj.Open dbConnectionString
    If dbObj.State <> adStateOpen And (dbType = e_databaseTypes_AccessFile Or e_databaseTypes_MicrosoftAccess2KFile Or e_databaseTypes_MicrosoftAccess97File) Then
        If centerScreen = True Then
            Password = GetSetting(App.Title, DEF_REGISTRY_CONNECTIONS & "\" & DEF_ACCESS, "Database Password", "", HKEY_LOCAL_MACHINE, "SOFTWARE\Database")
        End If
        If Password = "" Then
            frmPassword.caption = DetermineFilename(dbPath) & " Password"
            frmPassword.Show vbModal, IIf(centerScreen = False, Me, Nothing)
            If frmPassword.bWasCancelled = True Then
                Exit Sub
            End If
        End If
        dbConnectionString = dbConnectionString & "PWD=" & frmPassword.txtPassword & ";"
        dbObj.Open dbConnectionString
        Unload frmPassword
        mnuClearPassword.Visible = True
    Else
        mnuClearPassword.Visible = False
        Call SaveSetting(App.Title, DEF_REGISTRY_CONNECTIONS & "\" & DEF_ACCESS, "Database Password", "", HKEY_LOCAL_MACHINE, "SOFTWARE\Database")
    End If
    If dbObj.State = adStateOpen Then
        Set rsSchema = dbObj.OpenSchema(adSchemaTables, Array(Empty, Empty, Empty, "TABLE"))
        If Not rsSchema Is Nothing Then
            Do While Not rsSchema.EOF
                If UCase(Left(rsSchema!Table_name, 4)) <> "MSYS" Then
                    If UCase(Left(rsSchema!Table_name, 11)) <> "SWITCHBOARD" Then
                        newTableName = rsSchema!Table_name
                        cmbTables.AddItem newTableName
                        frmPurgeDate.cmbTables.AddItem newTableName
                        frmRenameTable.cmbTables.AddItem newTableName
                    End If
                End If
                rsSchema.MoveNext
            Loop
            cmbTables.AddItem DEF_CUSTOM_SQL
        End If
    End If
    rsSchema.Close
    If cmbTables.ListCount > 0 Then
        If defaultTable = "" Then
            cmbTables.ListIndex = 0
            frmPurgeDate.cmbTables.ListIndex = 0
            frmRenameTable.cmbTables.ListIndex = 0
        Else
            For nCount = 0 To cmbTables.ListCount - 1
                If cmbTables.List(nCount) = defaultTable Then
                    cmbTables.ListIndex = nCount
                    frmPurgeDate.cmbTables.ListIndex = nCount
                    frmRenameTable.cmbTables.ListIndex = nCount
                    Exit For
                End If
            Next nCount
            If cmbTables.ListIndex = -1 Then
                cmbTables.ListIndex = 0
                frmPurgeDate.cmbTables.ListIndex = 0
                frmRenameTable.cmbTables.ListIndex = 0
            End If
        End If
    End If
    
    Set gridData.DataSource = adoData

    Set rsSchema = Nothing
    LockTextBox = False
End Sub

Private Sub chkEditMode_Click()
    mnuEditMode.Checked = IIf(chkEditMode.Value = vbChecked, True, False)
    If chkEditMode.Value = vbChecked Then
        gridData.AllowAddNew = True
        gridData.AllowDelete = True
        gridData.AllowUpdate = True
    Else
        gridData.AllowAddNew = False
        gridData.AllowDelete = False
        gridData.AllowUpdate = False
    End If
    If chkEditMode.Value = vbChecked And cmbTables.Text <> DEF_CUSTOM_SQL Then
        mnuEdit.Enabled = True
    Else
        mnuEdit.Enabled = False
    End If
    
End Sub

Private Sub cmbTables_Change()
    Call cmbTables_Click
End Sub

Private Sub cmbTables_Click()
    Call LoadData
End Sub
Public Sub LoadData()
    
    On Error GoTo e_Trap
    Call chkEditMode_Click
    If cmbTables.Text = DEF_CUSTOM_SQL Then
        mnuShowSQL.Checked = True
        Call Form_Resize
        On Error Resume Next
        If LockTextBox = False Then
            txtSqlStatement.SetFocus
            txtSqlStatement.SelStart = 0
            txtSqlStatement.SelLength = Len(txtSqlStatement)
        End If
        Exit Sub
    End If
    
    LockTextBox = True
    If cmbTables.Text = "" Then Exit Sub
    imgLoading.Visible = True
    Me.Refresh
    SortOrder = True
    LastColumnSort = 0
    Set gridData.DataSource = Nothing
    adoData.RecordSource = ""
    adoData.ConnectionString = ""
    adoData.ConnectionString = dbConnectionString
    adoData.RecordSource = "SELECT * FROM " & ResolveTable(cmbTables.Text)
    adoData.Refresh
    txtSqlStatement = adoData.RecordSource
    If adoData.Recordset.Fields.Count = 0 Then
        gridData.ClearFields
    Else
        Set gridData.DataSource = adoData.Recordset
        gridData.ClearFields
        gridData.ReBind
    End If
    lblStatus.caption = "Record Count: " & adoData.Recordset.RecordCount
    imgLoading.Visible = False
    LockTextBox = False
    Exit Sub
e_Trap:
    lblStatus.caption = "Error: " & Err.Description & " (" & Err.Number & ")"
    imgLoading.Visible = False
    LockTextBox = False

End Sub

Private Sub gridData_HeadClick(ByVal ColIndex As Integer)
Dim startingSql As String
Dim lastSql As String

    On Error GoTo e_Trap
    LockTextBox = True
    imgLoading.Visible = True
    Me.Refresh
    Call LockWindow(gridData.hwnd)
    If LastColumnSort = ColIndex Then
        SortOrder = Not SortOrder
    Else
        SortOrder = True
    End If
    lastSql = adoData.RecordSource
    If cmbTables.Text = DEF_CUSTOM_SQL Then
        If InStr(1, UCase(txtSqlStatement), "ORDER BY") <> 0 Then
            startingSql = Mid(txtSqlStatement, 1, InStr(1, UCase(txtSqlStatement), "ORDER BY") - 2)
            adoData.RecordSource = startingSql & " ORDER BY " & ResolveTable(adoData.Recordset.Fields(ColIndex).Name) & " " & IIf(SortOrder, "ASC", "DESC")
        Else
            adoData.RecordSource = txtSqlStatement & " ORDER BY " & ResolveTable(adoData.Recordset.Fields(ColIndex).Name) & " " & IIf(SortOrder, "ASC", "DESC")
        End If
    Else
        adoData.RecordSource = "SELECT * FROM " & ResolveTable(cmbTables.Text) & " ORDER BY " & ResolveTable(adoData.Recordset.Fields(ColIndex).Name) & " " & IIf(SortOrder, "ASC", "DESC")
    End If
    LastColumnSort = ColIndex
    txtSqlStatement = adoData.RecordSource
    adoData.Refresh
    Set gridData.DataSource = adoData
    cmbTables.SetFocus
    Call ReleaseWindow
    imgLoading.Visible = False
    LockTextBox = False
    Exit Sub
e_Trap:
    lblStatus.caption = "Order Error: " & Err.Description & " (" & Err.Number & ")"
    If adoData.Recordset Is Nothing And lastSql <> "" Then
        adoData.RecordSource = lastSql
        adoData.Refresh
    End If
    Call ReleaseWindow
    LockTextBox = False
End Sub



Private Sub mnuAddColumn_Click()
    frmAddColumn.sTableName = ResolveTable(cmbTables.Text)
    frmAddColumn.Show vbModal, Me
End Sub

Private Sub mnuAddTable_Click()
    frmAddTable.Show vbModal, Me
End Sub

Private Sub mnuClearPassword_Click()
    mnuClearPassword.Visible = False
    Call SaveSetting(App.Title, DEF_REGISTRY_CONNECTIONS & "\" & DEF_ACCESS, "Database Password", "", HKEY_LOCAL_MACHINE, "SOFTWARE\Database")
End Sub

Private Sub mnuCompact_Click()
Dim JetEngine As New JRO.JetEngine
Dim tempFilename As String
Dim startingType As String
Dim endingType As String
Dim existingTable As String
Dim errorMessage As String
    
    On Error GoTo e_Trap
    
    If Dir(dbPath) = "" Or Trim(dbPath) = "" Then Exit Sub
    
    Call frmConnecting.ShowConnecting("Connecting to " & "Compacting Access Database...", Me)
    If dbObj.State = adStateOpen Then
        On Error Resume Next
        existingTable = frmMain.cmbTables.Text
        Set dbObj = New ADODB.Connection
        Set adoData = New ADODB.Connection
        adoData.Recordset.ActiveConnection = Nothing
        On Error GoTo e_Trap
    End If
    
    tempFilename = Mid(dbPath, 1, Len(dbPath) - 4) & "2.mdb"
    If Dir(tempFilename) <> "" Then
        Kill tempFilename
    End If
    
    startingType = "Provider=Microsoft.Jet.OLEDB." & DEF_ACCESS97_OLEDB & ";" & "Data Source=" & dbPath & ";"
    
    startingType = "Jet OLEDB:Engine Type=4;" & "Data Source=" & dbPath
    endingType = "Provider=Microsoft.Jet.OLEDB." & DEF_ACCESS2K_OLEDB & ";" & "Data Source=" & tempFilename & ";"
    
    JetEngine.CompactDatabase startingType, endingType
            
    If Dir(tempFilename) <> "" Then
        Kill dbPath
        Call FileCopy(tempFilename, dbPath)
        Kill tempFilename
    End If
    
    frmConnecting.Hide
    Call frmMain.SetupDatabase(existingTable)
    Exit Sub
e_Trap:
    errorMessage = "Error: CompactAccessDatabase - " & Err.Description & " (" & Err.Number & ")"
    Call MsgBox(errorMessage, vbOKOnly + vbCritical, "Compact Error")
    frmConnecting.Hide
    Call frmMain.SetupDatabase(existingTable)
End Sub
Private Sub mnuDeleteColumn_Click()
Dim ret As Integer
    On Error GoTo e_Trap
    If gridData.Row <= 0 And gridData.SelStartCol <> -1 And adoData.Recordset.Fields.Count > 1 Then
        ret = MessageBox(Me.hwnd, "Are you sure you want to delete " & adoData.Recordset.Fields(gridData.SelStartCol).Name & " from " & cmbTables.Text & "?", vbYesNo + vbQuestion, "Delete Column")
        If ret = vbYes Then
            Call dbObj.Execute("ALTER TABLE " & ResolveTable(cmbTables.Text) & " DROP COLUMN [" & adoData.Recordset.Fields(gridData.SelStartCol).Name & "]")
            Call SetupDatabase(cmbTables.Text)
        End If
    End If
e_Trap:
    Exit Sub
End Sub

Private Sub mnuDropTable_Click()
Dim ret As Integer
    If cmbTables.Text = DEF_CUSTOM_SQL Then
        Exit Sub
    End If
    ret = MessageBox(Me.hwnd, "Are you sure you want to drop Table: " & cmbTables.Text & "?", vbYesNo + vbQuestion, "Drop Table")
    If ret = vbYes Then
        Call dbObj.Execute("DROP TABLE " & ResolveTable(cmbTables.Text))
        Call GetTableList
    End If
End Sub

Private Sub mnuEditMode_Click()
    chkEditMode.Value = IIf(chkEditMode.Value = vbChecked, vbUnchecked, vbChecked)
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Function SelectFile(Title As String, filter As String, flags As Long, defaultExtension As String, Optional saveFile As Boolean = True, Optional lastFilename As String) As String
Dim sOpen As SelectedFile
Dim filename As String
Dim ret As Integer

    On Error GoTo e_Browse
    FileDialog.sFilter = filter
    FileDialog.flags = flags
    FileDialog.sDlgTitle = Title
    FileDialog.sInitDir = DetermineDirectory(lastFilename)
    FileDialog.sFile = DetermineFilename(lastFilename)
    
    Do While filename = ""
        If saveFile = False Then
            sOpen = ShowOpen(Me.hwnd, True)
        Else
            sOpen = ShowSave(Me.hwnd, True)
        End If
        If sOpen.sFiles(1) = "" Then
            ret = MessageBox(Me.hwnd, "Please select a " & Title, vbOKCancel + vbInformation, "Missing Filename")
            If ret = vbCancel Then
                Exit Function
            End If
        Else
            filename = sOpen.sLastDirectory & sOpen.sFiles(1)
            If InStr(1, filename, ".") = 0 Then
                If LCase(Right(filename, 4)) <> "." & defaultExtension Then
                    filename = filename & "." & defaultExtension
                End If
            End If
            SelectFile = filename
        End If
    Loop
    Exit Function
e_Browse:
    SelectFile = ""
    Exit Function
End Function


Private Sub mnuOpenAccess_Click()
Dim openFile As String
Dim lastFile As String
    On Error GoTo e_Trap
    Call SaveDefaultTable
    
    lastFile = GetSetting(App.Title, DEF_REGISTRY_CONNECTIONS & "\" & DEF_ACCESS, "Database Path", "", HKEY_LOCAL_MACHINE, "SOFTWARE\Database")
    openFile = SelectFile("Select Database", "Microsoft Access Database (*.mdb)" & Chr$(0) & "*.mdb", &H4 + &H1000, "mdb", False, lastFile)
    If openFile <> "" Then
        LastOpenedType = e_LastOpened_Access
        dbPath = openFile
        dbType = e_databaseTypes_AccessFile
        If lastFile <> openFile Then
            Call SaveSetting(App.Title, DEF_REGISTRY_CONNECTIONS & "\" & DEF_ACCESS, "Database Password", "", HKEY_LOCAL_MACHINE, "SOFTWARE\Database")
        End If
        Call SaveSetting(App.Title, DEF_REGISTRY_CONNECTIONS & "\" & DEF_ACCESS, "Database Path", dbPath, HKEY_LOCAL_MACHINE, "SOFTWARE\Database")
        dbConnectionString = "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & dbPath & ";DefaultDir=" & DetermineDirectory(dbPath) & ";"
        Call SetupDatabase
    End If
    Exit Sub
e_Trap:
    Call MessageBox(Me.hwnd, "Error:" & Err.Description & " (" & Err.Number & ")", vbOKOnly + vbCritical, "Error")
End Sub


Private Sub mnuOpenNetwork_Click()
    Call SaveDefaultTable
    frmNetwork.Show vbModal, Me
End Sub

Private Sub mnuedit_Click()
    If cmbTables.Text = DEF_CUSTOM_SQL Then
        mnuDropTable.Enabled = False
    Else
        mnuDropTable.Enabled = True
    End If
    If adoData.Recordset Is Nothing Then
        mnuDeleteColumn.Enabled = False
        mnuRenameColumn.Enabled = False
    Else
        If gridData.Row <= 0 And gridData.SelStartCol <> -1 And adoData.Recordset.Fields.Count > 1 Then
            mnuDeleteColumn.Enabled = True
        Else
            mnuDeleteColumn.Enabled = False
        End If
        If cmbTables.Text <> DEF_CUSTOM_SQL And adoData.Recordset.Fields.Count > 0 Then
            mnuRenameColumn.Enabled = True
        Else
            mnuRenameColumn.Enabled = False
        End If
    End If
End Sub

Private Sub mnuPurgeDate_Click()
    frmPurgeDate.cmbTables.ListIndex = cmbTables.ListIndex
    frmPurgeDate.Show vbModal, Me
End Sub

Private Sub mnuReload_Click()
    Call SetupDatabase(cmbTables.Text)
End Sub

Private Sub mnuRenameColumn_Click()
    Call frmRenameColumn.ReloadColumns(gridData.SelStartCol)
    frmRenameColumn.Show vbModal, Me
End Sub

Private Sub mnuRenameTable_Click()
    frmRenameTable.cmbTables.ListIndex = cmbTables.ListIndex
    frmRenameTable.Show vbModal, Me
End Sub

Private Sub mnuShowSQL_Click()
    mnuShowSQL.Checked = Not mnuShowSQL.Checked
    Call Form_Resize
End Sub

Private Sub mnuViewColumns_Click()
    frmViewColumns.Show vbModal, Me
End Sub

Private Sub txtSqlStatement_Change()
    If Trim(txtSqlStatement) = "" Then
        cmdExecute.Enabled = False
    Else
        cmdExecute.Enabled = True
        cmdExecute.Default = True
        If LockTextBox = False Then
            LockTextBox = True
            cmbTables.ListIndex = cmbTables.ListCount - 1
            LockTextBox = False
        End If
    End If
End Sub

Private Sub txtSqlStatement_LostFocus()
    cmdExecute.Default = False
End Sub
