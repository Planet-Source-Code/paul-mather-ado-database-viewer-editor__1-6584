VERSION 5.00
Begin VB.Form frmNetwork 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Network Database"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4875
   Icon            =   "frmNetwork.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   4875
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2520
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1680
      Width           =   2175
   End
   Begin VB.TextBox txtUsername 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   2175
   End
   Begin VB.TextBox txtDatabaseName 
      Height          =   285
      Left            =   2520
      TabIndex        =   2
      Top             =   1080
      Width           =   2175
   End
   Begin VB.TextBox txtServerName 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   2175
   End
   Begin VB.ComboBox cmdDatabaseType 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   3480
      TabIndex        =   6
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "&Connect"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label lblPassword 
      AutoSize        =   -1  'True
      Caption         =   "Password"
      Height          =   195
      Left            =   2520
      TabIndex        =   11
      Top             =   1440
      Width           =   690
   End
   Begin VB.Label lblUsername 
      AutoSize        =   -1  'True
      Caption         =   "Username"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   720
   End
   Begin VB.Label lblDatabaseName 
      AutoSize        =   -1  'True
      Caption         =   "Database Name"
      Height          =   195
      Left            =   2520
      TabIndex        =   9
      Top             =   840
      Width           =   1155
   End
   Begin VB.Label lblServerName 
      AutoSize        =   -1  'True
      Caption         =   "Server Name"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   930
   End
   Begin VB.Label lblDatabaseType 
      AutoSize        =   -1  'True
      Caption         =   "Database Type"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmNetwork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim defaultTable As String

Private Sub cmdConnect_Click()
Dim servernameString As String

    dbConnectionString = BuildConnectString(cmdDatabaseType.ItemData(cmdDatabaseType.ListIndex), txtServerName, txtDatabaseName, txtUsername, txtPassword)
    
    Me.Hide
    servernameString = UCase(Mid(txtServerName, 1, 1)) & LCase(Mid(txtServerName, 2))
    If txtDatabaseName <> "" Then
        servernameString = servernameString & " : " & UCase(Mid(txtDatabaseName, 1, 1)) & LCase(Mid(txtDatabaseName, 2))
    End If
    Call frmConnecting.ShowConnecting("Connecting to " & BuildDatabaseName(cmdDatabaseType.ItemData(cmdDatabaseType.ListIndex), servernameString, e_LastOpened_Network), frmMain, cmdDatabaseType.ItemData(cmdDatabaseType.ListIndex))
    On Error GoTo e_Trap
    If dbObj.State = adStateOpen Then
        dbObj.Close
    End If
    dbObj.Open dbConnectionString
    If dbObj.State <> adStateOpen Then
        Call MessageBox(Me.hwnd, "Failed to open Database on " & txtServerName, vbOKOnly + vbCritical, "Connect Failure")
        frmConnecting.Hide
        Me.Show vbModal, frmMain
    Else
        dbType = cmdDatabaseType.ItemData(cmdDatabaseType.ListIndex)
        dbPath = UCase(Mid(txtServerName, 1, 1)) & LCase(Mid(txtServerName, 2))
        If txtDatabaseName <> "" Then
            dbPath = dbPath & " : " & UCase(Mid(txtDatabaseName, 1, 1)) & LCase(Mid(txtDatabaseName, 2))
        End If
        LastOpenedType = e_LastOpened_Network
        Call frmMain.SetupDatabase(defaultTable)
        Call SaveSetting(App.Title, DEF_REGISTRY_CONNECTIONS, "", cmdDatabaseType.ItemData(cmdDatabaseType.ListIndex), HKEY_LOCAL_MACHINE, "SOFTWARE\Database")
        Call SaveSetting(App.Title, DEF_REGISTRY_CONNECTIONS & "\" & cmdDatabaseType.Text, "Server Name", txtServerName, HKEY_LOCAL_MACHINE, "SOFTWARE\Database")
        Call SaveSetting(App.Title, DEF_REGISTRY_CONNECTIONS & "\" & cmdDatabaseType.Text, "Database Name", txtDatabaseName, HKEY_LOCAL_MACHINE, "SOFTWARE\Database")
        Call SaveSetting(App.Title, DEF_REGISTRY_CONNECTIONS & "\" & cmdDatabaseType.Text, "Username", txtUsername, HKEY_LOCAL_MACHINE, "SOFTWARE\Database")
        Call SaveSetting(App.Title, DEF_REGISTRY_CONNECTIONS & "\" & cmdDatabaseType.Text, "Password", txtPassword, HKEY_LOCAL_MACHINE, "SOFTWARE\Database")
        Unload Me
    End If
    Exit Sub
e_Trap:
    Call MessageBox(Me.hwnd, "Error Connecting to Database on " & txtServerName & vbCr & Err.Description & " (" & Err.Number & ")", vbCritical + vbOKOnly, "Connect Failure")
    frmConnecting.Hide
    Me.Show vbModal, frmMain
End Sub

Private Sub cmdDatabaseType_Change()
    Call cmdDatabaseType_Click
End Sub

Private Sub cmdDatabaseType_Click()
    If cmdDatabaseType.ItemData(cmdDatabaseType.ListIndex) = e_databaseTypes_DSNFile Then
        lblServerName.caption = "Data Source Name"
    Else
        lblServerName.caption = "Server Name"
    End If
    txtServerName = GetSetting(App.Title, DEF_REGISTRY_CONNECTIONS & "\" & cmdDatabaseType.Text, "Server Name", "", HKEY_LOCAL_MACHINE, "SOFTWARE\Database")
    txtDatabaseName = GetSetting(App.Title, DEF_REGISTRY_CONNECTIONS & "\" & cmdDatabaseType.Text, "Database Name", "", HKEY_LOCAL_MACHINE, "SOFTWARE\Database")
    txtUsername = GetSetting(App.Title, DEF_REGISTRY_CONNECTIONS & "\" & cmdDatabaseType.Text, "Username", "", HKEY_LOCAL_MACHINE, "SOFTWARE\Database")
    txtPassword = GetSetting(App.Title, DEF_REGISTRY_CONNECTIONS & "\" & cmdDatabaseType.Text, "Password", "", HKEY_LOCAL_MACHINE, "SOFTWARE\Database")
    defaultTable = GetSetting(App.Title, DEF_REGISTRY_CONNECTIONS & "\" & cmdDatabaseType.Text, "Default Table", "", HKEY_LOCAL_MACHINE, "SOFTWARE\Database")

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim nCount As Integer
Dim ListIndex As Integer

    cmdDatabaseType.AddItem DEF_ORACLE_CLIENT
'    cmdDatabaseType.AddItem DEF_ORACLE_odbc
    cmdDatabaseType.AddItem DEF_SQL_SERVER
    cmdDatabaseType.AddItem DEF_DSN_FILE
    
    cmdDatabaseType.ItemData(0) = e_databaseTypes_OracleMSDA
'    cmdDatabaseType.ItemData(1) = e_databaseTypes_OracleODBC
    cmdDatabaseType.ItemData(1) = e_databaseTypes_SQLserver
    cmdDatabaseType.ItemData(2) = e_databaseTypes_DSNFile
    
    On Error Resume Next
    ListIndex = GetSetting(App.Title, DEF_REGISTRY_CONNECTIONS, "", 0, HKEY_LOCAL_MACHINE, "SOFTWARE\Database")
    For nCount = 0 To cmdDatabaseType.ListCount - 1
        If cmdDatabaseType.ItemData(nCount) = ListIndex Then
            cmdDatabaseType.ListIndex = nCount
            Exit For
        End If
    Next nCount
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmMain.ZOrder
End Sub

Private Sub txtServerName_GotFocus()
    Call SelectText(txtServerName)
End Sub
Private Sub txtDatabaseName_GotFocus()
    Call SelectText(txtDatabaseName)
End Sub
Private Sub txtUsername_GotFocus()
    Call SelectText(txtUsername)
End Sub
Private Sub txtPassword_GotFocus()
    Call SelectText(txtPassword)
End Sub
Private Sub SelectText(ByRef textObj As TextBox)
    textObj.SelStart = 0
    textObj.SelLength = Len(textObj)
End Sub
