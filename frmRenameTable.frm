VERSION 5.00
Begin VB.Form frmRenameTable 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rename Table"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3360
   Icon            =   "frmRenameTable.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   3360
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbTables 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   3135
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdRename 
      Caption         =   "&Rename"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox txtNewName 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   3135
   End
   Begin VB.Label lblOldName 
      AutoSize        =   -1  'True
      Caption         =   "Existing Table Name"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblNewName 
      AutoSize        =   -1  'True
      Caption         =   "Enter New Table Name"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1665
   End
End
Attribute VB_Name = "frmRenameTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbTables_Change()
    Call cmbTables_Click
End Sub

Private Sub cmbTables_Click()
    frmMain.cmbTables.ListIndex = cmbTables.ListIndex
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdRename_Click()
Dim sqlStatement As String

    On Error GoTo e_Trap
    
    txtNewName = Trim(txtNewName)
    If txtNewName = "" Then
        Call MessageBox(Me.hwnd, "Please enter a valid table name", vbOKOnly + vbCritical, "Missing Info")
        txtNewName.SetFocus
        Exit Sub
    End If
    If txtNewName = cmbTables.Text Then
        Call MessageBox(Me.hwnd, "Please enter a different table name", vbOKOnly + vbCritical, "Missing Info")
        txtNewName.SetFocus
        Exit Sub
    End If

    sqlStatement = "SELECT " & ResolveTable(cmbTables.Text) & ".* INTO " & ResolveTable(txtNewName) & " FROM " & ResolveTable(cmbTables.Text)
    Call dbObj.Execute(sqlStatement)
    sqlStatement = "DROP TABLE " & ResolveTable(cmbTables.Text)
    Call dbObj.Execute(sqlStatement)
    
    Me.Hide
    Call frmMain.SetupDatabase(txtNewName)

    Exit Sub
e_Trap:
    Call MessageBox(Me.hwnd, "Rename Table Error: " & Err.Description & " (" & Err.Number & ")", vbOKOnly + vbCritical, "Rename Table Error")
End Sub

Private Sub Form_GotFocus()
    txtNewName.SetFocus
End Sub

