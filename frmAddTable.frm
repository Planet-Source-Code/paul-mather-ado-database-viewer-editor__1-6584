VERSION 5.00
Begin VB.Form frmAddTable 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add New Table"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3360
   Icon            =   "frmAddTable.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   3360
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbType 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1560
      Width           =   3135
   End
   Begin VB.TextBox txtFirstColumn 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   3135
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox txtNewName 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3135
   End
   Begin VB.Label lblNewType 
      AutoSize        =   -1  'True
      Caption         =   "New Column Type"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   1305
   End
   Begin VB.Label lblFirstColumn 
      AutoSize        =   -1  'True
      Caption         =   "Enter First Column Name"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   1740
   End
   Begin VB.Label lblNewName 
      AutoSize        =   -1  'True
      Caption         =   "Enter New Table Name"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1665
   End
End
Attribute VB_Name = "frmAddTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdOk_Click()
Dim sqlStatement As String
Dim typeString As String
    
    txtNewName = Replace(txtNewName, "'", "")
    txtNewName = Replace(txtNewName, """", "")
    txtNewName = Replace(txtNewName, "[", "")
    txtNewName = Replace(txtNewName, "]", "")
        
    txtFirstColumn = Replace(txtFirstColumn, "'", "")
    txtFirstColumn = Replace(txtFirstColumn, """", "")
    txtFirstColumn = Replace(txtFirstColumn, "[", "")
    txtFirstColumn = Replace(txtFirstColumn, "]", "")
        
    If Trim(txtNewName) = "" Then
        Call MessageBox(Me.hwnd, "Please Enter a Table Name", vbOKOnly + vbCritical, "Missing Data")
        txtNewName.SetFocus
        Exit Sub
    End If
    
    If Trim(txtFirstColumn) = "" Then
        Call MessageBox(Me.hwnd, "Please Enter a Column Name", vbOKOnly + vbCritical, "Missing Data")
        txtFirstColumn.SetFocus
        Exit Sub
    End If
    
    On Error GoTo e_Trap
    If dbObj.State <> adStateOpen Then
        dbObj.Open dbConnectionString
    End If
    Call dbObj.Execute("CREATE TABLE " & ResolveTable(txtNewName))
    
    Select Case cmbType.ItemData(cmbType.ListIndex)
        Case adVarChar
            typeString = "TEXT"
        Case adVarNumeric
            typeString = "FLOAT"
        Case adInteger
            typeString = "INTEGER"
        Case adDate
            typeString = "DATETIME"
        Case adBinary
            typeString = "BIT"
    End Select
    
    sqlStatement = "ALTER TABLE " & ResolveTable(txtNewName) & " ADD [" & txtFirstColumn & "] " & typeString
    On Error GoTo e_Trap
    Call dbObj.Execute(sqlStatement)
    
    Me.Hide
    Call frmMain.SetupDatabase(txtNewName)
    
    Me.Hide
    Exit Sub
e_Trap:
    Call MessageBox(Me.hwnd, "Error: " & Err.Description & " (" & Err.Number & ")", vbOKOnly + vbCritical, "Column Add Error")
    Exit Sub
End Sub

Private Sub Form_Load()
    cmbType.Clear
    cmbType.AddItem "Text"
    cmbType.AddItem "Float"
    cmbType.AddItem "Integer"
    cmbType.AddItem "Date/Time"
    cmbType.AddItem "Bit"
    
    cmbType.ItemData(0) = adVarChar
    cmbType.ItemData(1) = adVarNumeric
    cmbType.ItemData(2) = adInteger
    cmbType.ItemData(3) = adDate
    cmbType.ItemData(4) = adBinary
    
    cmbType.ListIndex = 0
End Sub
