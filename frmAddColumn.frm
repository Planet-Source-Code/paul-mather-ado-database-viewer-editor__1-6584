VERSION 5.00
Begin VB.Form frmAddColumn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add New Column"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3435
   Icon            =   "frmAddColumn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   3435
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtNewName 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3135
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.ComboBox cmbType 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   960
      Width           =   3135
   End
   Begin VB.Label lblNewName 
      AutoSize        =   -1  'True
      Caption         =   "Enter New Column Name"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1785
   End
   Begin VB.Label lblNewType 
      AutoSize        =   -1  'True
      Caption         =   "New Column Type"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1305
   End
End
Attribute VB_Name = "frmAddColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bWasCancelled As Boolean
Public sTableName As String
Private Sub cmdCancel_Click()
    bWasCancelled = True
    Me.Hide
End Sub

Private Sub cmdOk_Click()
Dim sqlStatement As String
Dim typeString As String

    txtNewName = Replace(txtNewName, "'", "")
    txtNewName = Replace(txtNewName, """", "")
    txtNewName = Replace(txtNewName, "[", "")
    txtNewName = Replace(txtNewName, "]", "")
        
    If Trim(txtNewName) = "" Then
        Call MessageBox(Me.hwnd, "Please Enter a Column Name", vbOKOnly + vbCritical, "Missing Data")
        txtNewName.SetFocus
        Exit Sub
    End If
    
    Select Case cmbType.ItemData(cmbType.ListIndex)
        Case adVarChar
            typeString = "TEXT"
        Case adVarNumeric
            typeString = "FLOAT"
        Case adInteger
            typeString = "INTEGER"
        Case adDate
            typeString = "DATETIME"
        Case adBoolean
            typeString = "BIT"
    End Select
    
    sqlStatement = "ALTER TABLE " & sTableName & " ADD [" & txtNewName & "] " & typeString
    On Error GoTo e_Trap
    Call dbObj.Execute(sqlStatement)
    
    Me.Hide
    Call frmMain.SetupDatabase(sTableName)
    bWasCancelled = False
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
    cmbType.AddItem "Boolean"
    
    cmbType.ItemData(0) = adVarChar
    cmbType.ItemData(1) = adVarNumeric
    cmbType.ItemData(2) = adInteger
    cmbType.ItemData(3) = adDate
    cmbType.ItemData(4) = adBoolean
    
    cmbType.ListIndex = 0
End Sub
