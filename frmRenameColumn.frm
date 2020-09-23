VERSION 5.00
Begin VB.Form frmRenameColumn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rename Column"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3405
   Icon            =   "frmRenameColumn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   3405
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtNewName 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   3135
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
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.ComboBox cmbColumns 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   3135
   End
   Begin VB.Label lblNewName 
      AutoSize        =   -1  'True
      Caption         =   "Enter New Column Name"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   1785
   End
   Begin VB.Label lblOldName 
      AutoSize        =   -1  'True
      Caption         =   "Existing Column Name"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmRenameColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub ReloadColumns(defaultIndex As Integer)
Dim nCount As Integer

    cmbColumns.Clear
    With frmMain.adoData.Recordset.Fields
        For nCount = 0 To .Count - 1
            cmbColumns.AddItem .Item(nCount).Name
        Next nCount
    End With
    If defaultIndex = -1 Then
        cmbColumns.ListIndex = 0
    ElseIf defaultIndex <= cmbColumns.ListCount - 1 Then
        cmbColumns.ListIndex = defaultIndex
    End If
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdRename_Click()
    
    On Error GoTo e_Trap
    
    If txtNewName = "" Then
        Call MessageBox(Me.hwnd, "Please enter a valid column name", vbOKOnly + vbCritical, "Missing Info")
        txtNewName.SetFocus
        Exit Sub
    End If
    If txtNewName = cmbColumns.Text Then
        Call MessageBox(Me.hwnd, "Please enter a different column name", vbOKOnly + vbCritical, "Missing Info")
        txtNewName.SetFocus
        Exit Sub
    End If
    Me.Hide
    Call frmMain.SetupDatabase(frmMain.cmbTables.Text)
    Exit Sub
e_Trap:
    Call MessageBox(Me.hwnd, "Rename Column Error: " & Err.Description & " (" & Err.Number & ")", vbOKOnly + vbCritical, "Rename Column Error")
End Sub
