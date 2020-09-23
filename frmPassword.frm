VERSION 5.00
Begin VB.Form frmPassword 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database Password"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3195
   Icon            =   "frmPassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   3195
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkRemember 
      Caption         =   "Save Password"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2415
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label lblPassword 
      AutoSize        =   -1  'True
      Caption         =   "Database Password:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1470
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bWasCancelled As Boolean
Private Sub cmdCancel_Click()
    bWasCancelled = True
    frmPassword.Hide
End Sub
Private Sub cmdOk_Click()
    Call SaveSetting(App.Title, DEF_REGISTRY_CONNECTIONS & "\" & DEF_ACCESS, "Save Option", chkRemember.Value, HKEY_LOCAL_MACHINE, "SOFTWARE\Database")
    If chkRemember.Value = vbChecked Then
        Call SaveSetting(App.Title, DEF_REGISTRY_CONNECTIONS & "\" & DEF_ACCESS, "Database Password", txtPassword, HKEY_LOCAL_MACHINE, "SOFTWARE\Database")
    Else
        Call SaveSetting(App.Title, DEF_REGISTRY_CONNECTIONS & "\" & DEF_ACCESS, "Database Password", "", HKEY_LOCAL_MACHINE, "SOFTWARE\Database")
    End If
    bWasCancelled = False
    frmPassword.Hide
End Sub

Private Sub Form_GotFocus()
    txtPassword.SetFocus
End Sub

Private Sub txtPassword_GotFocus()
    If Len(txtPassword) > 0 Then
        txtPassword.SelStart = 0
        txtPassword.SelLength = Len(txtPassword)
    End If
End Sub

Private Sub Form_Load()
    chkRemember.Value = GetSetting(App.Title, DEF_REGISTRY_CONNECTIONS & "\" & DEF_ACCESS, "Save Option", vbChecked, HKEY_LOCAL_MACHINE, "SOFTWARE\Database")
    txtPassword = GetSetting(App.Title, DEF_REGISTRY_CONNECTIONS & "\" & DEF_ACCESS, "Database Password", "", HKEY_LOCAL_MACHINE, "SOFTWARE\Database")
End Sub
