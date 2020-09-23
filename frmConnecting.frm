VERSION 5.00
Begin VB.Form frmConnecting 
   BackColor       =   &H000000FF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   720
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   2475
   ControlBox      =   0   'False
   ForeColor       =   &H00808080&
   Icon            =   "frmConnecting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   720
   ScaleWidth      =   2475
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label lblConnecting 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Connecting"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   1170
   End
   Begin VB.Image imgConnecting 
      Height          =   480
      Left            =   240
      Picture         =   "frmConnecting.frx":000C
      Top             =   120
      Width           =   480
   End
   Begin VB.Shape shpBack 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   2415
   End
End
Attribute VB_Name = "frmConnecting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const DEF_BORDER_SIZE As Integer = 30
Public Sub ShowConnecting(ByVal caption As String, Optional ByRef parentObj As Object, Optional ByVal databaseType As e_DatabaseTypes = e_DatabaseTypes_Undefined)
    If databaseType = e_DatabaseTypes_Undefined Then
        databaseType = dbType
    End If
    
    If databaseType = e_databaseTypes_OracleMSDA Or databaseType = e_databaseTypes_OracleODBC Then
        Me.BackColor = vbRed
    ElseIf databaseType = e_databaseTypes_SQLserver Then
        Me.BackColor = vbGreen
    ElseIf databaseType = e_databaseTypes_DSNFile Then
        Me.BackColor = vbYellow
    Else
        Me.BackColor = vbBlue
    End If
    lblConnecting.caption = caption
    Me.Width = lblConnecting.Left + lblConnecting.Width + 300
    shpBack.Width = Me.Width - DEF_BORDER_SIZE * 2 - IIf(Me.BorderStyle = 0, -5, 25)
    shpBack.Height = Me.Height - DEF_BORDER_SIZE * 2 - IIf(Me.BorderStyle = 0, -5, 25)
    
    If parentObj Is Nothing Then
        Me.Top = WorkAreaHeight / 2 - Me.Height / 2
        Me.Left = WorkAreaWidth / 2 - Me.Width / 2
    End If
    Me.Show , parentObj
    DoEvents
End Sub

Private Sub Form_Load()
    shpBack.Top = DEF_BORDER_SIZE
    shpBack.Left = DEF_BORDER_SIZE
End Sub
