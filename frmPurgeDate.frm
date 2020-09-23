VERSION 5.00
Begin VB.Form frmPurgeDate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Purge Database Table by Date"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5805
   Icon            =   "frmPurgeDate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   5805
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdQuick 
      Caption         =   "1 Year"
      Height          =   375
      Index           =   3
      Left            =   4920
      TabIndex        =   5
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton cmdQuick 
      Caption         =   "1 Month"
      Height          =   375
      Index           =   2
      Left            =   4200
      TabIndex        =   4
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton cmdQuick 
      Caption         =   "1 Week"
      Height          =   375
      Index           =   1
      Left            =   3480
      TabIndex        =   3
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton cmdQuick 
      Caption         =   "1 Day"
      Height          =   375
      Index           =   0
      Left            =   2760
      TabIndex        =   2
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox txtDate 
      Height          =   285
      Left            =   2760
      TabIndex        =   6
      Top             =   960
      Width           =   2895
   End
   Begin VB.TextBox txtSqlStatement 
      Height          =   495
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   1680
      Width           =   5535
   End
   Begin VB.ComboBox cmbDateColumn 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   960
      Width           =   2415
   End
   Begin VB.ComboBox cmbTables 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
   Begin VB.CommandButton cmdPurge 
      Caption         =   "&Purge"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   4320
      TabIndex        =   9
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      Caption         =   "Oldest Date (mm/dd/yy hh:mm:ss AMPM)"
      Height          =   195
      Left            =   2760
      TabIndex        =   13
      Top             =   720
      Width           =   2925
   End
   Begin VB.Label lblSqlStatement 
      AutoSize        =   -1  'True
      Caption         =   "SQL Statement"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   1440
      Width           =   1080
   End
   Begin VB.Label lblDateColumn 
      AutoSize        =   -1  'True
      Caption         =   "Date Column"
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   720
      Width           =   915
   End
   Begin VB.Label lblTableName 
      AutoSize        =   -1  'True
      Caption         =   "Table Name"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   870
   End
End
Attribute VB_Name = "frmPurgeDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbDateColumn_Change()
    Call BuildSQl
End Sub

Private Sub cmbDateColumn_Click()
    Call BuildSQl
End Sub

Private Sub cmbTables_Change()
    Call cmbTables_Click
End Sub

Private Sub ReloadColumns()
Dim rsNew As New ADODB.Recordset
Dim nCount As Integer
    On Error GoTo e_Trap
    cmbDateColumn.Clear
    Call rsNew.Open(ResolveTable(cmbTables.Text), dbConnectionString, , , adCmdTable)
    If Not rsNew.BOF Then
        For nCount = 0 To rsNew.Fields.Count - 1
            With rsNew.Fields.Item(nCount)
                If .Type = adDate Or .Type = adDBDate Or .Type = adDBFileTime Or .Type = adDBTimeStamp Or .Type = adDBTime Then
                    cmbDateColumn.AddItem .Name
                End If
            End With
        Next nCount
    End If
    If cmbDateColumn.ListCount > 0 Then
        cmbDateColumn.ListIndex = 0
    End If
    Set rsNew = Nothing
    Exit Sub
e_Trap:
    Exit Sub
End Sub

Private Sub cmbTables_Click()
    frmMain.cmbTables.ListIndex = cmbTables.ListIndex
    Call ReloadColumns
    Call BuildSQl
    Call txtSqlStatement_Change
End Sub

Private Sub cmdExit_Click()
    Me.Hide
End Sub

Private Sub BuildSQl()
    If cmbDateColumn.ListCount = 0 Or IsDate(txtDate) = False Then
        txtSqlStatement = ""
    Else
        If dbType = e_databaseTypes_AccessFile Or dbType = e_databaseTypes_MicrosoftAccess2KFile Or dbType = e_databaseTypes_MicrosoftAccess97File Or dbType = e_databaseTypes_DSNFile Then
            txtSqlStatement = "DELETE FROM " & ResolveTable(cmbTables.Text) & " WHERE " & ResolveTable(cmbDateColumn) & " < #" & Format(Trim(txtDate), "m/d/yyyy hh:nn:ss AMPM") & "#"
        Else
            txtSqlStatement = "DELETE FROM " & ResolveTable(cmbTables.Text) & " WHERE " & ResolveTable(cmbDateColumn) & "<'" & Format(Trim(txtDate), "m/d/yyyy hh:nn:ss AMPM") & "'"
        End If
    End If
End Sub

Private Sub cmdPurge_Click()
Dim ret As Integer

    ret = MessageBox(Me.hwnd, "Are you sure you want to purge Table: " & cmbTables.Text & "?", vbYesNo + vbQuestion, "Purge Table")
    If ret = vbYes Then
        On Error GoTo e_Trap
        frmMain.imgLoading.Visible = True
        frmMain.Refresh
        dbObj.Execute txtSqlStatement.Text
        Sleep 1000
        DoEvents
        Call frmMain.LoadData
    End If
    Exit Sub
e_Trap:
    frmMain.imgLoading.Visible = False
    frmMain.lblStatus.caption = "Purge Error: " & Err.Description & " (" & Err.Number & ")"
End Sub

Private Sub cmdQuick_Click(Index As Integer)
    If Index = 0 Then
        txtDate = DateAdd("d", -1, Now)
    ElseIf Index = 1 Then
        txtDate = DateAdd("ww", -1, Date)
    ElseIf Index = 2 Then
        txtDate = DateAdd("m", -1, Date)
    ElseIf Index = 3 Then
        txtDate = DateAdd("yyyy", -1, Date)
    End If
End Sub

Private Sub Form_GotFocus()
    Call BuildSQl
End Sub

Private Sub Form_Load()
    Call cmdQuick_Click(2)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Me.Hide
        Cancel = 1
        Exit Sub
    End If
End Sub

Private Sub mnuExit_Click()
    Me.Hide
End Sub

Private Sub txtDate_Change()
    Call BuildSQl
End Sub

Private Sub txtSqlStatement_Change()
    If Len(Trim(txtSqlStatement)) = 0 Then
        cmdPurge.Enabled = False
    Else
        cmdPurge.Enabled = True
    End If
End Sub
