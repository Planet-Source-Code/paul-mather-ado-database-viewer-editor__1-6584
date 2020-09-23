VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmViewColumns 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Column Types"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7260
   Icon            =   "frmViewColumns.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   7260
   StartUpPosition =   1  'CenterOwner
   Begin MSFlexGridLib.MSFlexGrid gridColumnTypes 
      Height          =   2415
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   4260
      _Version        =   393216
      FixedCols       =   0
      BackColorBkg    =   12632256
      FocusRect       =   0
      HighLight       =   0
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   6000
      TabIndex        =   0
      Top             =   2640
      Width           =   1095
   End
End
Attribute VB_Name = "frmViewColumns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim nCount As Integer
Dim typeString As String

    With gridColumnTypes
        .Clear
        .Cols = 3
        
        .ColWidth(0) = 4000
        .ColWidth(1) = 2000
        .ColWidth(2) = 4000
    
        .ColAlignment(2) = flexAlignLeftCenter
        
        .TextMatrix(0, 0) = "Column Name"
        .TextMatrix(0, 1) = "Type"
        .TextMatrix(0, 2) = "Size"
        If Not frmMain.adoData.Recordset Is Nothing Then
            .Rows = frmMain.adoData.Recordset.Fields.Count + 1
            
            For nCount = 1 To frmMain.adoData.Recordset.Fields.Count
                .TextMatrix(nCount, 0) = frmMain.adoData.Recordset.Fields(nCount - 1).Name
                .TextMatrix(nCount, 1) = ConvType(frmMain.adoData.Recordset.Fields(nCount - 1).Type)
                .TextMatrix(nCount, 2) = frmMain.adoData.Recordset.Fields(nCount - 1).DefinedSize
            Next nCount
        End If
    End With
                          
End Sub

Private Function ConvType(ByVal TypeVal As Long) As String
  Select Case TypeVal
        Case adBigInt                    ' 20
            ConvType = "Big Integer"
        Case adBinary                    ' 128
            ConvType = "Binary"
        Case adBoolean                   ' 11
            ConvType = "Boolean"
        Case adBSTR                      ' 8 i.e. null terminated string
            ConvType = "Text"
        Case adChar                      ' 129
            ConvType = "Text"
        Case adCurrency                  ' 6
            ConvType = "Currency"
        Case adDate                      ' 7
            ConvType = "Date/Time"
        Case adDBDate                    ' 133
            ConvType = "Date/Time"
        Case adDBTime                    ' 134
            ConvType = "Date/Time"
        Case adDBTimeStamp               ' 135
            ConvType = "Date/Time"
        Case adDecimal                   ' 14
            ConvType = "Float"
        Case adDouble                    ' 5
            ConvType = "Float"
        Case adEmpty                     ' 0
            ConvType = "Empty"
        Case adError                     ' 10
            ConvType = "Error"
        Case adGUID                      ' 72
            ConvType = "GUID"
        Case adIDispatch                 ' 9
            ConvType = "IDispatch"
        Case adInteger                   ' 3
            ConvType = "Integer"
        Case adIUnknown                  ' 13
            ConvType = "Unknown"
        Case adLongVarBinary             ' 205
            ConvType = "Binary"
        Case adLongVarChar               ' 201
            ConvType = "Text"
        Case adLongVarWChar              ' 203
            ConvType = "Text"
        Case adNumeric                  ' 131
            ConvType = "Long"
        Case adSingle                    ' 4
            ConvType = "Single"
        Case adSmallInt                  ' 2
            ConvType = "Small Integer"
        Case adTinyInt                   ' 16
            ConvType = "Tiny Integer"
        Case adUnsignedBigInt            ' 21
            ConvType = "Big Integer"
        Case adUnsignedInt               ' 19
            ConvType = "Integer"
        Case adUnsignedSmallInt          ' 18
            ConvType = "Small Integer"
        Case adUnsignedTinyInt           ' 17
            ConvType = "Timy Integer"
        Case adUserDefined               ' 132
            ConvType = "UserDefined"
        Case adVarNumeric                 ' 139
            ConvType = "Long"
        Case adVarBinary                 ' 204
            ConvType = "Binary"
        Case adVarChar                   ' 200
            ConvType = "Text"
        Case adVariant                   ' 12
            ConvType = "Variant"
        Case adVarWChar                  ' 202
            ConvType = "Text"
        Case adWChar                     ' 130
            ConvType = "Text"
        Case Else
            ConvType = "Unknown"
   End Select
End Function

