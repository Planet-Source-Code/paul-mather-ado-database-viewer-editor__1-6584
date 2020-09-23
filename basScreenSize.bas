Attribute VB_Name = "basScreenSize"
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long

Public Const SPI_GETWORKAREA = 48
            
Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Property Get WorkAreaHeight() As Long
Dim apiRECT As RECT
    Call GetWorkAreaDimensions(apiRECT)
    WorkAreaHeight = (apiRECT.Bottom - apiRECT.Top) * Screen.TwipsPerPixelX
End Property
Public Property Get WorkAreaWidth() As Long
Dim apiRECT As RECT
    Call GetWorkAreaDimensions(apiRECT)
    WorkAreaWidth = (apiRECT.Right - apiRECT.Left) * Screen.TwipsPerPixelY
End Property
Public Property Get WorkAreaTop() As Long
Dim apiRECT As RECT
    Call GetWorkAreaDimensions(apiRECT)
    WorkAreaTop = apiRECT.Top * Screen.TwipsPerPixelX
End Property
Public Property Get WorkAreaBottom() As Long
Dim apiRECT As RECT
    Call GetWorkAreaDimensions(apiRECT)
    WorkAreaBottom = apiRECT.Bottom * Screen.TwipsPerPixelX
End Property
Public Property Get WorkAreaLeft() As Long
Dim apiRECT As RECT
    Call GetWorkAreaDimensions(apiRECT)
    WorkAreaLeft = apiRECT.Left * Screen.TwipsPerPixelY
End Property
Public Property Get WorkAreaRight() As Long
Dim apiRECT As RECT
    Call GetWorkAreaDimensions(apiRECT)
    WorkAreaRight = apiRECT.Right * Screen.TwipsPerPixelY
End Property

Private Sub GetWorkAreaDimensions(ByRef apiRECT As RECT)
Dim lRet As Long
    lRet = SystemParametersInfo(SPI_GETWORKAREA, vbNull, apiRECT, 0)
End Sub





