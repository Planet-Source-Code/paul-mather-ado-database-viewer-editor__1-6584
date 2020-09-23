Attribute VB_Name = "basLockFormSize"
Option Explicit

Private Const GWL_WNDPROC = -4
Private Const WM_GETMINMAXINFO = &H24
Dim minWidth As Long
Dim minHeight As Long
Dim maxWidth As Long
Dim maxHeight As Long

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type

Global lpPrevWndProc As Long
Global gHW As Long

Private Declare Function GetModuleFileName Lib "KERNEL32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Sub CopyMemoryToMinMaxInfo Lib "KERNEL32" Alias "RtlMoveMemory" (hpvDest As MINMAXINFO, ByVal hpvSource As Long, ByVal cbCopy As Long)
Private Declare Sub CopyMemoryFromMinMaxInfo Lib "KERNEL32" Alias "RtlMoveMemory" (ByVal hpvDest As Long, hpvSource As MINMAXINFO, ByVal cbCopy As Long)

Public Sub Hook(hWnd As Long, minimumWidth As Long, minimumHeight As Long, Optional maximumWidth As Long = 0, Optional maximumHeight As Long = 0)
    If DebugMode = False Then
        'Start subclassing.
        gHW = hWnd
        minWidth = minimumWidth
        minHeight = minimumHeight
        maxWidth = maximumWidth
        maxHeight = maximumHeight
            
        lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, AddressOf WindowProc)
    End If
End Sub

Public Sub Unhook()
    Dim temp As Long

    If DebugMode = False Then
        'Cease subclassing.
        temp = SetWindowLong(gHW, GWL_WNDPROC, lpPrevWndProc)
    End If
End Sub

Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim MinMax As MINMAXINFO

    'Check for request for min/max window sizes.
    If uMsg = WM_GETMINMAXINFO Then
        'Retrieve default MinMax settings
        CopyMemoryToMinMaxInfo MinMax, lParam, Len(MinMax)

        'Specify new minimum size for window.
        MinMax.ptMinTrackSize.x = minWidth / Screen.TwipsPerPixelX
        MinMax.ptMinTrackSize.y = minHeight / Screen.TwipsPerPixelY

        If maxWidth <> 0 Then
            'Specify new maximum size for window.
            MinMax.ptMaxTrackSize.x = maxWidth / Screen.TwipsPerPixelX
            MinMax.ptMaxTrackSize.y = maxHeight / Screen.TwipsPerPixelY
        End If

        'Copy local structure back.
        CopyMemoryFromMinMaxInfo lParam, MinMax, Len(MinMax)

        WindowProc = DefWindowProc(hw, uMsg, wParam, lParam)
    Else
        WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, lParam)
    End If
End Function

Public Property Get DebugMode() As Boolean
Dim strFileName As String
Dim lngCount As Long
    strFileName = String(255, 0)
    lngCount = GetModuleFileName(App.hInstance, strFileName, 255)
    strFileName = left(strFileName, lngCount)
    If UCase(Right(strFileName, 7)) <> "VB6.EXE" Then
        DebugMode = False
    Else
        DebugMode = True
    End If
End Property



