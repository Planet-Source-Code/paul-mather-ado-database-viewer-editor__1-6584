Attribute VB_Name = "basMessageBox"
Declare Function MessageBoxEx Lib "user32" Alias "MessageBoxExA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal uType As Long, ByVal wLanguageId As Long) As Long
Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function GetCurrentThreadId Lib "KERNEL32" () As Long
Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Const GWL_HINSTANCE = (-6)
Const SWP_NOSIZE = &H1
Const SWP_NOZORDER = &H4
Const SWP_NOACTIVATE = &H10
Const HCBT_ACTIVATE = 5
Const WH_CBT = 5

Type RECT
    left As Long
    top As Long
    Right As Long
    Bottom As Long
End Type

Dim hHook As Long
Dim parenthWnd As Long

Public Function MessageBox(ByVal hwnd As Long, ByVal Prompt As String, Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, Optional ByVal Title As String = "", Optional ByVal HelpFile As String, Optional ByVal Context, Optional ByVal centerForm As Boolean = True) As VbMsgBoxResult
Dim ret As Long
    Dim hInst As Long
    Dim Thread As Long
    'Set up the CBT hook
    parenthWnd = hwnd
    hInst = GetWindowLong(hwnd, GWL_HINSTANCE)
    Thread = GetCurrentThreadId()
    If centerForm = True Then
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterForm, hInst, Thread)
    Else
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterScreen, hInst, Thread)
    End If
    
    ret = MessageBoxEx(hwnd, Prompt, Title, Buttons, 0)
    MessageBox = ret
End Function

Private Function WinProcCenterScreen(ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim rectForm As RECT, rectMsg As RECT
    Dim x As Long, y As Long
    If lMsg = HCBT_ACTIVATE Then
        'Show the MsgBox at a fixed location (0,0)
        GetWindowRect wParam, rectMsg
        x = Screen.Width / Screen.TwipsPerPixelX / 2 - (rectMsg.Right - rectMsg.left) / 2
        y = Screen.Height / Screen.TwipsPerPixelY / 2 - (rectMsg.Bottom - rectMsg.top) / 2
        SetWindowPos wParam, 0, x, y, 0, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
        'Release the CBT hook
        UnhookWindowsHookEx hHook
    End If
    WinProcCenterScreen = False
End Function

Private Function WinProcCenterForm(ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim rectForm As RECT, rectMsg As RECT
    Dim x As Long, y As Long
    'On HCBT_ACTIVATE, show the MsgBox centered over Form1
    If lMsg = HCBT_ACTIVATE Then
        'Get the coordinates of the form and the message box so that
        'you can determine where the center of the form is located
        GetWindowRect parenthWnd, rectForm
        GetWindowRect wParam, rectMsg
        x = (rectForm.left + (rectForm.Right - rectForm.left) / 2) - ((rectMsg.Right - rectMsg.left) / 2)
        y = (rectForm.top + (rectForm.Bottom - rectForm.top) / 2) - ((rectMsg.Bottom - rectMsg.top) / 2)
        'Position the msgbox
        SetWindowPos wParam, 0, x, y, 0, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
        'Release the CBT hook
        UnhookWindowsHookEx hHook
     End If
     WinProcCenterForm = False
End Function

