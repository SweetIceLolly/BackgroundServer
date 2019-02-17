Attribute VB_Name = "modAPI"
Option Explicit

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
    ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
Private Const WM_CLIPBOARDUPDATE = &H31D    'Keyboard change message code

Public WindowListCode       As String       'Window list HTML code for window enumerations
Public WindowList(2)        As String       'Window list page
Public PrevWndProc          As Long         'Previous address of window procedure
Public ClipboardChangeTime  As String       'Last change time of clipboard
Public hkKeyboard           As Long         'Hook handle of keyboard hook
Public hkMouse              As Long         'Hook handle of mouse hook
Public kLastTime            As String       'Last keyboard event time
Public mLastTime            As String       'Last mouse event time
Public PrevFocusedWindow    As Long         'Focused window handle in the last second
Public FocusChangeTime      As String       'Last window focus change time

'Purpose:   Window procedure, to handle WM_CLIPBOARDUPDATE message for the window
'Args:      hWnd: Window handle; uMsg: Message code; wParam, lParam: Additional message info
Public Function WndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If uMsg = WM_CLIPBOARDUPDATE Then
        'Record the change time of clipboard when it's changed
        ClipboardChangeTime = Format(Now, "yyyy/mm/dd hh:mm:ss")
    End If
    WndProc = CallWindowProc(PrevWndProc, hwnd, uMsg, wParam, lParam)
End Function

'Purpose:   Hook procedure for keyboard hook, to record last time of keyboard event
'Args:      nCode: Message code; wParam, lParam: Additional message info
Public Function KeyboardHookProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    'Record the time of the keyboard event
    kLastTime = Format(Now, "yyyy/mm/dd hh:mm:ss")
    KeyboardHookProc = CallNextHookEx(hkKeyboard, nCode, wParam, lParam)
End Function

'Purpose:   Hook procedure for mouse hook, to record last time of mouse event
'Args:      nCode: Message code; wParam, lParam: Additional message info
Public Function MouseHookProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    'Record the time of the mouse event
    mLastTime = Format(Now, "yyyy/mm/dd hh:mm:ss")
    MouseHookProc = CallNextHookEx(hkMouse, nCode, wParam, lParam)
End Function

'Purpose:   To add info of the specified window into the list
'Args:      hWnd: The handle of the window whose info will be retrieved
Private Sub AddWindowInfo(ByVal hwnd As Long)
    Dim WindowName  As String * 255         'Window caption
    Dim ClassName   As String * 255         'Window class name
    Dim PID         As Long                 'Process ID of the window
    
    GetWindowThreadProcessId hwnd, PID                  'Retrieve the process id
    GetClassNameA hwnd, ClassName, 255                  'Retrieve the class name
    GetWindowTextA hwnd, WindowName, 255                'Retrieve the window name
    
    WindowListCode = WindowListCode & Replace(Replace(Replace(Replace(Replace(Replace(WindowList(1), _
        "¡¾HWND¡¿", "0x" & Hex(hwnd)), _
        "¡¾VISIBLE¡¿", IsWindowVisible(hwnd) <> 0), _
        "¡¾CLASS¡¿", Left(ClassName, InStr(ClassName, vbNullChar) - 1)), _
        "¡¾TEXT¡¿", Left(WindowName, InStr(WindowName, vbNullChar) - 1)), _
        "¡¾PID¡¿", PID), _
        "¡¾HWND_DEC¡¿", hwnd)
End Sub

'Purpose:   The callback procedure of EnumWindows() function
'Args:      hWnd: Current window handle
'           lParam: Additional information.
'Return:    A BOOL type value, TRUE means keep enumerate windows, FALSE means exit the enumeration
Public Function EnumProc(ByVal hwnd As Long, lParam As Long) As BOOL
    AddWindowInfo hwnd
    EnumProc = 1
End Function
