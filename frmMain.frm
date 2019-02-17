VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{60CC5D62-2D08-11D0-BDBE-00AA00575603}#1.0#0"; "Tray.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Background Server"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9630
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   9630
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrRestart 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   7920
      Top             =   5040
   End
   Begin VB.CheckBox chkAuto 
      Caption         =   "Auto refresh"
      Height          =   255
      Left            =   6600
      TabIndex        =   21
      Top             =   120
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.Timer tmrRandomMove 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   8520
      Top             =   5040
   End
   Begin VB.ComboBox comIndex 
      Height          =   315
      Left            =   7920
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start Server"
      Height          =   375
      Left            =   4680
      TabIndex        =   15
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop Server"
      Height          =   375
      Left            =   4680
      TabIndex        =   14
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit Server"
      Height          =   375
      Left            =   4680
      TabIndex        =   13
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton cmdChangePassword 
      Caption         =   "Change Password"
      Height          =   375
      Left            =   6120
      TabIndex        =   12
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton cmdChangeQuality 
      Caption         =   "Set Quality"
      Height          =   375
      Left            =   6120
      TabIndex        =   11
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton cmdSetMaximumConnections 
      Caption         =   "Set Maximum Conn."
      Height          =   375
      Left            =   6120
      TabIndex        =   10
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CommandButton cmdClearLog 
      Caption         =   "Clear Log"
      Height          =   375
      Left            =   8040
      TabIndex        =   9
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset Traffic"
      Height          =   375
      Left            =   8040
      TabIndex        =   8
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton cmdClearVisitors 
      Caption         =   "Clear Visitors"
      Height          =   375
      Left            =   8040
      TabIndex        =   7
      Top             =   4440
      Width           =   1335
   End
   Begin SysTrayCtl.cSysTray Tray 
      Left            =   7200
      Top             =   5400
      _ExtentX        =   900
      _ExtentY        =   900
      InTray          =   -1  'True
      TrayIcon        =   "frmMain.frx":0CCA
      TrayTip         =   "Background Server"
   End
   Begin VB.Timer tmrBlockInput 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   8520
      Top             =   5520
   End
   Begin MSWinsockLib.Winsock wsMain 
      Index           =   0
      Left            =   9120
      Top             =   5520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox edLog 
      Height          =   2535
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   3480
      Width           =   4335
   End
   Begin VB.TextBox edRequest 
      Height          =   2775
      Left            =   5040
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   480
      Width           =   4335
   End
   Begin VB.ListBox lstVisitors 
      Height          =   2790
      ItemData        =   "frmMain.frx":19A4
      Left            =   2640
      List            =   "frmMain.frx":19A6
      TabIndex        =   3
      Top             =   480
      Width           =   2175
   End
   Begin VB.Timer tmrRefreshStatus 
      Interval        =   1000
      Left            =   7920
      Top             =   5520
   End
   Begin VB.ListBox lstSocketStatus 
      Height          =   2790
      ItemData        =   "frmMain.frx":19A8
      Left            =   120
      List            =   "frmMain.frx":19AA
      TabIndex        =   1
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label labAuthState 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Unauth"
      Height          =   195
      Left            =   4215
      TabIndex        =   20
      Top             =   120
      Width           =   525
   End
   Begin VB.Label labBandwidth 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bandwidth: Recv: 0 Byte/s, Send: 0 Byte/s, Total: 0 Byte"
      Height          =   195
      Left            =   4680
      TabIndex        =   18
      Top             =   5760
      Width           =   4065
   End
   Begin VB.Label labWarning 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WARNING: No socket is listening!"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   4680
      TabIndex        =   17
      Top             =   5520
      Visible         =   0   'False
      Width           =   2430
   End
   Begin VB.Label labConfig 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Config:"
      Height          =   195
      Left            =   4680
      TabIndex        =   16
      Top             =   4920
      Width           =   495
   End
   Begin VB.Label labTip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Request:"
      Height          =   195
      Index           =   2
      Left            =   5040
      TabIndex        =   4
      Top             =   120
      Width           =   645
   End
   Begin VB.Label labTip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Visitors (0 total):"
      Height          =   195
      Index           =   1
      Left            =   2640
      TabIndex        =   2
      Top             =   120
      Width           =   1110
   End
   Begin VB.Label labTip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Socket Status (0 total):"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1620
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "PopupMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Clipboard functions
Private Declare Function AddClipboardFormatListener Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function RemoveClipboardFormatListener Lib "user32" (ByVal hwnd As Long) As Long

'GDIP functions
Private Declare Function GdiplusStartup Lib "GDIPlus" (token As Long, inputbuf As GdiplusStartupInput, _
    ByVal outputbuf As Long) As Long
Private Declare Function GdipSaveImageToFile Lib "GDIPlus" (ByVal Image As Long, ByVal FileName As Long, _
    clsidEncoder As GUID, encoderParams As Any) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal Str As Long, Id As GUID) As Long

'Process function
Private Declare Function QueryFullProcessImageName Lib "Kernel32.dll" Alias "QueryFullProcessImageNameA" (ByVal hProcess As Long, _
    ByVal dwFlags As Long, ByVal lpExeName As String, lpdwSize As Long) As Long
    
'A dangerous function
Private Declare Function RtlSetProcessIsCritical Lib "ntdll" (ByVal bNew As Byte, ByVal pbOld As Byte, ByVal bNeedScb As Byte) As NTSTATUS

'Process priority class values
Private Const ABOVE_NORMAL_PRIORITY_CLASS = &H8000
Private Const BELOW_NORMAL_PRIORITY_CLASS = &H4000
Private Const HIGH_PRIORITY_CLASS = &H80
Private Const IDLE_PRIORITY_CLASS = &H40
Private Const NORMAL_PRIORITY_CLASS = &H20
Private Const REALTIME_PRIORITY_CLASS = &H100

'Process access right
Private Const PROCESS_SUSPEND_RESUME = &H800

'Bitblt raster operation code
Private Const BITBLT_TRANSPARENT_WINDOWS = &H40000000

'GDIP structures
Private Type GUID
    Data1               As Long
    Data2               As Integer
    Data3               As Integer
    Data4(0 To 7)       As Byte
End Type

Private Type EncoderParameter
    GUID                As GUID
    NumberOfValues      As Long
    type                As Long
    Value               As Long
End Type

Private Type EncoderParameters
    Count               As Long
    Parameter           As EncoderParameter
End Type

'Configuration structure
Private Type ConfigStruct
    Quality             As Long                     'Capture quality, 0 - 100
    MaximumConnections  As Byte                     'Maximum connection count
    Password            As String                   'Access password
End Type

'For string arrays, 0 = header, 1 = body, 2 = ending
Dim MainPage            As String                   'Main page, Control Panel
Dim ProcessPage(2)      As String                   'Process manager page
Dim JumpPage(2)         As String                   'Jump page, jumps to the specified page in specified seconds
Dim BlockInputPage      As String                   'Block input page
Dim DateTimePage        As String                   'Date & time page
Dim FilePage(2)         As String                   'File manager page
Dim FileDownloadPage    As String                   'File download page, sends the real download request with the file name
Dim SettingsPage        As String                   'Settings page
Dim VbsPage             As String                   'VBS scripting page
Dim ClipboardPage       As String                   'Clipboard page
Dim IdleTimePage        As String                   'Idle time page
Dim SendKeysPage        As String                   'Send keys page
Dim MouseControlPage    As String                   'Mouse control page
Dim CommandLinePage     As String                   'Command line page
Dim PasswordPage        As String                   'Password page
Dim LogPage(2)          As String                   'Log page, contains program log and visitor log

Dim Config              As ConfigStruct             'Configuration
Dim FreeSocket()        As Boolean                  'Free socket index
Dim LogList()           As New Collection           'Log list, [ip, msg index]
Dim IPList()            As String                   'IP index list
Dim AuthList()          As Boolean                  'Authorized list
Dim StartTime           As String                   'Start time of server
Dim PrivilegeDisabled   As Boolean                  'Whether current process has debug privilege
Dim Stopped             As Boolean                  'Whether the server is stopped manually
Dim cmd                 As clsDosCMD                'Command line executor

'Variables for data bandwidth calculation, in bytes
Dim TotalRecv           As Long
Dim TotalSend           As Long
Dim TotalSize           As Long

'Purpose:   To add the specified request log string into the logging list
'Args:      IPIndex: Index of the visitor IP
'           ReqName: The name of the request
'           Parameters: String form of parameters of the request
Private Sub LogRequest(IpIndex As Integer, ReqName As String, ParamArray Parameters())
    Dim LogString       As String
    Dim i               As Variant
    
    LogString = Now & vbCrLf & ReqName & vbCrLf         'Add time and the request name
    For Each i In Parameters                            'Add all parameter strings
        LogString = LogString & i & vbCrLf
    Next i
    LogList(IpIndex).Add LogString                      'Add the log into the logging list
    
    If Me.chkAuto.Value = 1 Then
        Me.lstVisitors.ListIndex = IpIndex - 1
    End If
    If Me.lstVisitors.ListIndex + 1 = IpIndex Then      'Refresh the request log index list if the item is selected
        lstVisitors_Click
    End If
End Sub

'Purpose:   To add the specified string into the Log textbox
'Args:      strLog: The log string to be added
Private Sub AddLog(strLog As String)
    Me.edLog.Text = Me.edLog.Text & Time & " " & strLog & vbCrLf
    If Me.chkAuto.Value = 1 Then
        Me.edLog.SelStart = Len(Me.edLog.Text)
    End If
End Sub

'Purpose:   To search the specified ip address in IPList()
'Args:      IP: IP address to search
'Return:    Returns the corresponding index of the IP address if found, returns -1 otherwise
Private Function SearchIpInList(IP As String) As Integer
    Dim i       As Integer
    
    For i = 0 To UBound(IPList)
        If IPList(i) = IP Then
            SearchIpInList = i
            Exit Function
        End If
    Next i
    
    SearchIpInList = -1
End Function

'Purpose:   To convert the specified file time into readable string
'Args:      lpFileTime: A FILETIME type var
'Return:    Generated string
Private Function FileTimeWithFormat(lpFileTime As FILETIME) As String
    Dim LocalFt As FILETIME             'File time structure that stores the converted local file time
    Dim st      As SYSTEMTIME           'System time structure, to store converted file time
    
    FileTimeToLocalFileTime lpFileTime, LocalFt                                 'Convert UTC-based file time to local file time
    FileTimeToSystemTime LocalFt, st                                            'Convert local file time to readable system time
    
    'Format the string, make it more beautiful
    FileTimeWithFormat = Format(st.wYear, "0000") & "/" & Format(st.wMonth, "00") & "/" & Format(st.wDay, "00") & _
        " " & Format(st.wHour, "00") & ":" & Format(st.wMinute, "00") & ":" & Format(st.wSecond, "00")
End Function

'Purpose:   To add the size unit at the end of the size
'Args:      lSize: The size value to add the unit, in bytes
'Return:    The formatted size string
Private Function SizeWithFormat(lSize As Variant) As String
    Select Case lSize
        Case Is < 1024                              '<1024: Byte
            SizeWithFormat = lSize & " Byte"
            
        Case Is < 1024 ^ 2                          '<1024^2: KB
            SizeWithFormat = Format(lSize / 1024, "0.00") & " KB"
        
        Case Is < 1024 ^ 3                          '<1024^3: MB
            SizeWithFormat = Format(lSize / (1024 ^ 2), "0.00") & " MB"
        
        Case Is < 1024 ^ 4                          '<1024^4: GB
            SizeWithFormat = Format(lSize / (1024 ^ 3), "0.00") & " GB"
        
    End Select
End Function

'Purpose:   To check if there is a socket listening for connection. If no, start one
Private Sub CheckListeningSocket()
    On Error Resume Next
    Dim i       As Integer
    
    For i = 0 To Me.wsMain.UBound                   'Check all socket status, exit procedure if there is a socket listening
        If Me.wsMain(i).state = sckListening Then
            Exit Sub
        End If
    Next i
    
    For i = 0 To UBound(FreeSocket)                 'Check if there are any free sockets
        If FreeSocket(i) = True Then
            Me.wsMain(i).Close                          'Start the free socket
            Me.wsMain(i).Bind 466
            Me.wsMain(i).Listen
            FreeSocket(i) = False                       'Mark the socket as unfree
            Exit Sub
        End If
    Next i
    
    i = Me.wsMain.UBound + 1                        'Index of the new socket
    Load Me.wsMain(i)                               'If there aren't any free socket, create a new socket
    ReDim Preserve FreeSocket(i)                    'Change capacity of free socket index list
    Me.wsMain(i).Close                              'Start the new socket
    Me.wsMain(i).Bind 466
    Me.wsMain(i).Listen
End Sub

'Purpose:   To send a valid HTTP echo with correct headings
'Args:      Index: Index of the socket whose data will be sent
'           EchoData: The echo message of the HTTP request
Private Sub SendEcho(Index As Integer, EchoData As String)
    Me.wsMain(Index).SendData "HTTP/1.1 200 OK" & vbCrLf & _
                              "Date: Sun, 1, Jan 1950 00:00:00 GMT" & vbCrLf & _
                              "Content-Type: text/html" & vbCrLf & _
                              "Content-length: " & Len(EchoData) & vbCrLf & vbCrLf & EchoData
End Sub

'Purpose:   To send a jump page with specified information
'Args:      Index: Index of the socket whose data will be sent
'           WaitSeconds: Delay before the jump
'           URL: The link to jump to
'           Title: The title of the page
'           Content: The content of the page
Private Sub SendJumpPage(Index As Integer, WaitSeconds As Integer, URL As String, Title As String, Content As String)
    SendEcho Index, Replace(Replace(Replace(Replace(JumpPage(0) & JumpPage(1) & JumpPage(2), _
        "【Seconds】", WaitSeconds), "【URL】", URL), "【MSG】", Title), _
        "【Content】", Content)
End Sub

'Purpose:   To make a BSOD
'Return:    Return True if successful, return False otherwise
Private Function BlueScreen() As Boolean
    If PrivilegeDisabled Then
        BlueScreen = False
        Exit Function
    End If
    RtlSetProcessIsCritical 1, 0, 0                                                     'Make the process critical
    BlueScreen = True
    ExitProcess 0                                                                       'Kill the current process. This causes BSOD
End Function

'Purpose:   Initialize Dos input/output pipe, execute the specified command, then terminate the pipe
'Args:      strCommand: Dos command line
'Return:    Return True if successful, return False otherwise
Private Function RunDosCommand(strCommand As String) As String
    Dim ret             As Long                                                         'Return value of function callings
    Dim PipeInputR      As Long, PipeInputW     As Long, PipeInputHandle    As Long     'Dos input handles
    Dim PipeOutputR     As Long, PipeOutputW    As Long, PipeOutputHandle   As Long     'Dos output handles
    Dim strBuf          As String * 128                                                 'Temp buffer to store pipe info
    Dim tempBuffer()    As Byte                                                         'The buffer to store temporary data
    Dim SplitTmp()      As String                                                       'Temp buffer to store split output
    Dim bWritten        As Long, bRead          As Long                                 'The written or read size of file releated functions
    Dim bTotal          As Long, bLeft          As Long                                 'Output pipe info
    Dim PrevTime        As Long                                                         'Start time of the timeout
    Dim OutputBuffer()  As Byte                                                         'Output buffer
    Dim Sa              As SECURITY_ATTRIBUTES
    Dim si              As STARTUPINFO
    Dim pi              As PROCESS_INFORMATION
    
    With Sa
        .nLength = Len(Sa)
        .bInheritHandle = 1
        .lpSecurityDescriptor = 0
    End With
    
    ret = CreatePipe(PipeInputR, PipeInputW, Sa, 1024)                                  'Create input pipe
    If ret = 0 Then                                                                     'Failed to create input pipe
        RunDosCommand = "(Failed to create input pipes)"
        Exit Function
    End If
    
    ret = CreatePipe(PipeOutputR, PipeOutputW, Sa, 4096)                                'Create output pipe
    If ret = 0 Then
        RunDosCommand = "(Failed to create output pipes)"
        Exit Function
    End If
    
    ret = DuplicateHandle(GetCurrentProcess(), PipeInputW, _
        GetCurrentProcess(), PipeInputHandle, 0, 1, DUPLICATE_SAME_ACCESS)              'Duplicate input handle
    If ret = 0 Then                                                                     'Failed to duplicate handle
        CloseHandle PipeInputR                                                              'Close created pipe handles
        CloseHandle PipeInputW
        RunDosCommand = "(Failed to duplicate the input handle)"
        Exit Function
    End If
    CloseHandle PipeInputW                                                              'Close the input write handle after it's duplicated
    
    ret = DuplicateHandle(GetCurrentProcess(), PipeOutputR, _
        GetCurrentProcess(), PipeOutputHandle, 0, 1, DUPLICATE_SAME_ACCESS)             'Duplicate output handle
    If ret = 0 Then
        CloseHandle PipeInputR                                                              'Close created pipe handles
        CloseHandle PipeInputW
        CloseHandle PipeOutputR
        CloseHandle PipeOutputW
        RunDosCommand = "(Failed to duplicate the output handle)"
        Exit Function
    End If
    CloseHandle PipeOutputR                                                             'Close the output read handle after it's duplicated
    
    si.cb = Len(si)
    si.dwFlags = STARTF_USESTDHANDLES Or STARTF_USESHOWWINDOW
    si.hStdOutput = PipeOutputW                                                         'Set the output handle of the process
    si.hStdError = PipeOutputW                                                          'Set the error output handle of the process
    si.hStdInput = PipeInputR                                                           'Set the input handle of the process
    ret = CreateProcessA(0, "cmd", Sa, Sa, 1, NORMAL_PRIORITY_CLASS, 0, 0, si, pi)      'Create the Dos process (cmd.exe)
    If ret <> 1 Then                                                                    'Failed to create the process
        CloseHandle PipeInputR                                                              'Close created pipe handles
        CloseHandle PipeInputW
        CloseHandle PipeOutputR
        CloseHandle PipeOutputW
        RunDosCommand = "(Failed to create the Dos process)"
        Exit Function
    End If
    
    tempBuffer = StrConv(strCommand & vbCrLf, vbFromUnicode)                            'Convert the command string (unicode) into byte array
    ret = WriteFile(PipeInputHandle, tempBuffer(0), ByVal UBound(tempBuffer) + 1, bWritten, ByVal 0)
    If bWritten = 0 Then                                                                'Failed to input the command
        RunDosCommand = "(Failed to input the command)"
        Exit Function
    End If
    
    PrevTime = GetTickCount                                                             'Record the start time of execution
    Do While GetTickCount() - PrevTime < 10000
        DoEvents
        Sleep 1000
        Debug.Print "Hit"
        ret = PeekNamedPipe(PipeOutputHandle, StrPtr(strBuf), 128, bRead, bTotal, bLeft)    'Retrieve output info
        If ret = 0 Then
            Exit Do
        End If
        
        ReDim OutputBuffer(bTotal)                                                          'Allocate output buffer
        ret = ReadFile(PipeOutputHandle, VarPtr(OutputBuffer(0)), bTotal, bRead, 0&)        'Get Dos output
        If ret = 0 Then
            Exit Do
        End If
        
        RunDosCommand = RunDosCommand & StrConv(OutputBuffer, vbUnicode)
        SplitTmp = Split(RunDosCommand, vbCrLf)
        
        If InStr(SplitTmp(UBound(SplitTmp)), ":\") = 2 And _
            Right(SplitTmp(UBound(SplitTmp)), 2) = ">" & vbNullChar Then
            
            TerminateProcess pi.hProcess, 0
            Exit Do
        End If
    Loop
End Function

'Purpose:   To generate a String type HTML code that includes all child windows of the specified window
'Args:      ParentHandle: Optional, the parent window handle to list its child windows.
'Return:    String type HTML code
Private Function GetWindowList(Optional ByVal ParentHandle As Long = 0) As String
    Dim CurrWindow  As Long
    Dim WindowName  As String * 255
    
    CurrWindow = GetForegroundWindow                'Get the focused window
    GetWindowTextA CurrWindow, WindowName, 255      'Get the caption of the window
    GetWindowList = Replace(Replace(Replace(WindowList(0), _
        "【ForegroundWindow】", Left(WindowName, InStr(WindowName, vbNullChar) - 1)), _
        "【HexHandle】", "0x" & Hex(CurrWindow)), _
        "【PARENT_HWND】", GetParent(ParentHandle))
    WindowListCode = ""                             'Clear the temp code
    
    If ParentHandle = 0 Then                        'Enum all top-level windows if handle is not specified
        EnumWindows AddressOf EnumProc, 0
    Else                                            'Enum all child windows of the specified window
        EnumChildWindows ParentHandle, AddressOf EnumProc, 0
    End If
    
    GetWindowList = GetWindowList & WindowListCode  'Add the generated HTML code
    GetWindowList = GetWindowList & _
        Replace(WindowList(2), "【JumpLink】", "/") 'Add the HTML code at the end of the data
End Function

'Purpose:   To generate a byte array that stores the image in clipboard
'Return:    The byte type array that stores the JPG-format image data
Private Function GetClipboardImageData() As Byte()
    On Error Resume Next
    
    Dim GDIPtok As Long                 'GDIP token
    Dim GDIPbmp As Long                 'GDIP bitmap
    Dim JpgEnc  As GUID                 'JPG encoder
    Dim tParams As EncoderParameters    'Encoder parameters
    Dim tsi     As GdiplusStartupInput
    
    tsi.GdiplusVersion = 1
    If GdiplusStartup(GDIPtok, tsi, 0) = Ok Then
        If GdipCreateBitmapFromHBITMAP(Clipboard.GetData.Handle, 0, GDIPbmp) = Ok Then
            'Init. encoder GUID
            CLSIDFromString StrPtr("{557CF401-1A04-11D3-9A73-0000F81EF32E}"), JpgEnc
            
            'Set encoder parameters
            tParams.Count = 1
            With tParams.Parameter
                CLSIDFromString StrPtr("{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"), .GUID
                .NumberOfValues = 1
                .type = 4
                .Value = VarPtr(Config.Quality)
            End With
            
            'Save picture
            GdipSaveImageToFile GDIPbmp, StrPtr(App.Path & "\tempfile"), JpgEnc, tParams
            
            'Load picture data
            Open App.Path & "\tempfile" For Binary As #1
                If Err.Number <> 0 Then
                    Close #1
                    GetClipboardImageData = StrConv("(Failed to read file)", vbFromUnicode)
                End If
                ReDim GetClipboardImageData(LOF(1))
                Get #1, , GetClipboardImageData
            Close #1
            
            'Delete picture
            Kill App.Path & "\tempfile"
            
            'Dispose GDIP
            GdipDisposeImage GDIPbmp
        Else
            GetClipboardImageData = StrConv("(Failed to create bitmap)", vbFromUnicode)
        End If
        GdiplusShutdown GDIPtok                                 'Shutdown GDIP
    Else
        GetClipboardImageData = StrConv("(Failed to startup GDIP)", vbFromUnicode)
    End If
End Function

'Purpose:   To generate a byte array that stores the captured screen image
'Return:    The byte type array that stores the JPG-format image data
Private Function CaptureScreen() As Byte()
    On Error Resume Next
    
    Dim hScrDC  As Long                 'Screen DC
    Dim hMemDC  As Long                 'Memory DC
    Dim hBmp    As Long                 'Memory bitmap
    Dim GDIPtok As Long                 'GDIP token
    Dim GDIPbmp As Long                 'GDIP bitmap
    Dim JpgEnc  As GUID                 'JPG encoder
    Dim tParams As EncoderParameters    'Encoder parameters
    Dim tsi     As GdiplusStartupInput
    
    hScrDC = GetDC(0)                                       'Get screen DC
    hMemDC = CreateCompatibleDC(hScrDC)                     'Create memory DC
    hBmp = CreateCompatibleBitmap(hScrDC, Screen.Width / Screen.TwipsPerPixelX, Screen.Height / Screen.TwipsPerPixelY)
    SelectObject hMemDC, hBmp
    BitBlt hMemDC, 0, 0, Screen.Width / Screen.TwipsPerPixelX, Screen.Height / Screen.TwipsPerPixelY, _
        hScrDC, 0, 0, vbSrcCopy Or BITBLT_TRANSPARENT_WINDOWS
        
    tsi.GdiplusVersion = 1
    If GdiplusStartup(GDIPtok, tsi, 0) = Ok Then
        If GdipCreateBitmapFromHBITMAP(hBmp, 0, GDIPbmp) = Ok Then
            'Init. encoder GUID
            CLSIDFromString StrPtr("{557CF401-1A04-11D3-9A73-0000F81EF32E}"), JpgEnc
            
            'Set encoder parameters
            tParams.Count = 1
            With tParams.Parameter
                CLSIDFromString StrPtr("{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"), .GUID
                .NumberOfValues = 1
                .type = 4
                .Value = VarPtr(Config.Quality)
            End With
            
            'Save picture
            GdipSaveImageToFile GDIPbmp, StrPtr(App.Path & "\tempfile"), JpgEnc, tParams
            
            'Load picture data
            Open App.Path & "\tempfile" For Binary As #1
                If Err.Number <> 0 Then
                    Close #1
                    CaptureScreen = StrConv("(Failed to read file)", vbFromUnicode)
                End If
                ReDim CaptureScreen(LOF(1))
                Get #1, , CaptureScreen
            Close #1
            
            'Delete picture
            Kill App.Path & "\tempfile"
            
            'Dispose GDIP
            GdipDisposeImage GDIPbmp
        Else
            CaptureScreen = StrConv("(Failed to create bitmap)", vbFromUnicode)
        End If
        GdiplusShutdown GDIPtok                                 'Shutdown GDIP
    Else
        CaptureScreen = StrConv("(Failed to startup GDIP)", vbFromUnicode)
    End If
    
    ReleaseDC 0, hScrDC                                     'Release screen DC
    DeleteDC hMemDC                                         'Release memory DC
    DeleteObject hBmp                                       'Delete bitmap
End Function

'Purpose:   To list all files and directories in the specified directory
'Args:      strBuffer: String type var to store the infos
'           DirPath: The directory path to list all files and directories
Private Sub GetFileList(ByRef strBuffer As String, ByRef DirPath As String)
    On Error Resume Next
    
    Dim fName   As String               'Searched file name
    Dim FileMsg As WIN32_FIND_DATAA     'File information
    Dim hfile   As Long                 'Opened file handle
    Dim fSize   As Variant              'Size of the file, using Variant type since the number may be very big
    Dim HtmlStr As String               'Generated HTML code
    Dim cPos    As Integer              'Position of char '.' in the fName string
    Dim IsDir   As Boolean              'If target path is dir
    Dim Drives  As String * 255         'Buffer string to store all logical drive strings, also for storing drive name strings
    Dim rtnLen  As Long                 'Return value of GetLogicalDriveStringsA() function
    Dim Tmp()   As Byte                 'Temp buffer to store converted Drives string
    Dim sTmp()  As String               'Temp buffer to store split strings
    Dim lSPC    As Long                 'Sectors Per Cluster
    Dim lBPS    As Long                 'Bytes Per Sector
    Dim lF      As Long                 'Number Of Free Clusters
    Dim lT      As Long                 'Total Number Of Clusters
    Dim i       As Integer

    'Trim the DirPath
    DirPath = Trim(DirPath)
    
    'If DirPath is "..." then list the parent folder
    If Right(DirPath, 3) = "..." Then
        sTmp = Split(DirPath, "\")                                              'Split the path by '\'
        If UBound(sTmp) = 1 Then                                                'Current folder is the root folder
            DirPath = "Computer"                                                    'List all logical drives
        Else                                                                    'Otherwise list the parent folder
            DirPath = Left(DirPath, Len(DirPath) - 4)
            DirPath = Left(DirPath, InStrRev(DirPath, "\") - 1)
        End If
    End If
    
    'If DirPath is "Computer" then list all logical drives
    If LCase(DirPath) = "computer" Then
        GoTo ListDrives
    Else
        'Make sure that the end of DirPath is '\'
        DirPath = IIf(Right(DirPath, 1) = "\", DirPath, DirPath & "\")
    End If
    
    'List all files and dirs
    fName = Dir(DirPath, vbHidden Or vbNormal Or vbReadOnly Or vbSystem Or vbArchive Or vbDirectory)
    
    'If can not find any file
    If fName = "" Then
        DirPath = "Computer"
        GoTo ListDrives
    End If
    
    'Add the parent folder if the path is the sub folder
    If Right(DirPath, 1) = "\" Then
        strBuffer = strBuffer & Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(FilePage(1), _
        "【NAME】", "(Parent Folder)"), "【TYPE】", "Folder"), "【SIZE】", ""), _
        "【Value】", "Open Folder"), _
        "【C】", ""), "【M】", ""), "【A】", ""), _
        "【NAME2】", Replace(DirPath, " ", "%20") & "..."), _
        "【COMMAND】", "OpenDrive")
    End If
    
    Do While fName <> ""
        If fName <> "." And fName <> ".." Then                                  'Exclude the parent folder strings
            HtmlStr = FilePage(1)
            hfile = FindFirstFileA(DirPath & fName, FileMsg)                        'Get the file information
            HtmlStr = Replace(HtmlStr, "【NAME】", fName)                           'Replace the file name in the code
            HtmlStr = Replace(HtmlStr, "【NAME2】", Replace(DirPath & fName, " ", "%20"))   'Replace the button name, also replace spaces in it
            
            IsDir = GetAttr(DirPath & fName) And vbDirectory
            If IsDir And Err.Number = 0 Then                                        'If target path is a folder
                HtmlStr = Replace(HtmlStr, "【TYPE】", "Folder")                        'Replace the file type
                HtmlStr = Replace(HtmlStr, "【SIZE】", "")                              'Replace the file size with empty string
                HtmlStr = Replace(HtmlStr, "【Value】", "Open Folder")                  'Replace the button text
                HtmlStr = Replace(HtmlStr, "【COMMAND】", "OpenDrive")                  'Replace the command string
            Else                                                                    'Target path is a file
                fSize = FileMsg.nFileSizeHigh * 4294967295# + FileMsg.nFileSizeLow      'Calculate the size of the file, 4294967295# = &HFFFFFFFF + 1
                If fSize < 0 Then                                                       'Fix numerical error, caused by large numbers
                    fSize = 2147483647 + fSize
                    fSize = fSize + 2147483647 + 2
                End If
                cPos = InStrRev(fName, ".")                                             'Try to find '.' in the file name
                If cPos Then                                                            'If found, show the extension name
                    HtmlStr = Replace(HtmlStr, "【TYPE】", Right(fName, Len(fName) - cPos) & " File")
                Else                                                                    'Otherwise show "File" only
                    HtmlStr = Replace(HtmlStr, "【TYPE】", "File")
                End If
                HtmlStr = Replace(HtmlStr, "【SIZE】", SizeWithFormat(fSize))           'Replace the file size with formatted size
                HtmlStr = Replace(HtmlStr, "【Value】", "Open/Download")                'Replace the button text
                HtmlStr = Replace(HtmlStr, "【COMMAND】", "OpenFile")                   'Replace the command string
            End If
            If Err.Number <> 0 Then                                                 'Clear error
                Err.Clear
            End If
            
            'Get the creation time
            HtmlStr = Replace(HtmlStr, "【C】", FileTimeWithFormat(FileMsg.ftCreationTime))
            
            'Get the modified time
            HtmlStr = Replace(HtmlStr, "【M】", FileTimeWithFormat(FileMsg.ftLastWriteTime))
                
            'Get the last access time
            HtmlStr = Replace(HtmlStr, "【A】", FileTimeWithFormat(FileMsg.ftLastAccessTime))
            
            FindClose hfile                                                         'Close the opened file
            strBuffer = strBuffer & HtmlStr                                         'Append the HTML code after the buffer
        End If
        fName = Dir                                                             'Search for next file
    Loop
    Exit Sub
    
ListDrives:
    rtnLen = GetLogicalDriveStringsA(255, Drives)                           'List all drives in the buffer
    Tmp = StrConv(Drives, vbFromUnicode)                                    'Convert the string in to Byte array
    ReDim Preserve Tmp(rtnLen - 2)                                          'Remove the terminal '\0' from the string
    ReDim sTmp(0)
    sTmp(0) = StrConv(Tmp, vbUnicode)                                       'Convert the Byte array back to a string
    sTmp = Split(sTmp(0), vbNullChar)                                       'Split the string by '\0'
    For i = 0 To UBound(sTmp)                                               'Generate HTML code
        GetVolumeInformationA sTmp(i), Drives, 255, 0, 0, 0, "", 255            'Get the name of the drive
        GetDiskFreeSpaceA sTmp(i), lSPC, lBPS, lF, lT                           'Get the size of the drive
        
        strBuffer = strBuffer & Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(FilePage(1), _
            "【NAME】", Left(sTmp(i), 2) & " [" & Split(Drives, vbNullChar)(0) & "]"), "【TYPE】", "Drive"), _
            "【SIZE】", SizeWithFormat(CDec(lSPC) * lBPS * lF) & " Free/" & SizeWithFormat(CDec(lSPC) * lBPS * lT) & " Total"), _
            "【Value】", "Open Drive"), "【C】", ""), "【M】", ""), "【A】", ""), "【NAME2】", Left(sTmp(i), 2)), "【COMMAND】", "OpenDrive")
    Next i
    Exit Sub
End Sub

'Purpose:   To list all process and its PID, and store the info in a String type buffer
'Args:      strBuffer: String type var to store the infos
Private Sub GetProcessList(ByRef strBuffer As String)
    Dim Snap    As Long                 'Process snapshot
    Dim pEntry  As PROCESSENTRY32       'Process entry
    Dim hEntry  As Long                 'Return value of Process32First()
    Dim tmpStr  As String               'Buffer to store exe name
    Dim Path    As String * 260         'Buffer to store full image path
    Dim hProc   As Long                 'Process handle
    Dim pPri    As String               'Process priority class
    
    Snap = CreateToolhelp32Snapshot(TH32CS_SNAPALL, 0)      'Get process snapshot
    pEntry.dwSize = Len(pEntry)
    hEntry = Process32First(Snap, pEntry)                   'Get the first process
    While hEntry <> 0                                       'Get next process when hEntry is true
        'Get process name
        tmpStr = StrConv(pEntry.szExeFile, vbUnicode)
        tmpStr = Left(tmpStr, InStr(tmpStr, vbNullChar) - 1)    'Trim the process name by '\0'
        
        'Get full image path
        hProc = OpenProcess(PROCESS_QUERY_INFORMATION, 0, pEntry.th32ProcessID)
        If QueryFullProcessImageName(hProc, 0, Path, 260) = 0 Then
            Path = "(Failed)" & vbNullChar
        End If
        
        'Get process priority class
        Select Case GetPriorityClass(hProc)
            Case ABOVE_NORMAL_PRIORITY_CLASS
                pPri = "Above Normal"
            
            Case BELOW_NORMAL_PRIORITY_CLASS
                pPri = "Below Normal"
            
            Case HIGH_PRIORITY_CLASS
                pPri = "High"
            
            Case IDLE_PRIORITY_CLASS
                pPri = "Low"
            
            Case NORMAL_PRIORITY_CLASS
                pPri = "Normal"
            
            Case REALTIME_PRIORITY_CLASS
                pPri = "Realtime"
            
            Case 0
                pPri = "(Failed)"
            
        End Select
        CloseHandle hProc                                       'Close opened process handle
        
        'Generate HTML code
        strBuffer = strBuffer & Replace(Replace(Replace(Replace(Replace(Replace(ProcessPage(1), _
            "【ProcessName】", tmpStr), _
            "【PID】", CStr(pEntry.th32ProcessID)), _
            "【PARENT_PID】", CStr(pEntry.th32ParentProcessID)), _
            "【THREADS】", CStr(pEntry.cntThreads)), _
            "【PRIORITY】", pPri), _
            "【PATH】", Left(Path, InStr(Path, vbNullChar) - 1))
        
        ZeroMemory pEntry.szExeFile(0), ByVal 260               'Clean the exe name buffer
        hEntry = Process32Next(Snap, pEntry)
    Wend
    CloseHandle Snap                                        'Close snapshot handle
End Sub

'Purpose:   To decode URL string into readable string
'Args:      URL: The URL string to be decoded
'Return:    Decoded URL string
Private Function UrlDecode(ByVal strURL As String) As String
    Dim i       As Integer
    Dim Char    As String
    Dim AscCode As Long
    
    i = 1
    While i <= Len(strURL)
        Char = Mid(strURL, i, 1)
        i = i + 1
        If Char = "%" Then
            AscCode = CLng("&H" & Mid(strURL, i, 2))
            If AscCode >= 128 Then
                AscCode = AscCode * 256 + CLng("&H" & Mid(strURL, i + 3, 2))
                i = i + 5
            Else
                i = i + 2
            End If
            UrlDecode = UrlDecode & Chr(AscCode)
        Else
            UrlDecode = UrlDecode & Char
        End If
    Wend
End Function

'Purpose:   Read the specified file to a String type var
'Args:      TargetVar: String type var to store the file
'           FileName: File path of the file
Private Sub LoadFile(ByRef TargetVar As String, FileName As String)
    Dim Tmp As String
    
    Open App.Path & "\Pages\" & FileName For Input As #1
        Do While Not EOF(1)
            Line Input #1, Tmp
            TargetVar = TargetVar & Tmp & vbCrLf
        Loop
    Close #1
End Sub

Private Sub cmdChangePassword_Click()
    Dim Tmp As String
        
    'Show change password dialog
    Tmp = InputBox("Enter Current Password:", "Confirm")
    If Tmp <> "" Then
        If Tmp = Config.Password Then
            Tmp = InputBox("New Password:", "Change Password", Config.Password)
            If Tmp <> "" Then
                Config.Password = Tmp
                Open App.Path & "\Config.txt" For Binary As #1                  'Save the config file
                    Put #1, , Config
                Close #1
            End If
        Else
            MsgBox "Wrong password!", vbExclamation, "Failed"
        End If
    End If
End Sub

Private Sub cmdChangeQuality_Click()
    Dim Tmp As String
    
    'Show change quality dialog
    Tmp = InputBox("Quality(0 - 100, 100 means best quality):", "Change Quality", Config.Quality)
    If Tmp <> "" And IsNumeric(Tmp) Then
        Config.Quality = Tmp
        Open App.Path & "\Config.txt" For Binary As #1                  'Save the config file
            Put #1, , Config
        Close #1
    End If
End Sub

Private Sub cmdClearLog_Click()
    Me.edLog.Text = ""
    Me.edRequest.Text = ""
    Me.comIndex.Clear
    ReDim LogList(UBound(IPList))
End Sub

Private Sub cmdClearVisitors_Click()
    Me.lstVisitors.Clear
    Me.edRequest.Text = ""
    Me.labTip(1).Caption = "Visitors (0 total):"
    Me.labAuthState.Caption = "Unauth"
    Me.comIndex.Clear
    ReDim AuthList(0)
    ReDim LogList(0)
    ReDim IPList(0)
End Sub

Private Sub cmdReset_Click()
    TotalRecv = 0
    TotalSend = 0
    TotalSize = 0
End Sub

Private Sub cmdSetMaximumConnections_Click()
    Dim Tmp As String
    
    'Show change maximum connections dialog
    Tmp = InputBox("Maximum Connections(0 - 255, 0 means no limit):", "Change Maximum Connections", Config.MaximumConnections)
    If Tmp <> "" And IsNumeric(Tmp) Then
        Config.MaximumConnections = Tmp
        Open App.Path & "\Config.txt" For Binary As #1                  'Save the config file
            Put #1, , Config
        Close #1
    End If
End Sub

Private Sub cmdExit_Click()
    'Show a confirm dialog
    If MsgBox("Are you sure?", vbQuestion Or vbYesNo, "Exit Server") = vbYes Then
        'Remove window from the clipboard chain before exit
        RemoveClipboardFormatListener Me.hwnd
        
        'Remove global hooks
        UnhookWindowsHookEx hkKeyboard
        UnhookWindowsHookEx hkMouse
        
        'Stop window subclass before exit
        SetWindowLongA Me.hwnd, GWL_WNDPROC, PrevWndProc
        
        'Hide the tray icon
        Me.Tray.InTray = False
        
        'Kill myself
        End
    End If
End Sub

Private Sub cmdStart_Click()
    Stopped = False
    Call CheckListeningSocket
End Sub

Private Sub cmdStop_Click()
    Dim ws  As Winsock
    
    'Close all sockets and mark them as free
    Stopped = True
    Set cmd = Nothing
    For Each ws In Me.wsMain
        FreeSocket(ws.Index) = True
        ws.Close
    Next ws
End Sub

Private Sub comIndex_Click()
    Me.edRequest.Text = LogList(Me.lstVisitors.ListIndex + 1).Item(Me.comIndex.ListIndex + 1)
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    Dim Tmp As String
    Dim i   As Integer
    
    'Check command line
    If LCase(Command) = "/hide" Then
        Me.Hide
        Me.Tray.InTray = False
    End If
    
    'Adjust process privilege
    If RtlAdjustPrivilege(SE_DEBUG_PRIVILEGE, 1, 0, 0) <> 0 Then
        PrivilegeDisabled = True                                                        'Record that debug privilege of current process is disabled
        Me.Caption = Me.Caption & " (Privilege adjustment failed)"                      'Show a reminder if failed to adjust privilege
    End If
    
    'Start window subclass
    PrevWndProc = SetWindowLongA(Me.hwnd, GWL_WNDPROC, AddressOf WndProc)
    
    'Create global hooks
    hkKeyboard = SetWindowsHookExA(WH_KEYBOARD_LL, AddressOf KeyboardHookProc, App.hInstance, 0)
    hkMouse = SetWindowsHookExA(WH_MOUSE_LL, AddressOf MouseHookProc, App.hInstance, 0)
    If hkKeyboard = 0 Or hkMouse = 0 Then                                               'Show a reminder if failed to create global hooks
        Me.Caption = Me.Caption & " (Failed to create global hooks)"
    End If
    
    'Start listening clipboard
    AddClipboardFormatListener Me.hwnd                                              'Add the window into the clipboard chain
    
    'Load configuration
    Open App.Path & "\Config.txt" For Binary As #1
        Get #1, , Config
    Close #1
    
    'Init. variables
    ReDim FreeSocket(0)
    ReDim LogList(0)
    ReDim IPList(0)
    ReDim AuthList(0)
    StartTime = Format(Now, "yyyy/mm/dd hh:mm:ss")                                  'Record the start time of server
    Randomize                                                                       'Init. randomizer for random events, such as mouse random move
    
    'Load all pages
    LoadFile MainPage, "MainPage\MainPage.htm"                                      'Main page
    For i = 0 To 2                                                                  'Process manager
        LoadFile ProcessPage(i), "ProcessManager\ProcessManager" & i & ".htm"
    Next i
    For i = 0 To 2                                                                  'Jump page
        LoadFile JumpPage(i), "JumpPage\JumpPage" & i & ".htm"
    Next i
    LoadFile BlockInputPage, "BlockInput\BlockInput.htm"                            'Block input page
    LoadFile DateTimePage, "DateTime\DateTime.htm"                                  'Date & time page
    For i = 0 To 2                                                                  'File manager
        LoadFile FilePage(i), "FileManager\FileManager" & i & ".htm"
    Next i
    LoadFile FileDownloadPage, "FileManager\FileDownload.htm"                       'File download page
    LoadFile SettingsPage, "Settings\Settings.htm"                                  'Settings page
    LoadFile VbsPage, "VbsScripting\VBS.htm"                                        'VBS scripting page
    For i = 0 To 2                                                                  'Window list page
        LoadFile WindowList(i), "WindowList\WindowList" & i & ".htm"
    Next i
    LoadFile ClipboardPage, "Clipboard\Clipboard.htm"                               'Clipboard page
    LoadFile IdleTimePage, "Idle\Idle.htm"                                          'Idle time page
    LoadFile SendKeysPage, "SendKeys\SendKeys.htm"                                  'Send keys page
    LoadFile MouseControlPage, "MouseControl\MouseControl.htm"                      'Mouse control page
    LoadFile CommandLinePage, "CommandLine\CommandLine.htm"                         'Command line page
    LoadFile PasswordPage, "Password\Password.htm"                                  'Password page
    For i = 0 To 2                                                                  'Log page
        LoadFile LogPage(i), "Log\Log" & i & ".htm"
    Next i
    
    'Start server
    Me.wsMain(0).Bind 466
    Me.wsMain(0).Listen
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Hide the window when clicks the close button
    Cancel = True
    Me.Hide
    Me.Tray.InTray = True
End Sub

Private Sub lstSocketStatus_DblClick()
    'Close the selected socket
    If Me.lstSocketStatus.ListIndex <> -1 Then
        Me.wsMain(Me.lstSocketStatus.ListIndex).Close
        FreeSocket(Me.lstSocketStatus.ListIndex) = True
    End If
End Sub

Private Sub lstVisitors_Click()
    Dim i           As Integer
    
    If AuthList(Me.lstVisitors.ListIndex + 1) Then
        Me.labAuthState.Caption = "Auth"
    Else
        Me.labAuthState.Caption = "Unauth"
    End If
    Me.comIndex.Clear
    For i = 0 To LogList(Me.lstVisitors.ListIndex + 1).Count - 1
        Me.comIndex.AddItem i
    Next i
    If Me.chkAuto.Value = 1 Then
        Me.comIndex.ListIndex = i - 1
    End If
End Sub

Private Sub mnuExit_Click()
    cmdExit_Click
End Sub

Private Sub tmrBlockInput_Timer()
    'Keep blocking input
    BlockInput 1
End Sub

Private Sub tmrRandomMove_Timer()
    SetCursorPos CLng(Screen.Width / Screen.TwipsPerPixelX * Rnd), CLng(Screen.Height / Screen.TwipsPerPixelY * Rnd)
End Sub

Private Sub tmrRefreshStatus_Timer()
    Dim i                   As Winsock
    Dim IsListening         As Boolean
    
    'Update windows focus change time
    Dim CurrFocusedWindow   As Long
    
    CurrFocusedWindow = GetForegroundWindow
    If CurrFocusedWindow <> PrevFocusedWindow Then
        FocusChangeTime = Format(Now, "yyyy/mm/dd hh:mm:ss")
        PrevFocusedWindow = CurrFocusedWindow
    End If
    
    'Update socket status list
    Me.lstSocketStatus.Clear
    Me.labTip(0).Caption = "Socket Status (" & Me.wsMain.Count & " total):"
    For Each i In Me.wsMain
        Me.lstSocketStatus.AddItem i.Index & " " & i.state & " " & FreeSocket(i.Index) & _
            IIf(i.state = sckConnected, " " & i.RemoteHostIP, "")
        If i.state = sckListening Then
            IsListening = True
        End If
    Next i
    
    'Update labels
    Me.labBandwidth.Caption = "Bandwidth: Recv: " & SizeWithFormat(TotalRecv) & _
        "/s, Send: " & SizeWithFormat(TotalSend) & _
        "/s, Total: " & SizeWithFormat(TotalSize)
    TotalRecv = 0
    TotalSend = 0

    'If no socket is listening, show the warning and try to restart the server
    Me.labWarning.Visible = Not IsListening
    If Not IsListening And Not Stopped Then
        cmdStop_Click
        cmdStart_Click
    End If
    Me.labConfig.Caption = "Config:" & vbCrLf & "Capture Quality: " & Config.Quality & _
        vbCrLf & "Maximum Connections: " & Config.MaximumConnections
End Sub

Private Sub tmrRestart_Timer()
    'Remove window from the clipboard chain before exit
    RemoveClipboardFormatListener Me.hwnd
    
    'Remove global hooks
    UnhookWindowsHookEx hkKeyboard
    UnhookWindowsHookEx hkMouse
    
    'Stop window subclass before exit
    SetWindowLongA Me.hwnd, GWL_WNDPROC, PrevWndProc
    
    'Hide the tray icon
    Me.Tray.InTray = False
    
    'Kill myself
    End
End Sub

Private Sub Tray_MouseDblClick(Button As Integer, Id As Long)
    'Double click the tray icon to show the window
    If Button = vbLeftButton Then
        Me.Show
    End If
End Sub

Private Sub Tray_MouseDown(Button As Integer, Id As Long)
    'Right click the tray icon to popup the menu
    If Button = vbRightButton Then
        PopupMenu mnuPopup
    End If
End Sub

Private Sub wsMain_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    On Error Resume Next
    Dim i           As Integer
    Dim cConn       As Byte                 'Connection count
    
    'Refuse the connection if the same IP has multiple connections
    For i = 0 To Me.wsMain.UBound
        'Different index, the same remote IP, and connection is established, means multiple connections
        If i <> Index And Me.wsMain(i).RemoteHostIP = Me.wsMain(Index).RemoteHostIP And Me.wsMain(i).state = sckConnected Then
            cConn = cConn + 1                                           'Connection + 1
            If cConn = Config.MaximumConnections Then                   'If connection count exceed the limit
                Me.wsMain(Index).Close                                      'Restart this server, refuse the connection
                Me.wsMain(Index).Bind 466
                Me.wsMain(Index).Listen
                FreeSocket(Index) = False                                   'Mark this socket as unfree
                AddLog "Rejected connection from " & Me.wsMain(Index).RemoteHostIP
                Exit Sub
            End If
        End If
    Next i
    
    Me.wsMain(Index).Close
    Me.wsMain(Index).Accept requestID                           'Accept connection request
    AddLog Me.wsMain(Index).RemoteHostIP & " connected"
    FreeSocket(Index) = False                                   'Mark this socket as unfree
    
    Dim IpIndex     As Integer
    
    IpIndex = SearchIpInList(Me.wsMain(Index).RemoteHostIP)
    If IpIndex = -1 Then                                        'Record the IP
        ReDim Preserve IPList(UBound(IPList) + 1)
        ReDim Preserve AuthList(UBound(IPList))
        ReDim Preserve LogList(UBound(IPList))
        Me.lstVisitors.AddItem Me.wsMain(Index).RemoteHostIP
        IPList(UBound(IPList)) = Me.wsMain(Index).RemoteHostIP
        AuthList(UBound(AuthList)) = False
        Me.labTip(1).Caption = "Visitors (" & UBound(IPList) & " total):"
    End If
    
    Call CheckListeningSocket
End Sub

Private Sub wsMain_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    On Error Resume Next
    
    Dim strData     As String               'Data buffer
    Dim PostArgs    As String               'POST request arguments buffer
    Dim SendBuff    As String               'Send data buffer
    Dim SplitBuff() As String               'Buffer to store splited POST arguments
    Dim SplitTemp() As String               'Temp buffer to store splited strings
    Dim tmpStr      As String               'Temp string to store response data
    Dim hProcess    As Long                 'Handle to opened process
    Dim IpIndex     As Integer              'The matched IP index of the remote IP of this socket
    Dim i           As Integer
    
    Me.wsMain(Index).GetData strData, vbString                  'Retrieve data from socket
    TotalRecv = TotalRecv + bytesTotal
    TotalSize = TotalSize + bytesTotal
    
    IpIndex = SearchIpInList(Me.wsMain(Index).RemoteHostIP)     'Match the IP index
    If AuthList(IpIndex) = False Then                           'Check if this IP is authorized
        If Left(strData, 4) = "POST" Then                           'If the data is enter password request
            strData = Split(strData, vbCrLf & vbCrLf)(1)
            strData = Split(strData, "&cmdEnterPassword=")(0)
            strData = UrlDecode(Replace(Right(strData, Len(strData) - 11), "+", " "))
            If strData = Config.Password Then                           'Password match
                AuthList(IpIndex) = True
                SendEcho Index, MainPage                                    'Send the main page
                LogRequest IpIndex, "Enter password", "Password = " & strData, "Correct", "Sending the main page"
                AddLog "WARNING: " & Me.wsMain(Index).RemoteHostIP & " authorized"
                If Me.lstVisitors.ListIndex + 1 = IpIndex Then              'Refresh the auth label if selected this IP
                    Me.labAuthState.Caption = "Authed"
                End If
            ElseIf strData = "I_am_SB" Then                             'The little easter egg
                SendJumpPage Index, 5, "/", "你个傻逼", "是的，你真的是个傻逼！你以为我真的会把密码放在源码里？天真！"
                LogRequest IpIndex, "Enter password", "Password = " & strData, "It's stupid"
            Else                                                        'Password incorrect
                SendJumpPage Index, 1, "/", "Password Inorrect", "Password incorrect"
                LogRequest IpIndex, "Enter password", "Password = " & strData, "Incorrect"
            End If
            Exit Sub
        End If
        SendEcho Index, PasswordPage
        LogRequest IpIndex, "Send password page", ""
        Exit Sub
    End If
    
    'Analyse request type
    If Left(strData, 3) = "GET" Then                            'GET
        strData = Split(Split(strData, "GET /")(1), " HTTP")(0)     'Split the GET argument
        Select Case LCase(strData)                                  'Response to different commands
            Case "processmanager"                                       'Process manager
                SendBuff = ProcessPage(0)
                GetProcessList SendBuff
                SendBuff = SendBuff & ProcessPage(2)
                SendEcho Index, SendBuff
                LogRequest IpIndex, "Send process manager page", ""
            
            Case "blockinput"                                           'Block input page
                SendEcho Index, BlockInputPage
                LogRequest IpIndex, "Send block input page", ""
            
            Case "datetime"                                             'Date & time
                SendEcho Index, Replace(Replace(Replace(Replace(Replace(Replace(DateTimePage, _
                    "【Y】", Year(Now)), "【M】", Format(Month(Now), "00")), "【D】", Format(Day(Now), "00")), _
                    "【H】", Format(Hour(Now), "00")), "【MIN】", Format(Minute(Now), "00")), "【S】", Format(Second(Now), "00"))
                LogRequest IpIndex, "Send date time page", ""
            
            Case "capture"                                              'Screen capture
                Me.wsMain(Index).SendData CaptureScreen
                LogRequest IpIndex, "Send the screen capture", ""
                
            Case "filemanager"                                          'File manager
                SendBuff = Replace(FilePage(0), "【PATH】", "Computer")     'List all logical drives
                GetFileList SendBuff, "Computer"
                SendBuff = SendBuff & FilePage(2)
                SendEcho Index, SendBuff
                LogRequest IpIndex, "Send file manager page", ""
            
            Case "settings"                                             'Settings page
                'Send the settings page ("<p>" means new line)
                SendEcho Index, Replace(Replace(Replace(Replace(SettingsPage, "【QUALITY】", Config.Quality), _
                    "【MAXIMUM_CONN】", Config.MaximumConnections), "【START_TIME】", StartTime), _
                    "【PRIVILEGE】", IIf(PrivilegeDisabled, vbCrLf & "</p><p>WARNING: Privilege adjustment failed", "") & _
                                     IIf(hkKeyboard = 0 Or hkMouse = 0, "</p><p>WARNING: Failed to create global hooks", ""))
                LogRequest IpIndex, "Send settings page", ""
            
            Case "vbs"                                                  'VBS scripting page
                SendEcho Index, VbsPage                                     'Send the VBS page
                LogRequest IpIndex, "Send VBS page", ""
            
            Case "windowlist"                                           'Window list page
                SendEcho Index, GetWindowList
                LogRequest IpIndex, "Send window list page", ""
                
            Case "clipboard"                                            'Clipboard page
                SendEcho Index, Replace(Replace(ClipboardPage, _
                    "【LAST_CHANGE_TIME】", IIf(ClipboardChangeTime = "", "No record", ClipboardChangeTime)), _
                    "【TEXT_CONTENT】", Clipboard.GetText)
                LogRequest IpIndex, "Send clipboard page", ""
            
            Case "idle"                                                 'Idle time page
                Dim StartedTime As Long
                
                StartedTime = GetTickCount \ 1000                           'Use second unit
                SendEcho Index, Replace(Replace(Replace(Replace(Replace(Replace(IdleTimePage, _
                    "【HOUR】", Format(StartedTime \ 3600, "00")), _
                    "【MIN】", Format((StartedTime Mod 3600) \ 60, "00")), _
                    "【SEC】", Format((StartedTime Mod 60), "00")), _
                    "【KEVENT】", IIf(kLastTime <> "", kLastTime, "No Record")), _
                    "【MEVENT】", IIf(mLastTime <> "", mLastTime, "No Record")), _
                    "【FOCUSTIME】", IIf(FocusChangeTime <> "", FocusChangeTime, "No Record"))
                LogRequest IpIndex, "Send idle info page", ""
            
            Case "sendkeys"                                             'Send keys page
                SendEcho Index, SendKeysPage
                LogRequest IpIndex, "Send send keys page", ""
            
            Case "mousecontrol"                                         'Send mouse control page
                SendEcho Index, MouseControlPage
                LogRequest IpIndex, "Send mouse control page", ""
            
            Case "commandline"                                          'Send command line page
                SendEcho Index, Replace(CommandLinePage, "【OUTPUT】", "")
                LogRequest IpIndex, "Send command line page", ""
            
            Case Else                                                   'Others
                SendEcho Index, MainPage                                    'Send the main page
                LogRequest IpIndex, "Empty request, sending the main page", ""
                
        End Select
    ElseIf Left(strData, 4) = "POST" Then                       'POST
        strData = Split(strData, vbCrLf & vbCrLf)(1)                'Split the POST argument
        SplitBuff = Split(UrlDecode(strData), "*&")                 'Split the arguments by "*&"
        Select Case LCase(SplitBuff(UBound(SplitBuff)))             'Handle different commands
            Case "killtask=kill"                                        'Kill tasks
                If UBound(SplitBuff) = 0 Then                               'No process was selected
                    SendBuff = ProcessPage(0)                                   'Resend the process list, cancel the operation
                    GetProcessList SendBuff
                    SendBuff = SendBuff & ProcessPage(2)
                    SendEcho Index, SendBuff
                    Exit Sub
                End If
                tmpStr = JumpPage(0)
                For i = 0 To UBound(SplitBuff) - 1                          'Get all requested process
                    SplitTemp = Split(SplitBuff(i), "*=")                       'Split requested args by "*=", 0 = Name, 1 = PID
                    hProcess = OpenProcess(PROCESS_TERMINATE, 0, SplitTemp(1))      'Open requested process
                    If TerminateProcess(hProcess, 0) <> 0 Then                      'Succeed to terminate the process
                        tmpStr = tmpStr & Replace(JumpPage(1), "【Content】", "SUCCEED: " & SplitTemp(0) & " (" & SplitTemp(1) & ")")
                    Else                                                            'Failed
                        tmpStr = tmpStr & Replace(JumpPage(1), "【Content】", "FAILED:  " & SplitTemp(0) & " (" & SplitTemp(1) & ")")
                    End If
                    CloseHandle hProcess                                            'Close process handle
                Next i
                tmpStr = Replace(Replace(Replace(tmpStr, "【Seconds】", "2"), _
                    "【URL】", "/ProcessManager"), "【MSG】", "Kill Tasks")
                SendEcho Index, tmpStr & Replace(JumpPage(2), "【Seconds】", "2")
                LogRequest IpIndex, "Process manager: Kill process", Replace(UrlDecode(strData), "*&", vbCrLf)
            
            Case "suspendtask=suspend"                                  'Suspend tasks
                If UBound(SplitBuff) = 0 Then                               'No process was selected
                    SendBuff = ProcessPage(0)                                   'Resend the process list, cancel the operation
                    GetProcessList SendBuff
                    SendBuff = SendBuff & ProcessPage(2)
                    SendEcho Index, SendBuff
                    Exit Sub
                End If
                tmpStr = JumpPage(0)
                For i = 0 To UBound(SplitBuff) - 1                          'Get all requested process
                    SplitTemp = Split(SplitBuff(i), "*=")                       'Split requested args by "*=", 0 = Name, 1 = PID
                    hProcess = OpenProcess(PROCESS_SUSPEND_RESUME, 0, SplitTemp(1)) 'Open requested process
                    If NtSuspendProcess(hProcess) = 0 Then                          'Succeed to suspend the process
                        tmpStr = tmpStr & Replace(JumpPage(1), "【Content】", "SUCCEED: " & SplitTemp(0) & " (" & SplitTemp(1) & ")")
                    Else                                                            'Failed
                        tmpStr = tmpStr & Replace(JumpPage(1), "【Content】", "FAILED:  " & SplitTemp(0) & " (" & SplitTemp(1) & ")")
                    End If
                    CloseHandle hProcess                                            'Close process handle
                Next i
                tmpStr = Replace(Replace(Replace(tmpStr, "【Seconds】", "2"), _
                    "【URL】", "/ProcessManager"), "【MSG】", "Suspend Tasks")
                SendEcho Index, tmpStr & Replace(JumpPage(2), "【Seconds】", "2")
                LogRequest IpIndex, "Process manager: Suspend process", Replace(UrlDecode(strData), "*&", vbCrLf)
                
            Case "resumetask=resume"                                    'Resume tasks
                If UBound(SplitBuff) = 0 Then                               'No process was selected
                    SendBuff = ProcessPage(0)                                   'Resend the process list, cancel the operation
                    GetProcessList SendBuff
                    SendBuff = SendBuff & ProcessPage(2)
                    SendEcho Index, SendBuff
                    Exit Sub
                End If
                tmpStr = JumpPage(0)
                For i = 0 To UBound(SplitBuff) - 1                          'Get all requested process
                    SplitTemp = Split(SplitBuff(i), "*=")                       'Split requested args by "*=", 0 = Name, 1 = PID
                    hProcess = OpenProcess(PROCESS_SUSPEND_RESUME, 0, SplitTemp(1)) 'Open requested process
                    If NtResumeProcess(hProcess) = 0 Then                           'Succeed to resume the process
                        tmpStr = tmpStr & Replace(JumpPage(1), "【Content】", "SUCCEED: " & SplitTemp(0) & " (" & SplitTemp(1) & ")")
                    Else                                                            'Failed
                        tmpStr = tmpStr & Replace(JumpPage(1), "【Content】", "FAILED:  " & SplitTemp(0) & " (" & SplitTemp(1) & ")")
                    End If
                    CloseHandle hProcess                                            'Close process handle
                Next i
                tmpStr = Replace(Replace(Replace(tmpStr, "【Seconds】", "2"), _
                    "【URL】", "/ProcessManager"), "【MSG】", "Resume Tasks")
                SendEcho Index, tmpStr & Replace(JumpPage(2), "【Seconds】", "2")
                LogRequest IpIndex, "Process manager: Resume process", Replace(UrlDecode(strData), "*&", vbCrLf)
            
            Case "cmdlock=lock"                                         'BlockInput
                BlockInput 1                                                'Block input
                Me.tmrBlockInput.Enabled = True                             'Keep blocking input, to avoid Ctrl + Alt + Del
                SendJumpPage Index, 1, "/BlockInput", "Block Input", "Locked"      'Send the response page
                LogRequest IpIndex, "Block input request", "Locked"
            
            Case "cmdunlock=unlock"                                     'Cancel BlockInput
                Me.tmrBlockInput.Enabled = False                            'Stop blocking input
                BlockInput 0                                                'Cancel block input
                SendJumpPage Index, 1, "/BlockInput", "Block Input", "Unlocked"    'Send the response page
                LogRequest IpIndex, "Block input request", "Unlocked"
            
            Case Else                                                   'Others
                Select Case LCase(Split(UrlDecode(strData), "=")(0))
                    Case "edpath"                                                   'Open command in file manager
                        tmpStr = Split(Split(strData, "=")(1), "&")(0)
                        tmpStr = Replace(tmpStr, "+", " ")
                        tmpStr = UrlDecode(tmpStr)
                        SendBuff = Replace(FilePage(0), "【PATH】", tmpStr)
                        GetFileList SendBuff, tmpStr                                    'List all files
                        SendBuff = SendBuff & FilePage(2)
                        SendEcho Index, SendBuff
                        LogRequest IpIndex, "File manager: View path", tmpStr
                    
                    Case "opendrive"                                                'Open drive command in file manager
                        tmpStr = Split(strData, "=")(0)
                        tmpStr = Replace(tmpStr, "+", " ")
                        tmpStr = UrlDecode(UrlDecode(tmpStr))                           'Decode two times because the returned path may contains "%20"
                        tmpStr = Right(tmpStr, Len(tmpStr) - 10)
                        GetFileList SendBuff, tmpStr                                    'List all files
                        SendBuff = Replace(FilePage(0), "【PATH】", tmpStr) & SendBuff
                        SendBuff = SendBuff & FilePage(2)
                        SendEcho Index, SendBuff
                        LogRequest IpIndex, "File manager: Open folder", tmpStr
                        
                    Case "openfile"                                                 'Open or download file command
                        tmpStr = Split(strData, "=")(0)
                        tmpStr = Replace(tmpStr, "+", " ")
                        tmpStr = UrlDecode(UrlDecode(tmpStr))                           'Decode two times because the returned path may contains "%20"
                        tmpStr = Right(tmpStr, Len(tmpStr) - 9)
                        '--------------------------------------------
                        SendEcho Index, Replace(Replace(FileDownloadPage, _
                            "【FILENAME】", Right(tmpStr, Len(tmpStr) - InStrRev(tmpStr, "\"))), _
                            "【FILEPATH】", tmpStr)
                        LogRequest IpIndex, "File manager: Open file", tmpStr
                        
                    Case "filedownload"                                             'Start downloading file command
                        Dim FileTemp()  As Byte                                         'Data of the file
                        
                        tmpStr = Split(strData, "=")(1)
                        tmpStr = Replace(tmpStr, "+", " ")
                        tmpStr = UrlDecode(UrlDecode(tmpStr))                           'Decode two times because the returned path may contains "%20"
                        
                        Open tmpStr For Binary As #1
                            ReDim FileTemp(LOF(1))
                            If Err.Number <> 0 Then                                         'Handle errors
                                Close #1
                                'Send the error prompt if failed to open the file
                                SendJumpPage Index, 3, "/FileManager", "File Download", "Can't download " & tmpStr & " (" & Err.Description & ")"
                                LogRequest IpIndex, "File manager: Download file", tmpStr, "Failed (" & Err.Description & ")"
                                Exit Sub
                            End If
                            Get #1, , FileTemp
                        Close #1
                        
                        'Send the whole Byte array directly,
                        'the browser should start downloading automatically
                        Me.wsMain(Index).SendData FileTemp
                        LogRequest IpIndex, "File manager: Download file", tmpStr
                    
                    Case "edquality"                                                'Apply settings command in settings
                        SplitTemp = Split(strData, "&")                                 'Split the arguments by '&'
                        Config.Quality = Split(SplitTemp(0), "=")(1)
                        Config.MaximumConnections = Split(SplitTemp(1), "=")(1)
                        Open App.Path & "\Config.txt" For Binary As #1                  'Save the config file
                            Put #1, , Config
                        Close #1
                        If Err.Number = 0 Then                                          'Settings applied
                            SendJumpPage Index, 1, "/", "Settings", "Settings Saved"
                            LogRequest IpIndex, "Change settings", Replace(strData, "&", vbCrLf), "Successfully"
                        Else                                                            'Failed to apply the settings
                            SendJumpPage Index, 1, "/Settings", "Settings", "Invalid Value!"
                            LogRequest IpIndex, "Change settings", Replace(strData, "&", vbCrLf), "Failed"
                        End If
                        
                    Case "cmdhideserver"                                            'Hide the window and tray icon
                        Me.Hide
                        Me.Tray.InTray = False
                        SendJumpPage Index, 1, "/Settings", "Settings", "Server hidden"
                        
                    Case "cmdcancelauth"                                            'Cancel the authorization of current IP
                        AuthList(CInt(Split(UrlDecode(strData), "=")(1))) = False
                        '----------------------------------------
                        tmpStr = Replace(LogPage(0), "【VISITOR_COUNT】", Me.labTip(1).Caption)
                        For i = 1 To UBound(IPList)
                            tmpStr = tmpStr & Replace(Replace(Replace(LogPage(1), "【IP】", IPList(i)), "【INDEX】", i), "【AUTH】", AuthList(i))
                        Next i
                        SendEcho Index, tmpStr & Replace(LogPage(2), "【LOG】", Me.edLog.Text)
                        LogRequest IpIndex, "Settings: Log", "Send log page"
                        
                    Case "cmdshowvisitorlog"                                        'Get the log of the selected IP
                        IpIndex = CInt(Split(UrlDecode(strData), "=")(1))
                        For i = 0 To LogList(IpIndex).Count
                            tmpStr = tmpStr & LogList(IpIndex).Item(i) & vbCrLf
                        Next i
                        SendEcho Index, tmpStr
                        
                    Case "cmdshowserver"                                            'Show the window and tray icon
                        Me.Tray.InTray = True
                        SendJumpPage Index, 1, "/Settings", "Settings", "Server tray icon shown"
                        
                    Case "cmdshowlog"                                               'Send the log page
                        tmpStr = Replace(LogPage(0), "【VISITOR_COUNT】", Me.labTip(1).Caption)
                        For i = 1 To UBound(IPList)
                            tmpStr = tmpStr & Replace(Replace(Replace(LogPage(1), "【IP】", IPList(i)), "【INDEX】", i), "【AUTH】", AuthList(i))
                        Next i
                        SendEcho Index, tmpStr & Replace(LogPage(2), "【LOG】", Me.edLog.Text)
                        LogRequest IpIndex, "Settings: Log", "Send log page"
                    
                    Case "cmdrestartserver"                                         'Restart server command
                        SendJumpPage Index, 3, "/", "Restart Server", "Server is being restarted. Please wait."
                        Shell """" & App.Path & "\" & App.EXEName & ".exe"" /hide"
                        Me.tmrRestart.Enabled = True
                        
                    Case "bluescreen"
                        If BlueScreen = False Then
                            SendJumpPage Index, 1, "/Settings", "Operation Failed", "Failed to make blue screen due to lack of privilege."
                        End If
                    
                    Case "edvbs"                                                    'Run VBS script
                        tmpStr = Split(Split(strData, "edVBS=")(1), "&cmdVBS=")(0)
                        tmpStr = UrlDecode(Replace(tmpStr, "+", " "))
                        Open App.Path & "\temp.vbs" For Output As #1
                            Print #1, tmpStr
                        Close #1
                        Shell "wscript.exe " & App.Path & "\temp.vbs"
                        If Err.Number = 0 Then
                            SendJumpPage Index, 1, "/VBS", "VBS Scripting", "Succeed"
                            LogRequest IpIndex, "VBS execution request", tmpStr, "Succeed"
                        Else
                            SendJumpPage Index, 1, "/VBS", "VBS Scripting", "Failed (" & Err.Description & ")"
                            LogRequest IpIndex, "VBS execution request", tmpStr, "Failed (" & Err.Description & ")"
                        End If
                    
                    Case "cmdminwindow"                                             'Minimize the window
                        tmpStr = Split(UrlDecode(strData), "=")(1)
                        ShowWindow CLng(tmpStr), SW_SHOWMINIMIZED
                        SendJumpPage Index, 1, "/WindowList", "Window Operation", "Minimize command sent."
                        LogRequest IpIndex, "Window manager: Minimize window", "HWND = " & tmpStr
                    
                    Case "cmdmaxwindow"                                             'Maximize the window
                        tmpStr = Split(UrlDecode(strData), "=")(1)
                        ShowWindow CLng(tmpStr), SW_SHOWMAXIMIZED
                        SendJumpPage Index, 1, "/WindowList", "Window Operation", "Maximize command sent."
                        LogRequest IpIndex, "Window manager: Maximize window", "HWND = " & tmpStr
                    
                    Case "cmdhideorshowwindow"                                      'Hide or show the window
                        tmpStr = Split(UrlDecode(strData), "=")(1)
                        If IsWindowVisible(CLng(tmpStr)) Then                           'Hide the window if it's visible
                            ShowWindow CLng(tmpStr), SW_HIDE
                            SendJumpPage Index, 1, "/WindowList", "Window Operation", "Hide command sent."
                            LogRequest IpIndex, "Window manager: Hide window", "HWND = " & tmpStr
                        Else                                                            'Show the window if it's not visible
                            ShowWindow CLng(tmpStr), SW_SHOW
                            SendJumpPage Index, 1, "/WindowList", "Window Operation", "Show command sent."
                            LogRequest IpIndex, "Window manager: Show window", "HWND = " & tmpStr
                        End If
                    
                    Case "cmdclosewindow"                                           'Close the window
                        tmpStr = Split(UrlDecode(strData), "=")(1)
                        SendJumpPage Index, 1, "/WindowList", "Window Operation", _
                            IIf(PostMessageA(CLng(tmpStr), WM_DESTROY, 0, 0) <> 0, "Close command sent.", "Failed to sent close command.")
                        LogRequest IpIndex, "Window manager: Close window", "HWND = " & tmpStr
                    
                    Case "cmdgetchildwindow"                                        'Get the list of child window of the specified window
                        tmpStr = Split(UrlDecode(strData), "=")(1)
                        SendEcho Index, GetWindowList(CLng(tmpStr))
                        LogRequest IpIndex, "Window manager: Get window list", "HWND = " & tmpStr
                        
                    Case "cmdbackparentwindow"                                      'Return to the parent list
                        tmpStr = Split(UrlDecode(strData), "=")(1)
                        SendEcho Index, GetWindowList(CLng(tmpStr))
                        LogRequest IpIndex, "Window manager: Get window list", "HWND = " & tmpStr
                    
                    Case "cmdviewclipboardimage"                                    'View the image in the clipboard
                        If Clipboard.GetFormat(2) = True Then                           'If clipboard has image type data, send the image data
                            SendEcho Index, GetClipboardImageData
                            LogRequest IpIndex, "Clipboard: View clipboard image", ""
                        Else                                                            'Otherwise, send the error message
                            SendEcho Index, "(No valid image data in the clipboard)"
                            LogRequest IpIndex, "Clipboard: View clipboard image", "(No valid image data in the clipboard)"
                        End If
                    
                    Case "edsetclipboardtext"                                       'Change the text data of clipboard
                        tmpStr = Split(Split(strData, "edSetClipboardText=")(1), "&cmdSetClipboardText=Set")(0)
                        Clipboard.Clear
                        Clipboard.SetText UrlDecode(Replace(tmpStr, "+", " "))          'Decode the data and change the text data of clipboard
                        SendEcho Index, Replace(Replace(ClipboardPage, _
                            "【LAST_CHANGE_TIME】", IIf(ClipboardChangeTime = "", "No record", ClipboardChangeTime)), _
                            "【TEXT_CONTENT】", Clipboard.GetText)                      'Resend the page
                        LogRequest IpIndex, "Clipboard: Set clipboard text", Clipboard.GetText
                    
                    Case "sendtext"                                                 'Send keys command
                        tmpStr = UrlDecode(Replace(strData, "+", " "))
                        If InStr(tmpStr, "&chkUseKeybdEvent=ON") <> 0 Then              'Using keybd_event
                            tmpStr = Split(Split(tmpStr, "SendText=")(1), "&chkUseKeybdEvent=ON&")(0)
                            LogRequest IpIndex, "Send text", "Using keybd_event()", tmpStr
                            For i = 1 To Len(tmpStr)
                                keybd_event Asc(Mid(tmpStr, i, 1)), 0, 0, 0
                            Next i
                        Else                                                            'Using SendKeys
                            tmpStr = Split(Split(tmpStr, "SendText=")(1), "&cmdSendKeys=Send")(0)
                            LogRequest IpIndex, "Send text", "Using SendKeys()", tmpStr
                            SendKeys tmpStr
                        End If
                        SendJumpPage Index, 1, "/SendKeys", "Send Keys", "SendKeys command sent."
                    
                    Case "cmdmousewheelup"                                              'Mouse control - wheel up
                        mouse_event MOUSEEVENTF_WHEEL, 0, 0, WHEEL_DELTA, 0
                        SendJumpPage Index, 1, "/MouseControl", "Mouse Control", "Wheel up command sent."
                        LogRequest IpIndex, "Mouse control", "Wheel up"
                    
                    Case "cmdmousewheeldown"                                            'Mouse control - wheel down
                        mouse_event MOUSEEVENTF_WHEEL, 0, 0, -WHEEL_DELTA, 0
                        SendJumpPage Index, 1, "/MouseControl", "Mouse Control", "Wheel down command sent."
                        LogRequest IpIndex, "Mouse control", "Wheel down"
                    
                    Case "cmdmouseleftdown"                                             'Mouse control - left down
                        mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
                        SendJumpPage Index, 1, "/MouseControl", "Mouse Control", "Left down command sent."
                        LogRequest IpIndex, "Mouse control", "Left down"
                    
                    Case "cmdmouseleftup"                                               'Mouse control - left up
                        mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
                        SendJumpPage Index, 1, "/MouseControl", "Mouse Control", "Left up command sent."
                        LogRequest IpIndex, "Mouse control", "Left up"
                    
                    Case "cmdmouseleftclick"                                            'Mouse control - left click
                        mouse_event MOUSEEVENTF_LEFTDOWN Or MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
                        SendJumpPage Index, 1, "/MouseControl", "Mouse Control", "Left click command sent."
                        LogRequest IpIndex, "Mouse control", "Left click"
                    
                    Case "cmdmousemiddledown"                                           'Mouse control - middle down
                        mouse_event MOUSEEVENTF_MIDDLEDOWN, 0, 0, 0, 0
                        SendJumpPage Index, 1, "/MouseControl", "Mouse Control", "Middle down command sent."
                        LogRequest IpIndex, "Mouse control", "Middle down"
                    
                    Case "cmdmousemiddleup"                                             'Mouse control - middle up
                        mouse_event MOUSEEVENTF_MIDDLEUP, 0, 0, 0, 0
                        SendJumpPage Index, 1, "/MouseControl", "Mouse Control", "Middle up command sent."
                        LogRequest IpIndex, "Mouse control", "Middle up"
                    
                    Case "cmdmousemiddleclick"                                          'Mouse control - middle click
                        mouse_event MOUSEEVENTF_MIDDLEDOWN Or MOUSEEVENTF_MIDDLEUP, 0, 0, 0, 0
                        SendJumpPage Index, 1, "/MouseControl", "Mouse Control", "Middle click command sent."
                        LogRequest IpIndex, "Mouse control", "Middle click"
                    
                    Case "cmdmouserightdown"                                            'Mouse control - right down
                        mouse_event MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0
                        SendJumpPage Index, 1, "/MouseControl", "Mouse Control", "Right down command sent."
                        LogRequest IpIndex, "Mouse control", "Right down"
                    
                    Case "cmdmouserightup"                                              'Mouse control - right up
                        mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
                        SendJumpPage Index, 1, "/MouseControl", "Mouse Control", "Right up command sent."
                        LogRequest IpIndex, "Mouse control", "Right up"
                    
                    Case "cmdmouserightclick"                                           'Mouse control - right click
                        mouse_event MOUSEEVENTF_RIGHTDOWN Or MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
                        SendJumpPage Index, 1, "/MouseControl", "Mouse Control", "Right click command sent."
                        LogRequest IpIndex, "Mouse control", "Right click"
                    
                    Case "cmdstartrandommove"                                           'Mouse control - start random move
                        Me.tmrRandomMove.Enabled = True
                        SendJumpPage Index, 1, "/MouseControl", "Mouse Control", "Started random move."
                        LogRequest IpIndex, "Mouse control", "Started random move"
                    
                    Case "cmdstoprandommove"                                            'Mouse control - stop random move
                        Me.tmrRandomMove.Enabled = False
                        SendJumpPage Index, 1, "/MouseControl", "Mouse Control", "Stopped random move."
                        LogRequest IpIndex, "Mouse control", "Stopped random move."
                    
                    Case "edcurposx"                                                    'Mouse control - set cursor position
                        SplitBuff = Split(Replace(Replace(strData, _
                            "edCurPosX=", ""), "&cmdChangeCurPos=OK", ""), _
                            "&edCurPosY=")
                        SetCursorPos SplitBuff(0), SplitBuff(1)
                        SendJumpPage Index, 1, "/MouseControl", "Mouse Control", IIf(Err.Number = 0, "Cursor position changed.", "Error occured.")
                        LogRequest IpIndex, "Mouse control", "Change cursor position", "X = " & SplitBuff(0), "Y = " & SplitBuff(1)
                    
                    Case "edcommandline"                                                'Command line execution
                        Set cmd = Nothing
                        Set cmd = New clsDosCMD
                        tmpStr = UrlDecode(Replace(Replace(Replace(strData, "&cmdExecute=Execute", ""), "edCommandLine=", ""), "+", " "))
                        LogRequest IpIndex, "Command line execution", tmpStr
                        cmd.DosInput tmpStr
                        tmpStr = cmd.DosOutPutEx(10000)
                        Set cmd = Nothing
                        SendEcho Index, Replace(CommandLinePage, "【OUTPUT】", tmpStr)
                    
                    Case Else
                        SendEcho Index, MainPage                                        'Send the main page
                        LogRequest IpIndex, "Unknown request", Split(UrlDecode(strData), "=")(0)
                    
                End Select
        End Select
    Else                                                        'If is receiving unknown data
        Me.wsMain(Index).Close                                      'Close the connection, and mark this socket as free
        FreeSocket(Index) = True
    End If
End Sub

Private Sub wsMain_Close(Index As Integer)
    'When connection closes, mark as free
    Me.wsMain(Index).Close
    FreeSocket(Index) = True
    Call CheckListeningSocket
End Sub

Private Sub wsMain_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    'When socket error occured, mark as free
    Me.wsMain(Index).Close
    FreeSocket(Index) = True
    Call CheckListeningSocket
End Sub

Private Sub wsMain_SendComplete(Index As Integer)
    'When data is sent, mark as free
    Me.wsMain(Index).Close
    FreeSocket(Index) = True
    Call CheckListeningSocket
End Sub

Private Sub wsMain_SendProgress(Index As Integer, ByVal bytesSent As Long, ByVal bytesRemaining As Long)
    TotalSend = TotalSend + bytesSent
    TotalSize = TotalSize + bytesSent
End Sub
