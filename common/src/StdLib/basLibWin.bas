Attribute VB_Name = "basLibWin"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                                                           '
' LibWin                                                    '
'                                                           '
' (c) 2017-2024 Michael Brodsky, mbrodskiis@gmail.com       '
' (v) 20241107                                              '
'                                                           '
' All rights reserved. Unauthorized use prohibited.         '
'                                                           '
' DESCRIPTION                                               '
'                                                           '
' This module defines VBA versions of types and objects     '
' from the Windows API, and several useful functions that   '
' provide system-level access to and control of processes   '
' and windows.                                              '
'                                                           '
' DEPENDENCIES                                              '
'                                                           '
' (None)                                                    '
'                                                           '
' NOTES                                                     '
'                                                           '
' This version of the library has only been tested with     '
' MS ACCESS 365 (64-bit) implementations.                   '
'                                                           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Compare Database
Option Explicit

' Aggregate type containing message information from a thread's message queue.
' Note: The Windows API declares this type as MSG, however it is renamed here
' to WINMSG because the MSG token is too short and often conflicts with existing code.
Public Type WINMSG
    hwnd As Long                ' A handle to the window whose window procedure receives the message,
    message As Long             ' The message identifier,
    wParam As Long              ' Additional information about the message,
    lParam As Long              ' Additional information about the message,
    time As Long                ' The time at which the message was posted.
End Type

' Aggregate type containing information used by ShellExecuteEx.
Public Type SHELLEXECUTEINFO
        cbSize As Long          ' The size of this structure in bytes,
        fMask As Long           ' Indicates the content and validity of the other structure members (See 'SEE' constants, below),
        hwnd As LongPtr         ' A handle to the owner window, if any,
        lpVerb As String        ' A string that specifies the action to be performed,
        lpFile As String        ' String specifying the name of the file or object on which ShellExecuteEx will perform the action,
        lpParameters As String  ' String containing the application parameters,
        lpDirectory As String   ' String specifying the name of the working directory,
        nShow As Long           ' Specifies how an application is to be shown when it is opened,
        hInstApp As LongPtr     ' Contains the result code of the shelled application (See SE_ERR constants, below),
        lpIDList As LongPtr     ' ITEMIDLIST structure that uniquely identifies the file to execute,
        lpClass As String       ' Specifies either a ProgId, URI protocol scheme, or file extension,
        hkeyClass As LongPtr    ' A handle to the registry key for the file type,
        dwHotKey As Long        ' A keyboard shortcut to associate with the application,
        hIcon As LongPtr        ' A handle to the icon for the file type,
        hprocess As LongPtr     ' A handle to the newly started application.
End Type

'''''''''''''''''''''''''''''''''''''
' Windows API Function Declarations '
'''''''''''''''''''''''''''''''''''''

Public Declare PtrSafe Function CloseHandle Lib "kernel32" ( _
    ByVal hObject As LongPtr _
) As Long

Public Declare PtrSafe Function CoInitialize Lib "ole32" ( _
    pRef As Long _
) As Long

Public Declare PtrSafe Function CoInitializeEx Lib "ole32" ( _
    pRef As Long, _
    ByVal dwCoInit As Long _
) As Long

Public Declare PtrSafe Function CoUninitialize Lib "ole32" () As Long

Public Declare PtrSafe Function DispatchMessage Lib "user32" Alias "DispatchMessageA" ( _
    ByRef lpMsg As WINMSG _
) As Boolean

Public Declare PtrSafe Function EnumChildWindows Lib "user32" ( _
    ByVal hwnd As LongPtr, _
    ByVal lpEnumFunc As LongPtr, _
    ByVal lParam As LongPtr _
) As LongPtr

Public Declare PtrSafe Function EnumWindows Lib "user32" ( _
    ByVal lpEnumFunc As LongPtr, _
    ByVal lParam As LongPtr _
) As Boolean

Public Declare PtrSafe Function CreateEvent Lib "kernel32" Alias "CreateEventA" ( _
    ByVal lpEventAttributes As Long, _
    ByVal bManualReset As Long, _
    ByVal bInitialState As Long, _
    ByVal lpName As String _
) As LongPtr

Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" ( _
    ByVal pCaller As Long, _
    ByVal szURL As String, _
    ByVal szFileName As String, _
    ByVal dwReserved As Long, _
    ByVal lpfnCB As Long _
) As Long

Public Declare PtrSafe Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" ( _
    ByVal lpFile As String, _
    ByVal lpDirectory As String, _
    ByVal lpResult As String _
) As LongPtr

Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
    ByVal lpClassName As String, _
    ByVal lpWindowName As String _
) As LongPtr

Public Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr

Public Declare PtrSafe Function GetAncestor Lib "user32" ( _
    ByVal hwnd As LongPtr, _
    ByVal gaFlags As Integer _
) As LongPtr

Public Declare PtrSafe Function GetDesktopWindow Lib "user32" () As LongPtr

Public Declare PtrSafe Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" ( _
    ByVal lpName As String, _
    ByVal lpBuffer As String, _
    ByVal NSize As Long _
) As Long

Public Declare PtrSafe Function GetExitCodeProcess Lib "kernel32" ( _
    ByVal hObject As LongPtr, _
    lpExitCode As Long _
) As Boolean

Public Declare PtrSafe Function GetLastError Lib "kernel32" () As Long

Public Declare PtrSafe Function GetParent Lib "user32" ( _
    ByVal hwnd As LongPtr _
) As LongPtr

Public Declare PtrSafe Function GetProcessId Lib "kernel32" ( _
    ByVal hprocess As LongPtr _
) As Long

Public Declare PtrSafe Function GetWindow Lib "user32" ( _
    ByVal hwnd As LongPtr, _
    ByVal wCmd As Long _
) As LongPtr

Public Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" ( _
    ByVal hwnd As LongPtr, _
    ByVal nIndex As Long _
) As LongPtr

Public Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" ( _
    ByVal hwnd As LongPtr, _
    ByVal lpString As String, _
    ByVal cch As Long _
) As Long

Public Declare PtrSafe Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" ( _
    ByVal hwnd As LongPtr _
) As Long

Public Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" ( _
    ByVal hwnd As LongPtr, _
    lpdwProcessId As Long _
) As Long

Private Declare PtrSafe Function InternetGetConnectedState Lib "wininet.dll" ( _
    ByRef dwflags As Long, _
    ByVal dwReserved As Long _
) As Long

Public Declare PtrSafe Function IsTopWindow Lib "user32" ( _
    ByVal hwnd As LongPtr _
) As Boolean

Public Declare PtrSafe Function IsWindow Lib "user32" ( _
    ByVal hwnd As LongPtr _
) As Boolean

Public Declare PtrSafe Function IsWindowVisible Lib "user32" ( _
    ByVal hwnd As LongPtr _
) As Boolean
  
Public Declare PtrSafe Function OpenProcess Lib "kernel32" ( _
    ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Boolean, _
    ByVal dwProcessId As Long _
) As LongPtr

Public Declare PtrSafe Function PeekMessage Lib "user32" Alias "PeekMessageA" ( _
    ByRef lpMsg As WINMSG, _
    ByVal hwnd As LongPtr, _
    ByVal wMsgFilterMin As Long, _
    ByVal wMsgFilterMax As Long, _
    ByVal wRemoveMsg As Long _
) As Boolean

Public Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hwnd As LongPtr, _
    ByVal wMsg As Long, _
    ByVal wParam As LongPtr, _
    lParam As Any _
) As LongPtr

Public Declare PtrSafe Function SetForegroundWindow Lib "user32" ( _
    ByVal hwnd As LongPtr _
) As Long

Public Declare PtrSafe Function SetEnvironmentVariable Lib "kernel32" Alias "SetEnvironmentVariableA" ( _
    ByVal lpName As String, _
    ByVal lpValue As String _
) As Long
     
Public Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hwnd As LongPtr, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long _
) As LongPtr

Public Declare PtrSafe Function ShellExecuteEx Lib "shell32.dll" ( _
    info As SHELLEXECUTEINFO _
) As Boolean

Public Declare PtrSafe Sub Sleep Lib "kernel32" ( _
    ByVal dwMilliseconds As Long _
)
Public Declare PtrSafe Function ShowWindow Lib "user32" ( _
    ByVal hwnd As LongPtr, _
    ByVal shwCmd As Long _
) As Boolean

Public Declare PtrSafe Function TerminateProcess Lib "kernel32" ( _
    ByVal hprocess As LongPtr, _
    ByVal uExitCode As Long _
) As Long

Public Declare PtrSafe Function TranslateMessage Lib "user32" ( _
    ByRef lpMsg As WINMSG _
) As Boolean

Public Declare PtrSafe Function WaitForInputIdle Lib "user32" ( _
    ByVal hprocess As LongPtr, _
    ByVal dwMilliseconds As Long _
) As Long


Public Declare PtrSafe Function MsgWaitForMultipleObjectsEx Lib "user32" ( _
    ByVal nCount As Long, _
    pHandles As LongPtr, _
    ByVal dwMilliseconds As Long, _
    ByVal dwWakeMask As Long, _
    ByVal dwflags As Long _
) As Long

'''''''''''''''''''''''''
' Windows API Constants '
'''''''''''''''''''''''''

Public Const GA_PARENT = 1
Public Const GA_ROOT = 2
Public Const GA_ROOTOWNER = 3
    
Public Const GW_CHILD = 5
Public Const GW_ENABLEDPOPUP = 6
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GW_OWNER = 4

Public Const GWL_EXSTYLE = -20
Public Const GWLP_HINSTANCE = -6
Public Const GWLP_HWNDPARENT = -8
Public Const GWLP_ID = -12
Public Const GWL_STYLE = -16
Public Const GWLP_USERDATA = -21
Public Const GWLP_WNDPROC = -4

Public Const HOTKEYF_ALT = &H4
Public Const HOTKEYF_CONTROL = &H2
Public Const HOTKEYF_EXT = &H8
Public Const HOTKEYF_SHIFT = &H1

Public Const MWMO_ALERTABLE = &H2
Public Const MWMO_INPUTAVAILABLE = &H4
Public Const MWMO_WAITALL = &H1

Public Const PM_REMOVE = &H1

Public Const PROCESS_CREATE_PROCESS = &H80
Public Const PROCESS_CREATE_THREAD = &H2
Public Const PROCESS_DUP_HANDLE = &H40
Public Const PROCESS_QUERY_INFORMATION = &H400
Public Const PROCESS_QUERY_LIMITED_INFORMATION = &H1000
Public Const PROCESS_SET_INFORMATION = &H200
Public Const PROCESS_SET_QUOTA = &H100
Public Const PROCESS_SUSPEND_RESUME = &H800
Public Const PROCESS_TERMINATE = &H1
Public Const PROCESS_VM_OPERATION = &H8
Public Const PROCESS_VM_READ = &H10
Public Const PROCESS_VM_WRITE = &H20
Public Const PROCESS_SYNCHRONIZE = &H100000
Public Const PROCESS_DELETE = &H10000
Public Const PROCESS_READ_CONTROL = &H20000
Public Const PROCESS_WRITE_DAC = &H40000
Public Const PROCESS_WRITE_OWNER = &H80000
Public Const PROCESS_STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const PROCESS_STILL_ACTIVE = &H103
Public Const PROCESS_ALL_ACCESS = PROCESS_STANDARD_RIGHTS_REQUIRED + PROCESS_SYNCHRONIZE + &HFFFF&

Public Const QS_KEY = &H1
Public Const QS_MOUSEMOVE = &H2
Public Const QS_MOUSEBUTTON = &H4
Public Const QS_POSTMESSAGE = &H8
Public Const QS_TIMER = &H10
Public Const QS_PAINT = &H20
Public Const QS_SENDMESSAGE = &H40
Public Const QS_HOTKEY = &H80
Public Const QS_ALLPOSTMESSAGE = &H100
Public Const QS_MOUSE = QS_MOUSEMOVE Or QS_MOUSEBUTTON
Public Const QS_INPUT = QS_MOUSE Or QS_KEY
Public Const QS_ALLEVENTS = QS_INPUT Or QS_POSTMESSAGE Or QS_TIMER Or QS_PAINT Or QS_HOTKEY
Public Const QS_ALLINPUT = QS_INPUT Or QS_POSTMESSAGE Or QS_TIMER Or QS_PAINT Or QS_HOTKEY Or QS_SENDMESSAGE

Public Const S_OK = &H0
Public Const S_FALSE = &H1

Public Const SE_ERR_ACCESSDENIED = 5
Public Const SE_ERR_ASSOCINCOMPLETE = 27
Public Const SE_ERR_BAD_FORMAT = 11
Public Const SE_ERR_DDEBUSY = 30
Public Const SE_ERR_DDEFAIL = 29
Public Const SE_ERR_DDETIMEOUT = 28
Public Const SE_ERR_DLLNOTFOUND = 32
Public Const SE_ERR_FILE_NOT_FOUND = 2
Public Const SE_ERR_NOASSOC = 31
Public Const SE_ERR_OOM = 8
Public Const SE_ERR_PATH_NOT_FOUND = 3
Public Const SE_ERR_SHARE = 26
Public Const SE_ISERROR = SE_ERR_DLLNOTFOUND

Public Const SEE_MASK_DEFAULT = &H0
Public Const SEE_MASK_CLASSNAME = &H1
Public Const SEE_MASK_CLASSKEY = &H3
Public Const SEE_MASK_IDLIST = &H4
Public Const SEE_MASK_INVOKEIDLIST = &HC
Public Const SEE_MASK_ICON = &H10
Public Const SEE_MASK_HOTKEY = &H20
Public Const SEE_MASK_NOCLOSEPROCESS = &H40
Public Const SEE_MASK_CONNECTNETDRV = &H80
Public Const SEE_MASK_NOASYNC = &H100
Public Const SEE_MASK_DOENVSUBST = &H200
Public Const SEE_MASK_FLAG_NO_UI = &H400
Public Const SEE_MASK_UNICODE = &H4000
Public Const SEE_MASK_NO_CONSOLE = &H8000
Public Const SEE_MASK_ASYNCOK = &H100000
Public Const SEE_MASK_HMONITOR = &H200000
Public Const SEE_MASK_NOZONECHECKS = &H800000
Public Const SEE_MASK_NOQUERYCLASSSTORE = &H1000000  ' Not used
Public Const SEE_MASK_WAITFORINPUTIDLE = &H2000000
Public Const SEE_MASK_FLAG_LOG_USAGE = &H4000000
Public Const SEE_MASK_FLAG_HINST_IS_SITE = &H8000000

Public Const SW_HIDE = 0
Public Const SW_SHOWNORMAL = 1
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOW = 5
Public Const SW_MINIMIZE = 6
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNA = 8
Public Const SW_RESTORE = 9
Public Const SW_SHOWDEFAULT = 10
Public Const SW_FORCEMINIMIZE = 11
Public Const SW_MAXIMIZE = SW_SHOWMAXIMIZED

Public Const WA_ACTIVE = 1
Public Const WA_CLICKACTIVE = 2
Public Const WA_INACTIVE = 0

Public Const WAIT_OBJECT_0 = &H0
Public Const WAIT_ABANDONED = &H80
Public Const WAIT_IO_COMPLETION = &HC0
Public Const WAIT_TIMEOUT = &H102
Public Const WAIT_FAILED = &HFFFFFFFF
Public Const WAIT_INFINITE = -1

Public Const WS_POPUP = &H80000000

Public Const WS_EX_ACCEPTFILES = &H10
Public Const WS_EX_APPWINDOW = &H40000
Public Const WS_EX_CLIENTEDGE = &H200
Public Const WS_EX_COMPOSITED = &H2000000
Public Const WS_EX_CONTEXTHELP = &H400
Public Const WS_EX_CONTROLPARENT = &H10000
Public Const WS_EX_DLGMODALFRAME = &H1
Public Const WS_EX_LAYERED = &H80000
Public Const WS_EX_LAYOUTRTL = &H400000
Public Const WS_EX_LEFT = &H0
Public Const WS_EX_LEFTSCROLLBAR = &H4000
Public Const WS_EX_TOOLWINDOW = &H80
Public Const WS_EX_TOPMOST = &H8
Public Const WS_EX_TRANSPARENT = &H20
Public Const WS_EX_WINDOWEDGE = &H100

Public Const WM_NULL = &H0
Public Const WM_ACTIVATE = &H6
Public Const WM_CLOSE = &H10
Public Const WM_QUIT = &H12

'''''''''''''''''''''''''''''''''''''''''''''''''
' Human-Readable Windows API Error Descriptions '
'''''''''''''''''''''''''''''''''''''''''''''''''

Public Const kwinErrDescAccessDenied As String = "The operating system denied access to the specified file"
Public Const kwinErrDescAssocIncomplete As String = "The file name association is incomplete or invalid"
Public Const kwinErrDescBadFormat As String = "The .exe file is invalid (non-Win32 .exe or error in .exe image)"
Public Const kwinErrDescDdeBusy As String = "The DDE transaction could not be completed because other DDE transactions were being processed"
Public Const kwinErrDescDdeFail As String = "The DDE transaction failed"
Public Const kwinErrDescDdeTimeout As String = "The DDE transaction could not be completed because the request timed out"
Public Const kwinErrDescDllNotFound As String = "The specified DLL was not found"
Public Const kwinErrDescFileNotFound As String = "The specified file was not found"
Public Const kwinErrDescNoAssoc As String = "There is no application associated with the given file name extension, or file not printable"
Public Const kwinErrDescOutOfMemory As String = "There was not enough memory to complete the operation"
Public Const kwinErrDescPathNotFound As String = "The specified path was not found"
Public Const kwinErrDescShare As String = "A sharing violation occurred"
Public Const kwinErrWaitAbandoned As String = "Wait abandoned"
Public Const kwinErrWaitFailed As String = "Wait failed"
Public Const kwinErrWaitTimeout As String = "Wait timeout"

'''''''''''''''''''''''''''''''''''
' Windows API Based VBA Functions '
'''''''''''''''''''''''''''''''''''

Public Function GetInternetConnectedState() As Boolean
    ' Returns TRUE if an active internet connection exists,
    ' else returns FALSE.
    '
    On Error Resume Next
    GetInternetConnectedState = InternetGetConnectedState(0&, 0&)
End Function

Public Function HwndFromPartialText( _
    ByVal aText As String _
) As LongPtr
    '
    ' Returns the window handle of the first window found having a
    ' caption that partially matches the given text.
    '
    Dim hwnd As LongPtr
    Dim txt As String
    
    HwndFromPartialText = False
    hwnd = FindWindow(vbNullString, vbNullString)
    Do While hwnd <> 0
        txt = String(GetWindowTextLength(hwnd) + 1, Chr$(0))
        GetWindowText hwnd, txt, Len(txt)
        txt = Left$(txt, Len(txt) - 1)
        If InStr(1, txt, aText) > 0 Then
            HwndFromPartialText = hwnd
            Exit Do
        End If
        hwnd = GetWindow(hwnd, GW_HWNDNEXT)
    Loop
End Function

Public Function PidToHwnd( _
    ByVal aPid As Long _
) As LongPtr
    '
    ' Returns the window handle of the top-level window belonging
    ' to the given process (task) id.
    '
    Dim hwnd As LongPtr
    
    hwnd = FindWindow(vbNullString, vbNullString)
    Do While hwnd <> 0
        hwnd = GetWindow(hwnd, GW_HWNDNEXT)
        If IsWindowVisible(hwnd) Then
            If GetParent(hwnd) = 0 Then
                Dim noOwner As Boolean, wstyle As LongPtr
                
                noOwner = (GetWindow(hwnd, GW_OWNER) = 0)
                wstyle = GetWindowLongPtr(hwnd, GWL_EXSTYLE)
                If (((wstyle And WS_EX_TOOLWINDOW) = 0) And noOwner) Or _
                ((wstyle And WS_EX_APPWINDOW) And Not noOwner) Then
                    Dim thid As Long, pid As Long
                    
                    thid = GetWindowThreadProcessId(hwnd, pid)
                    If pid = aPid Then
                        hwnd = GetAncestor(hwnd, GA_ROOT)
                        PidToHwnd = hwnd
                        Exit Function
                    End If
                End If
            End If
        End If
    Loop
End Function

Public Function HprocToHwnd( _
    ByVal aProc As LongPtr _
) As LongPtr
    '
    ' Returns the window handle of the top-level window belonging
    ' to the given process handle.
    '
    HprocToHwnd = PidToHwnd(GetProcessId(aProc))
End Function

Public Function ProcessWaitForExit( _
    ByVal aHprocess As LongPtr, _
    Optional ByVal aTimeout As Long = WAIT_INFINITE _
) As Long
    '
    ' Waits for the process having the given process handle to terminate
    ' for the given timeout period, in milliseconds, and returns one of:
    '   WAIT_OBJECT_0:  the process terminated successfully,
    '   WAIT_TIMEOUT:   the wait period timed out before the process terminated,
    '   WAIT_FAILED:    the wait failed because of a Windows API call error.
    '
    Dim status As Long
    Dim Start As Double
    
    DoEvents
    If GetExitCodeProcess(aHprocess, status) = 0 Then
        ProcessWaitForExit = WAIT_FAILED
    Else
        Start = Timer()
        Do
            DoEvents
            GetExitCodeProcess aHprocess, status
        Loop While status = PROCESS_STILL_ACTIVE And Not TimeoutMs(Timer(), Start, aTimeout)
        If status = PROCESS_STILL_ACTIVE Then ProcessWaitForExit = WAIT_TIMEOUT
    End If
End Function

Private Function ProgramHwnd( _
    ByVal aProgramName As String _
) As LongPtr
    '
    ' Returns the given executable's first window handle, or 0
    ' if the program isn't running or has no windows.
    '
    Dim result As Object, v As Variant
    
    Set result = GetObject("winmgmts:").ExecQuery("select * from win32_process where name='" & aProgramName & "'")
    For Each v In result
        ProgramHwnd = PidToHwnd(v.ProcessId)
        If ProgramHwnd <> 0 Then Exit For
    Next
End Function

Public Function ProgramIsRunning( _
    ByVal aProgramName As String _
) As Boolean
    '
    ' Returns TRUE if the given executable is running,
    ' else returns FALSE.
    '
    ProgramIsRunning = (GetObject("winmgmts:") _
    .ExecQuery("select * from win32_process where name='" & aProgramName & "'").count > 0)
End Function

Public Function ShellEx( _
    ByVal aFile As String, _
    Optional ByVal aShowCmd As Integer = vbNormalFocus, _
    Optional ByVal aOperation As String = "open", _
    Optional ByVal aParameters As String = vbNullString, _
    Optional ByVal aDirectory As String = vbNullString, _
    Optional ByVal aHwnd As LongPtr = 0, _
    Optional ByVal aTimeout As Long = 0, _
    Optional ByVal aFlags As Long = SEE_MASK_NOCLOSEPROCESS _
) As Long
    '
    ' An extended version of the VBA built-in Shell() function that takes
    ' optional parameters, similar to ShellExecute(), and a timeout
    ' parameter that, when non-zero, waits for the shelled application to
    ' finish before returning to the caller. If the ShellEx() function
    ' successfully executes the named file, it returns the process (task)
    ' ID of the started program, unless a timeout is specified, where it
    ' returns 0 after the program completes. If the ShellEx() function can't
    ' start the named program, an error occurs. If a timeout occurs, the
    ' valid process id is returned so the caller can close the program.
    '
    Dim info As SHELLEXECUTEINFO
    
    On Error GoTo Finally
    With info
        .cbSize = LenB(info)
        .fMask = aFlags
        .hwnd = aHwnd
        .lpDirectory = aDirectory
        .lpFile = aFile
        .lpParameters = aParameters
        .lpVerb = aOperation
        .nShow = aShowCmd
        If ShellExecuteEx(info) Then
            If .hInstApp > SE_ISERROR Then
                ShellEx = GetProcessId(.hprocess)
                If aTimeout <> 0 Then
                    Select Case ProcessWaitForExit(.hprocess, aTimeout)
                        Case WAIT_OBJECT_0:
                            ShellEx = 0 ' Success
                        Case WAIT_TIMEOUT:
                            ' Timeout, return the pid.
                        Case Else:
                            Err.Raise UsrErr(Err.LastDllError), "LibWin", "Wait failed"
                    End Select
                End If
            Else
                Err.Raise UsrErr(CLng(.hInstApp)), "LibWin", DllErrDescription(CLng(.hInstApp))
            End If
        Else
            Err.Raise UsrErr(Err.LastDllError), "LibWin", "ShellEx command failed"
        End If
        
Finally:
    If .hprocess <> 0 Then CloseHandle .hprocess
    If Err.Number <> 0 Then Err.Raise Err.Number
    End With
End Function

Public Function WindowHasPopup( _
    ByVal aHwnd As LongPtr _
) As Boolean
    '
    ' Returns TRUE if the given window handle has a child window
    ' that is a popup, else returns FALSE.
    '
    Dim child As LongPtr
    
    child = GetWindow(aHwnd, GW_CHILD)  'Find Child
    If child <> 0 Then WindowHasPopup = ((GetWindowLongPtr(child, GWL_EXSTYLE) And WS_POPUP) <> 0)
    Do While Not (child = 0 Or WindowHasPopup)
        child = GetWindow(child, GW_HWNDNEXT) 'Continue Enumeration
        If child <> 0 Then WindowHasPopup = ((GetWindowLongPtr(child, GWL_EXSTYLE) And WS_POPUP) <> 0)
    Loop
End Function


