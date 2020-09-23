Attribute VB_Name = "mFunctions"
Option Explicit

Public Const WM_KEYDOWN As Long = &H100
Public Const WM_KEYUP As Long = &H101
Public Const WM_SETFOCUS As Long = &H7
Public Const VK_RETURN As Long = &HD
Public Const WM_SETTEXT As Long = &HC
Public Const VK_CONTROL As Long = &H11
Public Const WS_EX_WINDOWEDGE = &H100

Public Enum enWindowStyles
    WS_BORDER = &H800000
    WS_CAPTION = &HC00000
    WS_CHILD = &H40000000
    WS_CLIPCHILDREN = &H2000000
    WS_CLIPSIBLINGS = &H4000000
    WS_DISABLED = &H8000000
    WS_DLGFRAME = &H400000
    WS_EX_ACCEPTFILES = &H10&
    WS_EX_DLGMODALFRAME = &H1&
    WS_EX_NOPARENTNOTIFY = &H4&
    WS_EX_TOPMOST = &H8&
    WS_EX_TRANSPARENT = &H20&
    WS_EX_TOOLWINDOW = &H80&
    WS_GROUP = &H20000
    WS_HSCROLL = &H100000
    WS_MAXIMIZE = &H1000000
    WS_MAXIMIZEBOX = &H10000
    WS_MINIMIZE = &H20000000
    WS_MINIMIZEBOX = &H20000
    WS_OVERLAPPED = &H0&
    WS_POPUP = &H80000000
    WS_SYSMENU = &H80000
    WS_TABSTOP = &H10000
    WS_THICKFRAME = &H40000
    WS_VISIBLE = &H10000000
    WS_VSCROLL = &H200000
    '\\ New from 95/NT4 onwards
    WS_EX_MDICHILD = &H40
    WS_EX_CLIENTEDGE = &H200
    WS_EX_CONTEXTHELP = &H400
    WS_EX_RIGHT = &H1000
    WS_EX_LEFT = &H0
    WS_EX_RTLREADING = &H2000
    WS_EX_LTRREADING = &H0
    WS_EX_LEFTSCROLLBAR = &H4000
    WS_EX_RIGHTSCROLLBAR = &H0
    WS_EX_CONTROLPARENT = &H10000
    WS_EX_STATICEDGE = &H20000
    WS_EX_APPWINDOW = &H40000
    WS_EX_OVERLAPPEDWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_CLIENTEDGE)
    WS_EX_PALETTEWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW Or WS_EX_TOPMOST)
End Enum

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Type OSVERSIONINFOEX
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Public Type DllVersionInfo
    cbSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
End Type

Public Const WM_CLOSE = &H10
Public Const WM_ACTIVATE As Long = &H6
Public Const WM_NCACTIVATE As Long = &H86
Public Const SW_HIDE As Long = 0
Public Const SW_SHOW = 5
Public Const GWL_STYLE = (-16)
Public Const GWL_EXSTYLE = (-20)
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOOWNERZORDER As Long = &H200
Public Const SWP_NOACTIVATE As Long = &H10
Public Const GW_HWNDNEXT As Long = 2
Public Const HWND_BOTTOM As Long = 1
Public Const VER_PLATFORM_WIN32s = 0
Public Const VER_PLATFORM_WIN32_WINDOWS = 1
Public Const VER_PLATFORM_WIN32_NT = 2

Public Declare Function GetTickCount Lib "kernel32.dll" () As Long
Public Declare Function SetActiveWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function GetFocus Lib "user32.dll" () As Long
Public Declare Function InitCommonControls Lib "Comctl32.dll" () As Long
Public Declare Function SetParent Lib "user32.dll" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal Hwnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function IsWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function IsWindowVisible Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function SetFocusA Lib "user32.dll" Alias "SetFocus" (ByVal hwnd As Long) As Long
Public Declare Function GetParent Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function IsZoomed Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function IsIconic Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32.dll" () As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetNextWindow Lib "user32" Alias "GetWindow" (ByVal hwnd As Long, ByVal wFlag As Long) As Long
Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Public Declare Function EnableMenuItem Lib "user32.dll" (ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long
Public Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetClassName Lib "user32.dll" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function DllGetVersion Lib "Shlwapi.dll" (dwVersion As DllVersionInfo) As Long
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFOEX) As Long
Public Declare Function FlashWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
Public Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long


Public Function OSVersion() As Long
    
    Dim udtOSVersion As OSVERSIONINFOEX
    Dim lMajorVersion  As Long
    Dim lMinorVersion As Long
    Dim lPlatformID As Long
    Dim sAns As Long
    
    
    udtOSVersion.dwOSVersionInfoSize = Len(udtOSVersion)
    GetVersionEx udtOSVersion
    lMajorVersion = udtOSVersion.dwMajorVersion
    lMinorVersion = udtOSVersion.dwMinorVersion
    lPlatformID = udtOSVersion.dwPlatformId
    
    Select Case lMajorVersion
        Case 5
        
            ' Added the following to give suppport for Windows XP!
            If lMinorVersion = 0 Then
            
                sAns = 2004
                
            ElseIf lMinorVersion = 1 Then
            
                sAns = 2005
            
            End If
            
                
                
        Case 4
            If lPlatformID = VER_PLATFORM_WIN32_NT Then
                sAns = 2003
            Else
                sAns = IIf(lMinorVersion = 0, _
                2002, 2001)
            End If
        Case 3
            If lPlatformID = VER_PLATFORM_WIN32_NT Then
                sAns = 2000
            Else
                sAns = 1999
            End If
            
        Case Else
            sAns = -1
    End Select
                    
    OSVersion = sAns
    
End Function
Public Function IEVersionShort() As Long
    Dim udtVersionInfo As DllVersionInfo
    udtVersionInfo.cbSize = Len(udtVersionInfo)
    Call DllGetVersion(udtVersionInfo)
    IEVersionShort = udtVersionInfo.dwMajorVersion
End Function


Public Function IEVersionLong() As String
    Dim udtVersionInfo As DllVersionInfo
    udtVersionInfo.cbSize = Len(udtVersionInfo)
    Call DllGetVersion(udtVersionInfo)
    IEVersionLong = "Internet Explorer " & _
    udtVersionInfo.dwMajorVersion & "." & _
    udtVersionInfo.dwMinorVersion & "." & _
    udtVersionInfo.dwBuildNumber
End Function

Function InternetExplorerPath() As String
    Const HKEY_LOCAL_MACHINE = &H80000002
    ' get the path from the registry, or return a null string
    InternetExplorerPath = ReadRegistry(HKEY_LOCAL_MACHINE, _
        "Software\Microsoft\Windows\CurrentVersion\App Paths\IEXPLORE.EXE", _
        "Path")
    ' get rid of the trailing semi-colon, if there
    If Right$(InternetExplorerPath, 1) = ";" Then
        InternetExplorerPath = Left$(InternetExplorerPath, _
            Len(InternetExplorerPath) - 1)
    End If
        
End Function


Public Function RemoveUnderscore(S As String) As String
    Dim Out As String, I As Long, C As String
    For I = 1 To Len(S)
        C = Mid(S, I, 1)
        If C = "&" Then
            Out = Out & "&&"
        Else
            Out = Out & C
        End If
    Next I
    RemoveUnderscore = Out
End Function
Public Function n_GetWindowText(hwnd As Long) As String
    Dim L As Long, S As String
    L = GetWindowTextLength(hwnd) + 1
    S = String(L, " ")
    GetWindowText hwnd, S, L
    n_GetWindowText = Mid(S, 1, Len(S) - 1)
End Function
Public Function SetWindowStyle(ByVal hwnd As Long, ByVal extended_style As Boolean, ByVal style_value As enWindowStyles, ByVal new_value As Boolean, ByVal brefresh As Boolean) As Long
    Dim style_type As Long
    Dim style As Long
    
    style_type = IIf(extended_style, GWL_EXSTYLE, GWL_STYLE)
    style = GetWindowLong(hwnd, style_type)
    style = IIf(new_value, style Or style_value, style And Not style_value)
    If brefresh Then ShowWindow hwnd, SW_HIDE
    SetWindowLong hwnd, style_type, style
    If brefresh Then ShowWindow hwnd, SW_SHOW
    SetWindowStyle = SetWindowPos(hwnd, 0, 0, 0, 0, 0, SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER)
End Function
Public Sub CloseHwnd(Whwnd)
    PostMessage Whwnd, WM_CLOSE, 0&, ByVal 0&
End Sub

