Attribute VB_Name = "modCBDeclares"
'--------------------------------------------------------------------------------------------
'   Name:           modCBDeclares (MODCBDECLARES.BAS)
'   Type:           Declarations Module
'   Description:    All my favorite declarations and utility macros and functions and stuff.
'
'   Author:         Klaus H. Probst [kprobst@vbbox.com]
'   URL:            http://www.vbbox.com/
'   Copyright:      None
'   Usage:          You may use this code as you see fit, provided that you assume all
'                   responsibilities for doing so.
'   Distribution:   If you intend to distribute the file(s) that make up this sample to
'                   any WWW site, online service, electronic bulletin board system (BBS),
'                   CD or any other electronic or physical media, you must notify me in
'                   advance to obtain my express permission.
'
'   Notes:
'
'   Dependencies:
'
'       (none)
'
'   Updated:        03/27/01
'   Revision:       1
'
'--------------------------------------------------------------------------------------------
#If Not CB_NO_MAIN_SYMBOLS Then

Option Explicit
DefLng A-Z

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Type POINTAPI
    x As Long
    y As Long
End Type

Type POINTS
    x As Integer
    y As Integer
End Type

Type NMHDR
    hwndFrom As Long
    idfrom As Long
    Code As Long
End Type

Type MEASUREITEMSTRUCT
    CtlType As Long
    CtlID As Long
    itemID As Long
    itemWidth As Long
    ItemHeight As Long
    ItemData As Long
End Type

' DRAWITEMSTRUCT for ownerdraw
Type DRAWITEMSTRUCT
    CtlType As Long
    CtlID As Long
    itemID As Long
    itemAction As Long
    itemState As Long
    hwndItem As Long
    hDC As Long
    rcItem As RECT
    ItemData As Long
End Type

Type DELETEITEMSTRUCT
    CtlType As Long
    CtlID As Long
    itemID As Long
    hwndItem As Long
    ItemData As Long
End Type


Type DRAWTEXTPARAMS
    cbSize As Long
    iTabLength As Long
    iLeftMargin As Long
    iRightMargin As Long
    uiLengthDrawn As Long
End Type

Type SIZE
    cx As Long
    cy As Long
End Type


Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Type PICTDESC
    cbSize As Long
    picType As Long
    hImage As Long
    Data1 As Long
    Data2 As Long
End Type

Type WINDOWPLACEMENT
    Length As Long
    Flags As Long
    ShowCmd As Long
    ptMinPosition As POINTAPI
    ptMaxPosition As POINTAPI
    rcNormalPosition As RECT
End Type

Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type

Type WINDOWPOS
    hWnd As Long
    hWndInsertAfter As Long
    x As Long
    y As Long
    cx As Long
    cy As Long
    Flags As Long
End Type

Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Public Const MAX_PATH = 260

Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * MAX_PATH
    szTypeName As String * 80
End Type

Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

' Logical Font
Private Const LF_FACESIZE = 32
Private Const LF_FULLFACESIZE = 64

Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName As String * LF_FACESIZE
End Type

Type tagMSG         '// just using "MSG" screws things up everywhere
    hWnd As Long
    Message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

Type SCROLLINFO
    cbSize As Long
    fMask As Long
    nMin As Long
    nMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long
End Type

Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Type COPYDATASTRUCT
    dwData As Long
    cbData As Long
    lpData As Long
End Type


'// Window attribute and creation functions
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long

'// Window text functions
Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long

'// Time/Date functions
Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long

'// File manipulation functions
Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Declare Function SearchPath Lib "kernel32" Alias "SearchPathA" (ByVal lpPath As String, ByVal lpFileName As String, ByVal lpExtension As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Declare Function RemoveDirectory Lib "kernel32" Alias "RemoveDirectoryA" (ByVal lpPathName As String) As Long
Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Declare Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, lpFilePart As Any) As Long
Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Declare Function FindFirstChangeNotification Lib "kernel32" Alias "FindFirstChangeNotificationA" (ByVal lpPathName As String, ByVal bWatchSubtree As Long, ByVal dwNotifyFilter As Long) As Long
Declare Function FindCloseChangeNotification Lib "kernel32" (ByVal hChangeHandle As Long) As Long
Declare Function FindNextChangeNotification Lib "kernel32" (ByVal hChangeHandle As Long) As Long
Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

'// "Legacy" file operation APIs
Declare Function lopen Lib "kernel32" Alias "_lopen" (ByVal lpPathName As String, ByVal iReadWrite As Long) As Long
Declare Function lclose Lib "kernel32" Alias "_lclose" (ByVal hFile As Long) As Long
Declare Function lcreat Lib "kernel32" Alias "_lcreat" (ByVal lpPathName As String, ByVal iAttribute As Long) As Long
Declare Function llseek Lib "kernel32" Alias "_llseek" (ByVal hFile As Long, ByVal lOffset As Long, ByVal iOrigin As Long) As Long
Declare Function lread Lib "kernel32" Alias "_lread" (ByVal hFile As Long, lpBuffer As Any, ByVal wBytes As Long) As Long
Declare Function lwrite Lib "kernel32" Alias "_lwrite" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal wBytes As Long) As Long
Declare Function hread Lib "kernel32" Alias "_hread" (ByVal hFile As Long, lpBuffer As Any, ByVal lBytes As Long) As Long
Declare Function hwrite Lib "kernel32" Alias "_hwrite" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal lBytes As Long) As Long

'// Library and process functions
Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

'// Shell library functions
Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As Any, ByVal lpParameters As Any, ByVal lpDirectory As Any, ByVal nShowCmd As Long) As Long
Declare Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal hDrop As Long, ByVal UINT As Long, ByVal lpStr As String, ByVal ch As Long) As Long
Declare Function DragQueryPoint Lib "shell32.dll" (ByVal hDrop As Long, lpPoint As POINTAPI) As Long
Declare Sub DragFinish Lib "shell32.dll" (ByVal hDrop As Long)
Declare Sub DragAcceptFiles Lib "shell32.dll" (ByVal hWnd As Long, ByVal fAccept As Long)

'// Window Z-order and placement APIs
Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function OpenIcon Lib "user32" (ByVal hWnd As Long) As Long
Declare Function CloseWindow Lib "user32" (ByVal hWnd As Long) As Long
Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function GetWindowPlacement Lib "user32" (ByVal hWnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Declare Function SetWindowPlacement Lib "user32" (ByVal hWnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Declare Function BeginDeferWindowPos Lib "user32" (ByVal nNumWindows As Long) As Long
Declare Function DeferWindowPos Lib "user32" (ByVal hWinPosInfo As Long, ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function EndDeferWindowPos Lib "user32" (ByVal hWinPosInfo As Long) As Long
Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Declare Function SetActiveWindow Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetForegroundWindow Lib "user32" () As Long
Declare Function GetActiveWindow Lib "user32" () As Long
Declare Function win32_SetFocus Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Declare Function win32_GetFocus Lib "user32" Alias "GetFocus" () As Long
Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Declare Function IsWindowEnabled Lib "user32" (ByVal hWnd As Long) As Long
Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Declare Function ShowWindowAsync Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Declare Function AdjustWindowRectEx Lib "user32" (lpRect As RECT, ByVal dsStyle As Long, ByVal bMenu As Long, ByVal dwEsStyle As Long) As Long

'// Window location
Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd As Long, ByVal hWndChild As Long, ByVal lpszClassName As Any, ByVal lpszWindow As Any) As Long
Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long


'// Keyboard and mouse input functions
Declare Function GetCapture Lib "user32" () As Long
Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Declare Function ReleaseCapture Lib "user32" () As Long
Declare Function GetKBCodePage Lib "user32" () As Long
Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Declare Function SetKeyboardState Lib "user32" (lppbKeyState As Byte) As Long
Declare Function GetKeyboardType Lib "user32" (ByVal nTypeFlag As Long) As Long
Declare Function GetKeyNameText Lib "user32" Alias "GetKeyNameTextA" (ByVal lParam As Long, ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function ToAscii Lib "user32" (ByVal uVirtKey As Long, ByVal uScanCode As Long, lpbKeyState As Byte, lpwTransKey As Long, ByVal fuState As Long) As Long
Declare Function ToUnicode Lib "user32" (ByVal wVirtKey As Long, ByVal wScanCode As Long, lpKeyState As Byte, ByVal pwszBuff As String, ByVal cchBuff As Long, ByVal wFlags As Long) As Long
Declare Function OemKeyScan Lib "user32" (ByVal wOemChar As Long) As Long
Declare Function VkKeyScan Lib "user32" Alias "VkKeyScanA" (ByVal cChar As Byte) As Integer
Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long
Declare Function GetInputState Lib "user32" () As Long
Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal dwData As Long, ByVal dwExtraInfo As Long)
Declare Function AttachThreadInput Lib "user32" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessID As Long) As Long
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'// OLE/COM functions
Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (pPictDesc As PICTDESC, refiid As GUID, ByVal fPictureOwnsHandle As Long, ppvObj As Object) As Long    '// ppvObj = StdPicture
Declare Sub OleTranslateColor Lib "oleaut32.dll" (ByVal ColorIn As Long, ByVal hPal As Long, ByRef RGBColorOut As Long)
Declare Function CoTaskMemAlloc Lib "OLE32.DLL" (ByVal cb As Long) As Long
Declare Function CoTaskMemRealloc Lib "OLE32.DLL" (ByVal pv As Long, ByVal cb As Long) As Long
Declare Sub CoTaskMemFree Lib "OLE32.DLL" (ByVal pv As Long)
Declare Function CoInitialize Lib "OLE32.DLL" (ByVal pvReserved As Any) As Long
Declare Sub CoUninitialize Lib "OLE32.DLL" ()

'// Memory functions
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal nCount As Long)
Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)
Declare Function GetProcessHeap Lib "kernel32" () As Long
Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
Declare Function LocalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal uBytes As Long) As Long
Public Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long

'// Messaging functions
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'// Window graphic functions
Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long
Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long
Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As Any, ByVal bErase As Long) As Long

'// Mapping functions
Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Declare Function MapWindowPoints Lib "user32" (ByVal hwndFrom As Long, ByVal hwndTo As Long, lppt As Any, ByVal cPoints As Long) As Long
Declare Function ChildWindowFromPoint Lib "user32" (ByVal hWnd As Long, ByVal xPoint As Long, ByVal yPoint As Long) As Long

'// System Functions
Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Declare Function GetDesktopWindow Lib "user32" () As Long
Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, arguments As Long) As Long

'// Font functions
Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long

'// RECT Functions
Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Declare Function UnionRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
Declare Function IsRectEmpty Lib "user32" (lpRect As RECT) As Long
Declare Function SetRectEmpty Lib "user32" (lpRect As RECT) As Long
Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Declare Function SubtractRect Lib "user32" (lprcDst As RECT, lprcSrc1 As RECT, lprcSrc2 As RECT) As Long
Declare Function IntersectRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
Declare Function CopyRect Lib "user32" (lpDestRect As RECT, lpSourceRect As RECT) As Long

'// Cursor functions
Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
Declare Function GetCursor Lib "user32" () As Long
Declare Function GetClipCursor Lib "user32" (lprc As RECT) As Long

'// Bitmap functions
Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Declare Function GetBitmapDimensionEx Lib "gdi32" (ByVal hBitmap As Long, lpDimension As SIZE) As Long

'// DC Functions
Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Declare Function GetTextColor Lib "gdi32" (ByVal hDC As Long) As Long
Declare Function WindowFromDC Lib "user32" (ByVal hDC As Long) As Long
Declare Function GetBkColor Lib "gdi32" (ByVal hDC As Long) As Long
Declare Function GetBkMode Lib "gdi32" (ByVal hDC As Long) As Long
Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long

'// Drawing Functions
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long
Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Declare Function DrawIcon Lib "user32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Declare Function DrawAnimatedRects Lib "user32" (ByVal hWnd As Long, ByVal idAni As Long, lprcFrom As RECT, lprcTo As RECT) As Long
Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hDC As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Any, ByVal wParam As Any, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal fuFlags As Long) As Long
Declare Function DrawFrameControl Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal uType As Long, ByVal uState As Long) As Long
Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, lpPoint As Any) As Long
Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Declare Function ExtFloodFill Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Declare Function FrameRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Declare Function InvertRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long
Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Declare Function DrawCaption Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long, ByRef lprc As RECT, ByVal uFlags As Long) As Long

'// Text rendering functions
Declare Function ExtTextOut Lib "gdi32" Alias "ExtTextOutA" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal wOptions As Long, lpRect As Any, ByVal lpString As String, ByVal nCount As Long, lpDx As Long) As Long
Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hDC As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, lpDrawTextParams As Any) As Long
Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hDC As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As SIZE) As Long
Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long


'// Resource APIs
Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal iImageType As Long, ByVal cx As Long, ByVal cy As Long, ByVal fFlags As Long) As Long
Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As String) As Long
Declare Function LoadCursorFromFile Lib "user32" Alias "LoadCursorFromFileA" (ByVal lpFileName As String) As Long
Declare Function LoadBitmap Lib "user32" Alias "LoadBitmapA" (ByVal hInstance As Long, ByVal lpBitmapName As String) As Long
Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Long
Declare Function CopyCursor Lib "user32" (ByVal hcur As Long) As Long
Declare Function CopyIcon Lib "user32" (ByVal hIcon As Long) As Long
Declare Function LoadIcon Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As String) As Long
Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Declare Function LoadString Lib "user32" Alias "LoadStringA" (ByVal hInstance As Long, ByVal wID As Long, ByVal lpBuffer As String, ByVal nBufferMax As Long) As Long

'// API string functions
Declare Function lstrcmpiA Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
Declare Function lstrcmpiW Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
Declare Function CharNextA Lib "user32" (ByVal lpsz As Any) As Long
Declare Function CharNextW Lib "user32" (ByVal lpsz As Any) As Long
Declare Function lstrlenA Lib "kernel32" (ByVal psString As Any) As Long
Declare Function lstrlenW Lib "kernel32" (ByVal psString As Any) As Long
Declare Function lstrcpyA Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
Declare Function lstrcpyW Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
Declare Function lstrcpynA Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any, ByVal iMaxLength As Long) As Long
Declare Function lstrcpynW Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any, ByVal iMaxLength As Long) As Long
Declare Function wvsprintfA Lib "user32" (ByVal sBuffer As String, ByVal lpszFormat As String, ByRef arglist As Long) As Long
Declare Function wvsprintfW Lib "user32" (ByVal sBuffer As Long, ByVal lpszFormat As Long, ByRef arglist As Long) As Long
Declare Function CompareString Lib "kernel32" Alias "CompareStringA" (ByVal Locale As Long, ByVal dwCmpFlags As Long, ByVal lpString1 As Any, ByVal cchCount1 As Long, ByVal lpString2 As Any, ByVal cchCount2 As Long) As Long


#If UNICODE Then
Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyW" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
Declare Function lstrlen Lib "kernel32" Alias "lstrlenW" (ByVal psString As Any) As Long
Declare Function lstrcpyn Lib "kernel32" Alias "lstrcpynW" (ByVal lpString1 As Any, ByVal lpString2 As Any, ByVal iMaxLength As Long) As Long
Declare Function CharNext Lib "user32" Alias "CharNextW" (ByVal lpsz As Any) As Long
Declare Function wvsprintf Lib "user32" Alias "wvsprintfW" (ByVal sBuffer As Long, ByVal lpszFormat As Long, ByRef arglist As Long) As Long
Declare Function lstrcmpi Lib "kernel32" Alias "lstrcmpiW" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
#Else
Declare Function wvsprintf Lib "user32" Alias "wvsprintfA" (ByVal sBuffer As String, ByVal lpszFormat As String, ByRef arglist As Long) As Long
Declare Function CharNext Lib "user32" Alias "CharNextA" (ByVal lpsz As Any) As Long
Declare Function lstrcpyn Lib "kernel32" Alias "lstrcpynA" (ByVal lpString1 As Any, ByVal lpString2 As Any, ByVal iMaxLength As Long) As Long
Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal psString As Any) As Long
Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
Declare Function lstrcmpi Lib "kernel32" Alias "lstrcmpiA" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
#End If

Declare Function MultiByteToWideChar Lib "kernel32" (ByVal codepage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Any, ByVal cchWideChar As Long) As Long
Declare Function WideCharToMultiByte Lib "kernel32" (ByVal codepage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Any, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As String, lpUsedDefaultChar As Long) As Long


'// Math and number APIs
Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long

'// Message and hotkey registration
Declare Function RegisterHotkey Lib "user32" Alias "RegisterHotKey" (ByVal hWnd As Long, ByVal ID As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Declare Function UnregisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal ID As Long) As Long
Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long

'// Scrollbar APIs
Declare Function EnableScrollBar Lib "user32" (ByVal hWnd As Long, ByVal wSBflags As Long, ByVal wArrows As Long) As Long
Declare Function SetScrollPos Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long, ByVal nPos As Long, ByVal bRedraw As Long) As Long
Declare Function GetScrollPos Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long) As Long
Declare Function SetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal n As Long, lpcScrollInfo As SCROLLINFO, ByVal bool As Long) As Long
Declare Function GetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal n As Long, lpScrollInfo As SCROLLINFO) As Long

'// Environment APIs
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long


Public Const NORM_IGNORECASE = &H1                  '// ignore case
Public Const NORM_IGNORENONSPACE = &H2              '// ignore nonspacing chars
Public Const NORM_IGNORESYMBOLS = &H4               '// ignore symbols

Public Const NORM_IGNOREKANATYPE = &H10000          '// ignore kanatype
Public Const NORM_IGNOREWIDTH = &H20000             '// ignore width

Public Const CSTR_LESS_THAN = 1                     '// string 1 less than string 2
Public Const CSTR_EQUAL = 2                         '// string 1 equal to string 2
Public Const CSTR_GREATER_THAN = 3                  '// string 1 greater than string 2

'// Very bad. Should be obtained using MAKELCID()
Public Const LOCALE_SYSTEM_DEFAULT = &H800
Public Const LOCALE_USER_DEFAULT = &H400


'// DrawCaption flags
Public Const DC_ACTIVE = &H1
Public Const DC_SMALLCAP = &H2
Public Const DC_ICON = &H4
Public Const DC_TEXT = &H8
Public Const DC_INBUTTON = &H10
'#if(WINVER >= &H0500)
Public Const DC_GRADIENT = &H20


' GetWindow() Constants
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GW_OWNER = 4
Public Const GW_CHILD = 5
Public Const GW_MAX = 5


'// Heap constants
Public Const HEAP_ZERO_MEMORY = &H8
Public Const HEAP_GENERATE_EXCEPTIONS = &H4

'// LocalAlloc flags
Public Const LMEM_FIXED = &H0
Public Const LMEM_ZEROINIT = &H40
Public Const lptr = (LMEM_FIXED Or LMEM_ZEROINIT)


'// File notification constants
Public Const FILE_NOTIFY_CHANGE_FILE_NAME As Long = &H1
Public Const FILE_NOTIFY_CHANGE_DIR_NAME As Long = &H2
Public Const FILE_NOTIFY_CHANGE_ATTRIBUTES As Long = &H4
Public Const FILE_NOTIFY_CHANGE_SIZE As Long = &H8
Public Const FILE_NOTIFY_CHANGE_LAST_WRITE As Long = &H10
Public Const FILE_NOTIFY_CHANGE_LAST_ACCESS As Long = &H20
Public Const FILE_NOTIFY_CHANGE_CREATION As Long = &H40
Public Const FILE_NOTIFY_CHANGE_SECURITY As Long = &H100
Public Const FILE_NOTIFY_FLAGS = FILE_NOTIFY_CHANGE_ATTRIBUTES Or _
                                 FILE_NOTIFY_CHANGE_FILE_NAME Or _
                                 FILE_NOTIFY_CHANGE_LAST_WRITE


'// Wait function flags
Public Const INFINITE = &HFFFF
Public Const WAIT_OBJECT_0 = &H0
Public Const WAIT_ABANDONED = &H80
Public Const WAIT_IO_COMPLETION = &HC0
Public Const WAIT_TIMEOUT = &H102
Public Const STATUS_PENDING = &H103
Public Const WAIT_FAILED = 1        '//&HFFFFFFFF


'// keyb_event constants
Public Const KEYEVENTF_EXTENDEDKEY = &H1
Public Const KEYEVENTF_KEYUP = &H2
Public Const KEYEVENTF_UNICODE = &H4
Public Const KEYEVENTF_SCANCODE = &H8

'// mouse_event constants
Public Const MOUSEEVENTF_MOVE = &H1          '// mouse move */
Public Const MOUSEEVENTF_LEFTDOWN = &H2      '// left button down */
Public Const MOUSEEVENTF_LEFTUP = &H4        '// left button up */
Public Const MOUSEEVENTF_RIGHTDOWN = &H8     '// right button down */
Public Const MOUSEEVENTF_RIGHTUP = &H10      '// right button up */
Public Const MOUSEEVENTF_MIDDLEDOWN = &H20   '// middle button down */
Public Const MOUSEEVENTF_MIDDLEUP = &H40     '// middle button up */
Public Const MOUSEEVENTF_XDOWN = &H80        '// x button down */
Public Const MOUSEEVENTF_XUP = &H100         '// x button down */
Public Const MOUSEEVENTF_WHEEL = &H800       '// wheel button rolled */
Public Const MOUSEEVENTF_VIRTUALDESK = &H4000 '// map to entire virtual desktop */
Public Const MOUSEEVENTF_ABSOLUTE = &H8000   '// absolute move */


Public Const DRIVE_NONE = 0
Public Const DRIVE_ROOTMISSING = 1
Public Const DRIVE_REMOVABLE = 2
Public Const DRIVE_FIXED = 3
Public Const DRIVE_REMOTE = 4
Public Const DRIVE_CDROM = 5
Public Const DRIVE_RAMDISK = 6

Public Const OF_READ = &H0
Public Const OF_WRITE = &H1
Public Const OF_READWRITE = &H2
Public Const OF_SHARE_COMPAT = &H0
Public Const OF_SHARE_EXCLUSIVE = &H10
Public Const OF_SHARE_DENY_WRITE = &H20
Public Const OF_SHARE_DENY_READ = &H30
Public Const OF_SHARE_DENY_NONE = &H40
Public Const OF_PARSE = &H100
Public Const OF_DELETE = &H200
Public Const OF_VERIFY = &H400
Public Const OF_CANCEL = &H800
Public Const OF_CREATE = &H1000
Public Const OF_PROMPT = &H2000
Public Const OF_EXIST = &H4000
Public Const OF_REOPEN = &H8000

Public Const HFILE_ERROR = (-1)
Public Const ERROR_NO_MORE_FILES = 18

'// File attribute masks
Public Const FILE_ATTRIBUTE_READONLY = &H1
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_SYSTEM = &H4
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_TEMPORARY = &H100
Public Const FILE_ATTRIBUTE_COMPRESSED = &H800
Public Const FILE_ATTRIBUTE_OFFLINE = &H1000

'// ShGetFileInfo() flags
Public Const SHGFI_ICON = &H100                         '  get icon
Public Const SHGFI_DISPLAYNAME = &H200                  '  get display name
Public Const SHGFI_TYPENAME = &H400                     '  get type name
Public Const SHGFI_ATTRIBUTES = &H800                   '  get attributes
Public Const SHGFI_ICONLOCATION = &H1000                '  get icon location
Public Const SHGFI_EXETYPE = &H2000                     '  return exe type
Public Const SHGFI_SYSICONINDEX = &H4000                '  get system icon index
Public Const SHGFI_LINKOVERLAY = &H8000                 '  put a link overlay on icon
Public Const SHGFI_SELECTED = &H10000                   '  show icon in selected state
Public Const SHGFI_LARGEICON = &H0                      '  get large icon
Public Const SHGFI_SMALLICON = &H1                      '  get small icon
Public Const SHGFI_OPENICON = &H2                       '  get open icon
Public Const SHGFI_SHELLICONSIZE = &H4                  '  get shell size icon
Public Const SHGFI_PIDL = &H8                           '  pszPath is a pidl
Public Const SHGFI_USEFILEATTRIBUTES = &H10             '  use passed dwFileAttribute


'// Misc constants
Public Const INVALID_HANDLE_VALUE = -1
Public Const MAXDWORD As Long = &HFFFFFFFF

' OEM Resource Ordinal Numbers
Public Const OBM_CLOSE = 32754
Public Const OBM_UPARROW = 32753
Public Const OBM_DNARROW = 32752
Public Const OBM_RGARROW = 32751
Public Const OBM_LFARROW = 32750
Public Const OBM_REDUCE = 32749
Public Const OBM_ZOOM = 32748
Public Const OBM_RESTORE = 32747
Public Const OBM_REDUCED = 32746
Public Const OBM_ZOOMD = 32745
Public Const OBM_RESTORED = 32744
Public Const OBM_UPARROWD = 32743
Public Const OBM_DNARROWD = 32742
Public Const OBM_RGARROWD = 32741
Public Const OBM_LFARROWD = 32740
Public Const OBM_MNARROW = 32739
Public Const OBM_COMBO = 32738
Public Const OBM_UPARROWI = 32737
Public Const OBM_DNARROWI = 32736
Public Const OBM_RGARROWI = 32735
Public Const OBM_LFARROWI = 32734

Public Const OBM_OLD_CLOSE = 32767
Public Const OBM_SIZE = 32766
Public Const OBM_OLD_UPARROW = 32765
Public Const OBM_OLD_DNARROW = 32764
Public Const OBM_OLD_RGARROW = 32763
Public Const OBM_OLD_LFARROW = 32762
Public Const OBM_BTSIZE = 32761
Public Const OBM_CHECK = 32760
Public Const OBM_CHECKBOXES = 32759
Public Const OBM_BTNCORNERS = 32758
Public Const OBM_OLD_REDUCE = 32757
Public Const OBM_OLD_ZOOM = 32756
Public Const OBM_OLD_RESTORE = 32755

Public Const OCR_NORMAL = 32512
Public Const OCR_IBEAM = 32513
Public Const OCR_WAIT = 32514
Public Const OCR_CROSS = 32515
Public Const OCR_UP = 32516
Public Const OCR_SIZE = 32640
Public Const OCR_ICON = 32641
Public Const OCR_SIZENWSE = 32642
Public Const OCR_SIZENESW = 32643
Public Const OCR_SIZEWE = 32644
Public Const OCR_SIZENS = 32645
Public Const OCR_SIZEALL = 32646
Public Const OCR_ICOCUR = 32647
Public Const OCR_NO = 32648 ' not in win3.1

Public Const OIC_SAMPLE = 32512
Public Const OIC_HAND = 32513
Public Const OIC_QUES = 32514
Public Const OIC_BANG = 32515
Public Const OIC_NOTE = 32516

' Standard Icon IDs
Public Const IDI_APPLICATION = 32512&
Public Const IDI_HAND = 32513&
Public Const IDI_QUESTION = 32514&
Public Const IDI_EXCLAMATION = 32515&
Public Const IDI_ASTERISK = 32516&


' Standard Cursor IDs
Public Const IDC_ARROW = 32512&
Public Const IDC_IBEAM = 32513&
Public Const IDC_WAIT = 32514&
Public Const IDC_CROSS = 32515&
Public Const IDC_UPARROW = 32516&
Public Const IDC_SIZE = 32640&
Public Const IDC_ICON = 32641&
Public Const IDC_SIZENWSE = 32642&
Public Const IDC_SIZENESW = 32643&
Public Const IDC_SIZEWE = 32644&
Public Const IDC_SIZENS = 32645&
Public Const IDC_SIZEALL = 32646&
Public Const IDC_NO = 32648&
Public Const IDC_APPSTARTING = 32650&

'// LoadImage() image types
Public Const IMAGE_BITMAP = 0
Public Const IMAGE_ICON = 1
Public Const IMAGE_CURSOR = 2
Public Const IMAGE_ENHMETAFILE = 3

'// LoadImage() flags
Public Const LR_DEFAULTCOLOR = &H0
Public Const LR_MONOCHROME = &H1
Public Const LR_COLOR = &H2
Public Const LR_COPYRETURNORG = &H4
Public Const LR_COPYDELETEORG = &H8
Public Const LR_LOADFROMFILE = &H10
Public Const LR_LOADTRANSPARENT = &H20
Public Const LR_DEFAULTSIZE = &H40
Public Const LR_LOADMAP3DCOLORS = &H1000
Public Const LR_CREATEDIBHEADER = &H2000
Public Const LR_COPYFROMRESOURCE = &H4000
Public Const LR_SHARED = &H8000

'  Ternary raster operations
Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Public Const SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest
Public Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest
Public Const SRCINVERT = &H660046       ' (DWORD) dest = source XOR dest
Public Const SRCERASE = &H440328        ' (DWORD) dest = source AND (NOT dest )
Public Const NOTSRCCOPY = &H330008      ' (DWORD) dest = (NOT source)
Public Const NOTSRCERASE = &H1100A6     ' (DWORD) dest = (NOT src) AND (NOT dest)
Public Const MERGECOPY = &HC000CA       ' (DWORD) dest = (source AND pattern)
Public Const MERGEPAINT = &HBB0226      ' (DWORD) dest = (NOT source) OR dest
Public Const PATCOPY = &HF00021 ' (DWORD) dest = pattern
Public Const PATPAINT = &HFB0A09        ' (DWORD) dest = DPSnoo
Public Const PATINVERT = &H5A0049       ' (DWORD) dest = pattern XOR dest
Public Const DSTINVERT = &H550009       ' (DWORD) dest = (NOT dest)
Public Const BLACKNESS = &H42 ' (DWORD) dest = BLACK
Public Const WHITENESS = &HFF0062       ' (DWORD) dest = WHITE

'// DrawText() Format Flags
Public Const DT_TOP = &H0
Public Const DT_LEFT = &H0
Public Const DT_CENTER = &H1
Public Const DT_RIGHT = &H2
Public Const DT_VCENTER = &H4
Public Const DT_BOTTOM = &H8
Public Const DT_WORDBREAK = &H10
Public Const DT_SINGLELINE = &H20
Public Const DT_EXPANDTABS = &H40
Public Const DT_TABSTOP = &H80
Public Const DT_NOCLIP = &H100
Public Const DT_EXTERNALLEADING = &H200
Public Const DT_CALCRECT = &H400
Public Const DT_NOPREFIX = &H800
Public Const DT_INTERNAL = &H1000
Public Const DT_EDITCONTROL = &H2000
Public Const DT_PATH_ELLIPSIS = &H4000
Public Const DT_END_ELLIPSIS = &H8000
Public Const DT_MODIFYSTRING = &H10000
Public Const DT_RTLREADING = &H20000
Public Const DT_WORD_ELLIPSIS = &H40000

'// ExtFloodFill style flags
Public Const FLOODFILLBORDER = 0
Public Const FLOODFILLSURFACE = 1

Public Const RDW_INVALIDATE = &H1
Public Const RDW_INTERNALPAINT = &H2
Public Const RDW_ERASE = &H4

Public Const RDW_VALIDATE = &H8
Public Const RDW_NOINTERNALPAINT = &H10
Public Const RDW_NOERASE = &H20

Public Const RDW_NOCHILDREN = &H40
Public Const RDW_ALLCHILDREN = &H80

Public Const RDW_UPDATENOW = &H100
Public Const RDW_ERASENOW = &H200

Public Const RDW_FRAME = &H400
Public Const RDW_NOFRAME = &H800

Public Const PS_SOLID = 0
Public Const PS_DASH = 1                    '  -------
Public Const PS_DOT = 2                     '  .......
Public Const PS_DASHDOT = 3                 '  _._._._
Public Const PS_DASHDOTDOT = 4              '  _.._.._
Public Const PS_NULL = 5
Public Const PS_INSIDEFRAME = 6
Public Const PS_USERSTYLE = 7
Public Const PS_ALTERNATE = 8
Public Const PS_STYLE_MASK = &HF

Public Const PS_ENDCAP_ROUND = &H0
Public Const PS_ENDCAP_SQUARE = &H100
Public Const PS_ENDCAP_FLAT = &H200
Public Const PS_ENDCAP_MASK = &HF00

Public Const PS_JOIN_ROUND = &H0
Public Const PS_JOIN_BEVEL = &H1000
Public Const PS_JOIN_MITER = &H2000
Public Const PS_JOIN_MASK = &HF000

Public Const PS_COSMETIC = &H0
Public Const PS_GEOMETRIC = &H10000
Public Const PS_TYPE_MASK = &HF0000


'// ExtTextOut() flags
Public Const ETO_GRAYED = 1
Public Const ETO_OPAQUE = 2
Public Const ETO_CLIPPED = 4

'// Constants used in DrawFrameControl()
'// Values for uType
Public Const DFC_CAPTION = 1
Public Const DFC_MENU = 2
Public Const DFC_SCROLL = 3
Public Const DFC_BUTTON = 4

'// Values for uState
'// uType=DFC_CAPTION
Public Const DFCS_CAPTIONCLOSE = &H0
Public Const DFCS_CAPTIONMIN = &H1
Public Const DFCS_CAPTIONMAX = &H2
Public Const DFCS_CAPTIONRESTORE = &H3
Public Const DFCS_CAPTIONHELP = &H4

'// uType=DFC_MENU
Public Const DFCS_MENUARROW = &H0
Public Const DFCS_MENUCHECK = &H1
Public Const DFCS_MENUBULLET = &H2
Public Const DFCS_MENUARROWRIGHT = &H4

'// uType=DFC_SCROLL
Public Const DFCS_SCROLLUP = &H0
Public Const DFCS_SCROLLDOWN = &H1
Public Const DFCS_SCROLLLEFT = &H2
Public Const DFCS_SCROLLRIGHT = &H3
Public Const DFCS_SCROLLCOMBOBOX = &H5
Public Const DFCS_SCROLLSIZEGRIP = &H8
Public Const DFCS_SCROLLSIZEGRIPRIGHT = &H10

'// uType=DFC_BUTTON
Public Const DFCS_BUTTONCHECK = &H0
Public Const DFCS_BUTTONRADIOIMAGE = &H1
Public Const DFCS_BUTTONRADIOMASK = &H2
Public Const DFCS_BUTTONRADIO = &H4
Public Const DFCS_BUTTON3STATE = &H8
Public Const DFCS_BUTTONPUSH = &H10

'// OR flags for uState
Public Const DFCS_ADJUSTRECT = &H2000                    'Bounding rectangle is adjusted to exclude the surrounding edge of the push button.
Public Const DFCS_INACTIVE = &H100
Public Const DFCS_PUSHED = &H200
Public Const DFCS_CHECKED = &H400
Public Const DFCS_FLAT = &H4000
Public Const DFCS_MONO = &H8000
Public Const DFCS_TRANSPARENT = &H800   '// NT5/W98 only

'// used with GetStockObject()
Public Const GSO_WHITE_BRUSH = 0
Public Const GSO_LTGRAY_BRUSH = 1
Public Const GSO_GRAY_BRUSH = 2
Public Const GSO_DKGRAY_BRUSH = 3
Public Const GSO_BLACK_BRUSH = 4
Public Const GSO_NULL_BRUSH = 5
Public Const GSO_HOLLOW_BRUSH = GSO_NULL_BRUSH
Public Const GSO_WHITE_PEN = 6
Public Const GSO_BLACK_PEN = 7
Public Const GSO_NULL_PEN = 8
Public Const GSO_OEM_FIXED_FONT = 10
Public Const GSO_ANSI_FIXED_FONT = 11
Public Const GSO_ANSI_VAR_FONT = 12
Public Const GSO_SYSTEM_FONT = 13
Public Const GSO_DEVICE_DEFAULT_FONT = 14
Public Const GSO_DEFAULT_PALETTE = 15
Public Const GSO_SYSTEM_FIXED_FONT = 16
Public Const GSO_DEFAULT_FONT = 17

'// Owner draw control types
Public Const ODT_MENU = 1
Public Const ODT_LISTBOX = 2
Public Const ODT_COMBOBOX = 3
Public Const ODT_BUTTON = 4

'// Owner draw actions
Public Const ODA_DRAWENTIRE = &H1
Public Const ODA_SELECT = &H2
Public Const ODA_FOCUS = &H4

'// Owner draw state
Public Const ODS_SELECTED = &H1
Public Const ODS_GRAYED = &H2
Public Const ODS_DISABLED = &H4
Public Const ODS_CHECKED = &H8
Public Const ODS_FOCUS = &H10
Public Const ODS_DEFAULT = &H20
Public Const ODS_COMBOBOXEDIT = &H1000

'// GetSysColor() indexes
Public Const COLOR_SCROLLBAR = 0
Public Const COLOR_BACKGROUND = 1
Public Const COLOR_ACTIVECAPTION = 2
Public Const COLOR_INACTIVECAPTION = 3
Public Const COLOR_MENU = 4
Public Const COLOR_WINDOW = 5
Public Const COLOR_WINDOWFRAME = 6
Public Const COLOR_MENUTEXT = 7
Public Const COLOR_WINDOWTEXT = 8
Public Const COLOR_CAPTIONTEXT = 9
Public Const COLOR_ACTIVEBORDER = 10
Public Const COLOR_INACTIVEBORDER = 11
Public Const COLOR_APPWORKSPACE = 12
Public Const COLOR_HIGHLIGHT = 13
Public Const COLOR_HIGHLIGHTTEXT = 14
Public Const COLOR_BTNFACE = 15
Public Const COLOR_BTNSHADOW = 16
Public Const COLOR_GRAYTEXT = 17
Public Const COLOR_BTNTEXT = 18
Public Const COLOR_INACTIVECAPTIONTEXT = 19
Public Const COLOR_BTNHIGHLIGHT = 20
Public Const COLOR_3DDKSHADOW = 21
Public Const COLOR_3DLIGHT = 22
Public Const COLOR_INFOTEXT = 23
Public Const COLOR_INFOBK = 24


'// Flags for DrawState()
Public Const DST_COMPLEX = &H0
Public Const DST_TEXT = &H1
Public Const DST_PREFIXTEXT = &H2
Public Const DST_ICON = &H3
Public Const DST_BITMAP = &H4

'// State types for DrawState()
Public Const DSS_NORMAL = &H0
Public Const DSS_UNION = &H10                '// Gray string appearance
Public Const DSS_DISABLED = &H20
Public Const DSS_DEFAULT = &H40
Public Const DSS_MONO = &H80
Public Const DSS_RIGHT = &H8000

'// DrawEdge() constants
Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKENOUTER = &H2
Public Const BDR_RAISEDINNER = &H4
Public Const BDR_SUNKENINNER = &H8

Public Const BDR_OUTER = &H3
Public Const BDR_INNER = &HC
Public Const BDR_RAISED = &H5
Public Const BDR_SUNKEN = &HA

Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Public Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Public Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)

Public Const BF_LEFT = &H1
Public Const BF_TOP = &H2
Public Const BF_RIGHT = &H4
Public Const BF_BOTTOM = &H8

Public Const BF_TOPLEFT = (BF_TOP Or BF_LEFT)
Public Const BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)
Public Const BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
Public Const BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Public Const BF_DIAGONAL = &H10

Public Const BF_DIAGONAL_ENDTOPRIGHT = (BF_DIAGONAL Or BF_TOP Or BF_RIGHT)
Public Const BF_DIAGONAL_ENDTOPLEFT = (BF_DIAGONAL Or BF_TOP Or BF_LEFT)
Public Const BF_DIAGONAL_ENDBOTTOMLEFT = (BF_DIAGONAL Or BF_BOTTOM Or BF_LEFT)
Public Const BF_DIAGONAL_ENDBOTTOMRIGHT = (BF_DIAGONAL Or BF_BOTTOM Or BF_RIGHT)

Public Const BF_MIDDLE = &H800    ' Fill in the middle.
Public Const BF_SOFT = &H1000     ' Use for softer buttons.
Public Const BF_ADJUST = &H2000   ' Calculate the space left over.
Public Const BF_FLAT = &H4000     ' For flat rather than 3-D borders.
Public Const BF_MONO = &H8000     ' For monochrome borders.

'// Flags for DrawIconEx()
Public Const DI_MASK = &H1
Public Const DI_IMAGE = &H2
Public Const DI_NORMAL = &H3
Public Const DI_COMPAT = &H4
Public Const DI_DEFAULTSIZE = &H8

'// Constants for DrawAnimatedRects()
Public Const IDANI_OPEN = &H1
Public Const IDANI_CLOSE = &H2
Public Const IDANI_CAPTION = &H3


' SetWindowPos Flags
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOREDRAW = &H8
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_NOCOPYBITS = &H100
Public Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering

Public Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER

' SetWindowPos() hwndInsertAfter values
Public Const HWND_TOP = 0
Public Const HWND_BOTTOM = 1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

'// Used by Set and GetWindowPlacement
Public Const WPF_SETMINPOSITION = &H1
Public Const WPF_RESTORETOMAXIMIZED = &H2

'// ShowWindow() flags
Public Const SW_HIDE = 0
Public Const SW_SHOWNORMAL = 1
Public Const SW_NORMAL = 1
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_MAXIMIZE = 3
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOW = 5
Public Const SW_MINIMIZE = 6
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNA = 8
Public Const SW_RESTORE = 9
Public Const SW_SHOWDEFAULT = 10
Public Const SW_MAX = 10

' Window field offsets for GetWindowLong() and GetWindowWord()
Public Const GWL_WNDPROC = (-4)
Public Const GWL_HINSTANCE = (-6)
Public Const GWL_HWNDPARENT = (-8)
Public Const GWL_STYLE = (-16)
Public Const GWL_EXSTYLE = (-20)
Public Const GWL_USERDATA = (-21)
Public Const GWL_ID = (-12)

Public Const OUT_DEFAULT_PRECIS = 0
Public Const OUT_STRING_PRECIS = 1
Public Const OUT_CHARACTER_PRECIS = 2
Public Const OUT_STROKE_PRECIS = 3
Public Const OUT_TT_PRECIS = 4
Public Const OUT_DEVICE_PRECIS = 5
Public Const OUT_RASTER_PRECIS = 6
Public Const OUT_TT_ONLY_PRECIS = 7
Public Const OUT_OUTLINE_PRECIS = 8

Public Const CLIP_DEFAULT_PRECIS = 0
Public Const CLIP_CHARACTER_PRECIS = 1
Public Const CLIP_STROKE_PRECIS = 2
Public Const CLIP_MASK = &HF
Public Const CLIP_LH_ANGLES = 16
Public Const CLIP_TT_ALWAYS = 32
Public Const CLIP_EMBEDDED = 128

Public Const DEFAULT_QUALITY = 0
Public Const DRAFT_QUALITY = 1
Public Const PROOF_QUALITY = 2

Public Const DEFAULT_PITCH = 0
Public Const FIXED_PITCH = 1
Public Const VARIABLE_PITCH = 2

Public Const ANSI_CHARSET = 0
Public Const DEFAULT_CHARSET = 1
Public Const SYMBOL_CHARSET = 2
Public Const SHIFTJIS_CHARSET = 128
Public Const HANGEUL_CHARSET = 129
Public Const CHINESEBIG5_CHARSET = 136
Public Const OEM_CHARSET = 255

' Font Families
'
Public Const FF_DONTCARE = 0    '  Don't care or don't know.
Public Const FF_ROMAN = 16      '  Variable stroke width, serifed.

' Times Roman, Century Schoolbook, etc.
Public Const FF_SWISS = 32      '  Variable stroke width, sans-serifed.

' Helvetica, Swiss, etc.
Public Const FF_MODERN = 48     '  Constant stroke width, serifed or sans-serifed.

' Pica, Elite, Courier, etc.
Public Const FF_SCRIPT = 64     '  Cursive, etc.
Public Const FF_DECORATIVE = 80 '  Old English, etc.

'  EnumFonts Masks
Public Const RASTER_FONTTYPE = &H1
Public Const DEVICE_FONTTYPE = &H2
Public Const TRUETYPE_FONTTYPE = &H4

' Font Weights
Public Const FW_THIN = 100
Public Const FW_EXTRALIGHT = 200
Public Const FW_LIGHT = 300
Public Const FW_NORMAL = 400
Public Const FW_MEDIUM = 500
Public Const FW_SEMIBOLD = 600
Public Const FW_BOLD = 700
Public Const FW_EXTRABOLD = 800
Public Const FW_HEAVY = 900

Public Const FW_ULTRALIGHT = FW_EXTRALIGHT
Public Const FW_REGULAR = FW_NORMAL
Public Const FW_DEMIBOLD = FW_SEMIBOLD
Public Const FW_ULTRABOLD = FW_EXTRABOLD
Public Const FW_BLACK = FW_HEAVY

Public Const LOGPIXELSX = 88        '  Logical pixels/inch in X
Public Const LOGPIXELSY = 90        '  Logical pixels/inch in Y

'// WideChar/MB conversion constants
Public Const WC_DEFAULTCHECK = &H100       '  check for default char
Public Const WC_COMPOSITECHECK = &H200       '  convert composite to precomposed
Public Const WC_DISCARDNS = &H10        '  discard non-spacing chars
Public Const WC_SEPCHARS = &H20        '  generate separate chars
Public Const WC_DEFAULTCHAR = &H40        '  replace w/ default char
Public Const CP_ACP = 0  '  default to ANSI code page

'// Default window sizing
Public Const CW_USEDEFAULT = &H80000000

'// General window styles
Public Const WS_OVERLAPPED = &H0&
Public Const WS_POPUP = &H80000000
Public Const WS_CHILD = &H40000000
Public Const WS_MINIMIZE = &H20000000
Public Const WS_VISIBLE = &H10000000
Public Const WS_DISABLED = &H8000000
Public Const WS_CLIPSIBLINGS = &H4000000
Public Const WS_CLIPCHILDREN = &H2000000
Public Const WS_MAXIMIZE = &H1000000
Public Const WS_CAPTION = &HC00000                    '// WS_BORDER  Or  WS_DLGFRAME
Public Const WS_BORDER = &H800000
Public Const WS_DLGFRAME = &H400000
Public Const WS_VSCROLL = &H200000
Public Const WS_HSCROLL = &H100000
Public Const WS_SYSMENU = &H80000
Public Const WS_THICKFRAME = &H40000
Public Const WS_GROUP = &H20000
Public Const WS_TABSTOP = &H10000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_MAXIMIZEBOX = &H10000

Public Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or _
                                    WS_CAPTION Or _
                                    WS_SYSMENU Or _
                                    WS_THICKFRAME Or _
                                    WS_MINIMIZEBOX Or _
                                    WS_MAXIMIZEBOX)

Public Const WS_POPUPWINDOW = (WS_POPUP Or _
                               WS_BORDER Or _
                               WS_SYSMENU)

'//Extended Window Styles
Public Const WS_EX_DLGMODALFRAME = &H1&
Public Const WS_EX_NOPARENTNOTIFY = &H4&
Public Const WS_EX_TOPMOST = &H8&
Public Const WS_EX_ACCEPTFILES = &H10&
Public Const WS_EX_TRANSPARENT = &H20&
Public Const WS_EX_MDICHILD = &H40&
Public Const WS_EX_TOOLWINDOW = &H80&
Public Const WS_EX_WINDOWEDGE = &H100&
Public Const WS_EX_CLIENTEDGE = &H200&
Public Const WS_EX_CONTEXTHELP = &H400&
Public Const WS_EX_RIGHT = &H1000&
Public Const WS_EX_LEFT = &H0&
Public Const WS_EX_RTLREADING = &H2000&
Public Const WS_EX_LTRREADING = &H0&
Public Const WS_EX_LEFTSCROLLBAR = &H4000&
Public Const WS_EX_RIGHTSCROLLBAR = &H0&
Public Const WS_EX_CONTROLPARENT = &H10000
Public Const WS_EX_STATICEDGE = &H20000
Public Const WS_EX_APPWINDOW = &H40000
Public Const WS_EX_OVERLAPPEDWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_CLIENTEDGE)
Public Const WS_EX_PALETTEWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW Or WS_EX_TOPMOST)

'//Class styles
Public Const CS_VREDRAW = &H1
Public Const CS_HREDRAW = &H2
Public Const CS_DBLCLKS = &H8
Public Const CS_OWNDC = &H20
Public Const CS_CLASSDC = &H40
Public Const CS_PARENTDC = &H80
Public Const CS_NOCLOSE = &H200
Public Const CS_SAVEBITS = &H800
Public Const CS_BYTEALIGNCLIENT = &H1000
Public Const CS_BYTEALIGNWINDOW = &H2000
Public Const CS_GLOBALCLASS = &H4000


'// Listbox Styles
Public Const LBS_NOTIFY = &H1&
Public Const LBS_SORT = &H2&
Public Const LBS_NOREDRAW = &H4&
Public Const LBS_MULTIPLESEL = &H8&
Public Const LBS_OWNERDRAWFIXED = &H10&
Public Const LBS_OWNERDRAWVARIABLE = &H20&
Public Const LBS_HASSTRINGS = &H40&
Public Const LBS_USETABSTOPS = &H80&
Public Const LBS_NOINTEGRALHEIGHT = &H100&
Public Const LBS_MULTICOLUMN = &H200&
Public Const LBS_WANTKEYBOARDINPUT = &H400&
Public Const LBS_EXTENDEDSEL = &H800&
Public Const LBS_DISABLENOSCROLL = &H1000&
Public Const LBS_NODATA = &H2000&
Public Const LBS_NOSEL = &H4000&
Public Const LBS_STANDARD = (LBS_NOTIFY Or LBS_SORT Or WS_VSCROLL Or WS_BORDER)

'// Combo Box styles
Public Const CBS_SIMPLE = &H1&
Public Const CBS_DROPDOWN = &H2&
Public Const CBS_DROPDOWNLIST = &H3&
Public Const CBS_OWNERDRAWFIXED = &H10&
Public Const CBS_OWNERDRAWVARIABLE = &H20&
Public Const CBS_AUTOHSCROLL = &H40&
Public Const CBS_OEMCONVERT = &H80&
Public Const CBS_SORT = &H100&
Public Const CBS_HASSTRINGS = &H200&
Public Const CBS_NOINTEGRALHEIGHT = &H400&
Public Const CBS_DISABLENOSCROLL = &H800&
Public Const CBS_UPPERCASE = &H2000&
Public Const CBS_LOWERCASE = &H4000&

'// Edit Control Styles
Public Const ES_LEFT = &H0&
Public Const ES_CENTER = &H1&
Public Const ES_RIGHT = &H2&
Public Const ES_MULTILINE = &H4&
Public Const ES_UPPERCASE = &H8&
Public Const ES_LOWERCASE = &H10&
Public Const ES_PASSWORD = &H20&
Public Const ES_AUTOVSCROLL = &H40&
Public Const ES_AUTOHSCROLL = &H80&
Public Const ES_NOHIDESEL = &H100&
Public Const ES_OEMCONVERT = &H400&
Public Const ES_READONLY = &H800&
Public Const ES_WANTRETURN = &H1000&
Public Const ES_NUMBER = &H2000&

'// Static Control Constants
Public Const SS_LEFT = &H0&
Public Const SS_CENTER = &H1&
Public Const SS_RIGHT = &H2&
Public Const SS_ICON = &H3&
Public Const SS_BLACKRECT = &H4&
Public Const SS_GRAYRECT = &H5&
Public Const SS_WHITERECT = &H6&
Public Const SS_BLACKFRAME = &H7&
Public Const SS_GRAYFRAME = &H8&
Public Const SS_WHITEFRAME = &H9&
Public Const SS_USERITEM = &HA&
Public Const SS_SIMPLE = &HB&
Public Const SS_LEFTNOWORDWRAP = &HC&
Public Const SS_OWNERDRAW = &HD&
Public Const SS_BITMAP = &HE&
Public Const SS_ENHMETAFILE = &HF&
Public Const SS_ETCHEDHORZ = &H10&
Public Const SS_ETCHEDVERT = &H11&
Public Const SS_ETCHEDFRAME = &H12&
Public Const SS_TYPEMASK = &H1F&
Public Const SS_NOPREFIX = &H80&                '// Don't do "&" character translation
Public Const SS_NOTIFY = &H100&
Public Const SS_CENTERIMAGE = &H200&
Public Const SS_RIGHTJUST = &H400&
Public Const SS_REALSIZEIMAGE = &H800&
Public Const SS_SUNKEN = &H1000&
Public Const SS_ENDELLIPSIS = &H4000&
Public Const SS_PATHELLIPSIS = &H8000&
Public Const SS_WORDELLIPSIS = &HC000&
Public Const SS_ELLIPSISMASK = &HC000&

'// Button control styles
Public Const BS_PUSHBUTTON = &H0&
Public Const BS_DEFPUSHBUTTON = &H1&
Public Const BS_CHECKBOX = &H2&
Public Const BS_AUTOCHECKBOX = &H3&
Public Const BS_RADIOBUTTON = &H4&
Public Const BS_3STATE = &H5&
Public Const BS_AUTO3STATE = &H6&
Public Const BS_GROUPBOX = &H7&
Public Const BS_USERBUTTON = &H8&
Public Const BS_AUTORADIOBUTTON = &H9&
Public Const BS_OWNERDRAW = &HB&
Public Const BS_LEFTTEXT = &H20&
Public Const BS_TEXT = &H0&
Public Const BS_ICON = &H40&
Public Const BS_BITMAP = &H80&
Public Const BS_LEFT = &H100&
Public Const BS_RIGHT = &H200&
Public Const BS_CENTER = &H300&
Public Const BS_TOP = &H400&
Public Const BS_BOTTOM = &H800&
Public Const BS_VCENTER = &HC00&
Public Const BS_PUSHLIKE = &H1000&
Public Const BS_MULTILINE = &H2000&
Public Const BS_NOTIFY = &H4000&
Public Const BS_FLAT = &H8000&
Public Const BS_RIGHTBUTTON = BS_LEFTTEXT

'// Scroll Bar Styles
Public Const SBS_HORZ = &H0&
Public Const SBS_VERT = &H1&
Public Const SBS_TOPALIGN = &H2&
Public Const SBS_LEFTALIGN = &H2&
Public Const SBS_BOTTOMALIGN = &H4&
Public Const SBS_RIGHTALIGN = &H4&
Public Const SBS_SIZEBOXTOPLEFTALIGN = &H2&
Public Const SBS_SIZEBOXBOTTOMRIGHTALIGN = &H4&
Public Const SBS_SIZEBOX = &H8&
Public Const SBS_SIZEGRIP = &H10&

Public Const VK_SHIFT = &H10
Public Const VK_CONTROL = &H11
Public Const VK_MENU = &H12     '// Alt key

'// FormatMessage() flags
Public Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Public Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Public Const FORMAT_MESSAGE_FROM_STRING = &H400
Public Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Public Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Public Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
Public Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF

'  regular WinExec() codes
Public Const SE_ERR_FNF = 2                     '  file not found
Public Const SE_ERR_PNF = 3                     '  path not found
Public Const SE_ERR_ACCESSDENIED = 5            '  access denied
Public Const SE_ERR_OOM = 8                     '  out of memory
Public Const SE_ERR_DLLNOTFOUND = 32

'// Extended error results from ShellExecute(Ex)
Public Const SE_ERR_SHARE = 26
Public Const SE_ERR_ASSOCINCOMPLETE = 27
Public Const SE_ERR_DDETIMEOUT = 28
Public Const SE_ERR_DDEFAIL = 29
Public Const SE_ERR_DDEBUSY = 30
Public Const SE_ERR_NOASSOC = 31

' Scroll Bar Constants
Public Const SB_HORZ = 0
Public Const SB_VERT = 1
Public Const SB_CTL = 2
Public Const SB_BOTH = 3

'// General messages
Public Const WM_NULL = &H0
Public Const WM_CREATE = &H1
Public Const WM_DESTROY = &H2
Public Const WM_SETTEXT = &HC
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_QUIT = &H12
Public Const WM_SETCURSOR = &H20
Public Const WM_SETFONT = &H30
Public Const WM_GETFONT = &H31
Public Const WM_GETOBJECT = &H3D
Public Const WM_COPYDATA = &H4A
Public Const WM_NOTIFY = &H4E
Public Const WM_HELP = &H53
Public Const WM_NOTIFYFORMAT = &H55
    Public Const NFR_ANSI = 1
    Public Const NFR_UNICODE = 2
    Public Const NF_QUERY = 3
    Public Const NF_REQUERY = 4
Public Const WM_GETICON = &H7F
Public Const WM_SETICON = &H80
    Public Const ICON_SMALL = 0
    Public Const ICON_BIG = 1

Public Const WM_SYNCPAINT = &H88
Public Const WM_COMMAND = &H111
Public Const WM_TIMER = &H113
Public Const WM_PARENTNOTIFY = &H210
Public Const WM_DROPFILES = &H233
Public Const WM_APP = &H8000


'// WM_SYSCOMMAND and friends
Public Const WM_SYSCOMMAND = &H112
    '// Values for wParam in WM_SYSCOMMAND
    Public Const SC_SIZE = &HF000&
    Public Const SC_MOVE = &HF010&
    Public Const SC_MINIMIZE = &HF020&
    Public Const SC_MAXIMIZE = &HF030&
    Public Const SC_NEXTWINDOW = &HF040&
    Public Const SC_PREVWINDOW = &HF050&
    Public Const SC_CLOSE = &HF060&
    Public Const SC_VSCROLL = &HF070&
    Public Const SC_HSCROLL = &HF080&
    Public Const SC_MOUSEMENU = &HF090&
    Public Const SC_KEYMENU = &HF100&
    Public Const SC_ARRANGE = &HF110&
    Public Const SC_RESTORE = &HF120&
    Public Const SC_TASKLIST = &HF130&
    Public Const SC_SCREENSAVE = &HF140&
    Public Const SC_HOTKEY = &HF150&


Public Const WM_USER = &H400

'// Activation and focus
Public Const WM_ACTIVATE = &H6
    '//WM_ACTIVATE state values
    Public Const WA_INACTIVE = 0
    Public Const WA_ACTIVE = 1
    Public Const WA_CLICKACTIVE = 2

Public Const WM_SETFOCUS = &H7
Public Const WM_KILLFOCUS = &H8
Public Const WM_ENABLE = &HA
Public Const WM_CLOSE = &H10
Public Const WM_SHOWWINDOW = &H18
Public Const WM_ACTIVATEAPP = &H1C
Public Const WM_MOUSEACTIVATE = &H21
    '// WM_MOUSEACTIVEATE codes
    Public Const MA_ACTIVATE = 1
    Public Const MA_ACTIVATEANDEAT = 2
    Public Const MA_NOACTIVATE = 3
    Public Const MA_NOACTIVATEANDEAT = 4

Public Const WM_CHILDACTIVATE = &H22
Public Const WM_CAPTURECHANGED = &H215
Public Const WM_SETHOTKEY = &H32
Public Const WM_GETHOTKEY = &H33
Public Const WM_HOTKEY = &H312
    '// Keyboard modifiers for WM_HOTKEY
    Public Const MOD_ALT = &H1
    Public Const MOD_CONTROL = &H2
    Public Const MOD_SHIFT = &H4
    Public Const MOD_WIN = &H8

'// DDE messages
Public Const WM_DDE_FIRST = &H3E0
Public Const WM_DDE_INITIATE = (WM_DDE_FIRST)
Public Const WM_DDE_TERMINATE = (WM_DDE_FIRST + 1)
Public Const WM_DDE_ADVISE = (WM_DDE_FIRST + 2)
Public Const WM_DDE_UNADVISE = (WM_DDE_FIRST + 3)
Public Const WM_DDE_ACK = (WM_DDE_FIRST + 4)
Public Const WM_DDE_DATA = (WM_DDE_FIRST + 5)
Public Const WM_DDE_REQUEST = (WM_DDE_FIRST + 6)
Public Const WM_DDE_POKE = (WM_DDE_FIRST + 7)
Public Const WM_DDE_EXECUTE = (WM_DDE_FIRST + 8)
Public Const WM_DDE_LAST = (WM_DDE_FIRST + 8)

'// Printing
Public Const WM_SPOOLERSTATUS = &H2A
Public Const WM_PRINT = &H317
Public Const WM_PRINTCLIENT = &H318

'// Dialog-related messages
Public Const WM_INITDIALOG = &H110
Public Const WM_GETDLGCODE = &H87
Public Const WM_NEXTDLGCTL = &H28
Public Const WM_CANCELMODE = &H1F

'// SIzing and movement
Public Const WM_MOVE = &H3
Public Const WM_SIZE = &H5
    Public Const SIZE_RESTORED = 0
    Public Const SIZE_MINIMIZED = 1
    Public Const SIZE_MAXIMIZED = 2
    Public Const SIZE_MAXSHOW = 3
    Public Const SIZE_MAXHIDE = 4

Public Const WM_QUERYOPEN = &H13
Public Const WM_GETMINMAXINFO = &H24
Public Const WM_ENTERSIZEMOVE = &H231
Public Const WM_EXITSIZEMOVE = &H232
Public Const WM_WINDOWPOSCHANGING = &H46
Public Const WM_WINDOWPOSCHANGED = &H47

Public Const WM_SIZING = &H214
    Public Const WMSZ_LEFT = 1
    Public Const WMSZ_RIGHT = 2
    Public Const WMSZ_TOP = 3
    Public Const WMSZ_TOPLEFT = 4
    Public Const WMSZ_TOPRIGHT = 5
    Public Const WMSZ_BOTTOM = 6
    Public Const WMSZ_BOTTOMLEFT = 7
    Public Const WMSZ_BOTTOMRIGHT = 8

Public Const WM_MOVING = &H216

'// Journaling (CBT)
Public Const WM_QUEUESYNC = &H23
Public Const WM_CANCELJOURNAL = &H4B
Public Const WM_TCARD = &H52

'// Drawing and painting
Public Const WM_SETREDRAW = &HB
Public Const WM_PAINT = &HF
Public Const WM_COMPAREITEM = &H39
Public Const WM_QUERYDRAGICON = &H37
Public Const WM_ERASEBKGND = &H14
Public Const WM_PAINTICON = &H26
Public Const WM_ICONERASEBKGND = &H27
Public Const WM_DRAWITEM = &H2B
Public Const WM_MEASUREITEM = &H2C
Public Const WM_DELETEITEM = &H2D
Public Const WM_VKEYTOITEM = &H2E
Public Const WM_CHARTOITEM = &H2F

'// Environment messages
Public Const WM_COMPACTING = &H41
Public Const WM_COMMNOTIFY = &H44
Public Const WM_FONTCHANGE = &H1D
Public Const WM_TIMECHANGE = &H1E
Public Const WM_QUERYENDSESSION = &H11
Public Const WM_SYSCOLORCHANGE = &H15
Public Const WM_ENDSESSION = &H16
Public Const WM_WININICHANGE = &H1A
Public Const WM_SETTINGCHANGE = WM_WININICHANGE
Public Const WM_DEVMODECHANGE = &H1B
Public Const WM_DEVICECHANGE = &H219
Public Const WM_INPUTLANGCHANGEREQUEST = &H50
Public Const WM_INPUTLANGCHANGE = &H51
Public Const WM_USERCHANGED = &H54
Public Const WM_STYLECHANGING = &H7C                          '// This actually not an environmental thing
Public Const WM_STYLECHANGED = &H7D
Public Const WM_DISPLAYCHANGE = &H7E
Public Const WM_ENTERIDLE = &H121

'// Power-related messages
Public Const WM_POWER = &H48
    '// wParam for WM_POWER window message and DRV_POWER driver notification
    Public Const PWR_OK = 1
    Public Const PWR_FAIL = (-1)
    Public Const PWR_SUSPENDREQUEST = 1
    Public Const PWR_SUSPENDRESUME = 2
    Public Const PWR_CRITICALRESUME = 3

Public Const WM_POWERBROADCAST = &H218
    Public Const PBT_APMQUERYSUSPEND = &H0
    Public Const PBT_APMQUERYSTANDBY = &H1
    Public Const PBT_APMQUERYSUSPENDFAILED = &H2
    Public Const PBT_APMQUERYSTANDBYFAILED = &H3
    Public Const PBT_APMSUSPEND = &H4
    Public Const PBT_APMSTANDBY = &H5
    Public Const PBT_APMRESUMECRITICAL = &H6
    Public Const PBT_APMRESUMESUSPEND = &H7
    Public Const PBT_APMRESUMESTANDBY = &H8
    Public Const PBTF_APMRESUMEFROMFAILURE = &H1
    Public Const PBT_APMBATTERYLOW = &H9
    Public Const PBT_APMPOWERSTATUSCHANGE = &HA
    Public Const PBT_APMOEMEVENT = &HB
    Public Const PBT_APMRESUMEAUTOMATIC = &H12

'// Non-client messages
Public Const WM_NCCREATE = &H81
Public Const WM_NCDESTROY = &H82
Public Const WM_NCCALCSIZE = &H83
    Public Const WVR_ALIGNTOP = &H10
    Public Const WVR_ALIGNLEFT = &H20
    Public Const WVR_ALIGNBOTTOM = &H40
    Public Const WVR_ALIGNRIGHT = &H80
    Public Const WVR_HREDRAW = &H100
    Public Const WVR_VREDRAW = &H200
    Public Const WVR_REDRAW = (WVR_HREDRAW Or WVR_VREDRAW)
    Public Const WVR_VALIDRECTS = &H400

Public Const WM_NCHITTEST = &H84
    '// NC_HITTEST area codes
    Public Const HTERROR = (-2)
    Public Const HTTRANSPARENT = (-1)
    Public Const HTNOWHERE = 0
    Public Const HTCLIENT = 1
    Public Const HTCAPTION = 2
    Public Const HTSYSMENU = 3
    Public Const HTGROWBOX = 4
    Public Const HTSIZE = HTGROWBOX
    Public Const HTMENU = 5
    Public Const HTHSCROLL = 6
    Public Const HTVSCROLL = 7
    Public Const HTMINBUTTON = 8
    Public Const HTMAXBUTTON = 9
    Public Const HTLEFT = 10
    Public Const HTRIGHT = 11
    Public Const HTTOP = 12
    Public Const HTTOPLEFT = 13
    Public Const HTTOPRIGHT = 14
    Public Const HTBOTTOM = 15
    Public Const HTBOTTOMLEFT = 16
    Public Const HTBOTTOMRIGHT = 17
    Public Const HTBORDER = 18
    Public Const HTREDUCE = HTMINBUTTON
    Public Const HTZOOM = HTMAXBUTTON
    Public Const HTSIZEFIRST = HTLEFT
    Public Const HTSIZELAST = HTBOTTOMRIGHT
    Public Const HTOBJECT = 19
    Public Const HTCLOSE = 20
    Public Const HTHELP = 21

Public Const WM_NCPAINT = &H85
Public Const WM_NCACTIVATE = &H86
Public Const WM_NCMOUSEMOVE = &HA0
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const WM_NCLBUTTONUP = &HA2
Public Const WM_NCLBUTTONDBLCLK = &HA3
Public Const WM_NCRBUTTONDOWN = &HA4
Public Const WM_NCRBUTTONUP = &HA5
Public Const WM_NCRBUTTONDBLCLK = &HA6
Public Const WM_NCMBUTTONDOWN = &HA7
Public Const WM_NCMBUTTONUP = &HA8
Public Const WM_NCMBUTTONDBLCLK = &HA9

'// IME messages
Public Const WM_IME_STARTCOMPOSITION = &H10D
Public Const WM_IME_ENDCOMPOSITION = &H10E
Public Const WM_IME_COMPOSITION = &H10F
Public Const WM_IME_KEYLAST = &H10F
Public Const WM_IME_SETCONTEXT = &H281
Public Const WM_IME_NOTIFY = &H282
Public Const WM_IME_CONTROL = &H283
Public Const WM_IME_COMPOSITIONFULL = &H284
Public Const WM_IME_SELECT = &H285
Public Const WM_IME_CHAR = &H286
Public Const WM_IME_KEYDOWN = &H290
Public Const WM_IME_KEYUP = &H291

'// Keyboard messages
Public Const WM_KEYFIRST = &H100
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_CHAR = &H102
Public Const WM_DEADCHAR = &H103
Public Const WM_SYSKEYDOWN = &H104
Public Const WM_SYSKEYUP = &H105
Public Const WM_SYSCHAR = &H106
Public Const WM_SYSDEADCHAR = &H107
Public Const WM_KEYLAST = &H108

'// Scrolling
Public Const WM_HSCROLL = &H114
Public Const WM_VSCROLL = &H115

'// Menu messages
Public Const WM_INITMENU = &H116
Public Const WM_INITMENUPOPUP = &H117
Public Const WM_MENUSELECT = &H11F
Public Const WM_MENUCHAR = &H120
Public Const WM_MENURBUTTONUP = &H122
Public Const WM_MENUDRAG = &H123
Public Const WM_MENUGETOBJECT = &H124
Public Const WM_UNINITMENUPOPUP = &H125
Public Const WM_MENUCOMMAND = &H126
Public Const WM_ENTERMENULOOP = &H211
Public Const WM_EXITMENULOOP = &H212
Public Const WM_NEXTMENU = &H213
Public Const WM_CONTEXTMENU = &H7B

'// WM_CTLCOLOR* messages
Public Const WM_CTLCOLORMSGBOX = &H132
Public Const WM_CTLCOLOREDIT = &H133
Public Const WM_CTLCOLORLISTBOX = &H134
Public Const WM_CTLCOLORBTN = &H135
Public Const WM_CTLCOLORDLG = &H136
Public Const WM_CTLCOLORSCROLLBAR = &H137
Public Const WM_CTLCOLORSTATIC = &H138

'// MDI-related messages
Public Const WM_MDICREATE = &H220
Public Const WM_MDIDESTROY = &H221
Public Const WM_MDIACTIVATE = &H222
Public Const WM_MDIRESTORE = &H223
Public Const WM_MDINEXT = &H224
Public Const WM_MDIMAXIMIZE = &H225
Public Const WM_MDITILE = &H226
Public Const WM_MDICASCADE = &H227
Public Const WM_MDIICONARRANGE = &H228
Public Const WM_MDIGETACTIVE = &H229
Public Const WM_MDISETMENU = &H230
Public Const WM_MDIREFRESHMENU = &H234

'// Mouse messages
Public Const WM_MOUSEFIRST = &H200
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208
Public Const WM_MBUTTONDBLCLK = &H209
Public Const WM_MOUSEWHEEL = &H20A
Public Const WM_MOUSELAST = &H20A
Public Const WM_MOUSEHOVER = &H2A1   '// WINVER >= 2K or COMMCTL >=4.71
Public Const WM_MOUSELEAVE = &H2A3   '// WINVER >= 2K or COMMCTL >=4.71
    '//Key State Masks for Mouse Messages
    Public Const MK_LBUTTON = &H1
    Public Const MK_RBUTTON = &H2
    Public Const MK_SHIFT = &H4
    Public Const MK_CONTROL = &H8
    Public Const MK_MBUTTON = &H10

'// Edit messages (as in general editing, not the control)
Public Const WM_CUT = &H300
Public Const WM_COPY = &H301
Public Const WM_PASTE = &H302
Public Const WM_CLEAR = &H303
Public Const WM_UNDO = &H304

'// Clipboard messages
Public Const WM_RENDERFORMAT = &H305
Public Const WM_RENDERALLFORMATS = &H306
Public Const WM_DESTROYCLIPBOARD = &H307
Public Const WM_DRAWCLIPBOARD = &H308
Public Const WM_PAINTCLIPBOARD = &H309
Public Const WM_VSCROLLCLIPBOARD = &H30A
Public Const WM_SIZECLIPBOARD = &H30B
Public Const WM_ASKCBFORMATNAME = &H30C
Public Const WM_CHANGECBCHAIN = &H30D
Public Const WM_HSCROLLCLIPBOARD = &H30E

'// Palette messages
Public Const WM_QUERYNEWPALETTE = &H30F
Public Const WM_PALETTEISCHANGING = &H310
Public Const WM_PALETTECHANGED = &H311

'// Begin Control messages

'// Listbox Return Values
Public Const LB_OKAY = 0
Public Const LB_ERR = (-1)
Public Const LB_ERRSPACE = (-2)

'// Listbox Notification Codes
Public Const LBN_ERRSPACE = (-2)
Public Const LBN_SELCHANGE = 1
Public Const LBN_DBLCLK = 2
Public Const LBN_SELCANCEL = 3
Public Const LBN_SETFOCUS = 4
Public Const LBN_KILLFOCUS = 5

'// Listbox messages
Public Const LB_ADDSTRING = &H180
Public Const LB_INSERTSTRING = &H181
Public Const LB_DELETESTRING = &H182
Public Const LB_SELITEMRANGEEX = &H183
Public Const LB_RESETCONTENT = &H184
Public Const LB_SETSEL = &H185
Public Const LB_SETCURSEL = &H186
Public Const LB_GETSEL = &H187
Public Const LB_GETCURSEL = &H188
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_GETCOUNT = &H18B
Public Const LB_SELECTSTRING = &H18C
Public Const LB_DIR = &H18D
Public Const LB_GETTOPINDEX = &H18E
Public Const LB_FINDSTRING = &H18F
Public Const LB_GETSELCOUNT = &H190
Public Const LB_GETSELITEMS = &H191
Public Const LB_SETTABSTOPS = &H192
Public Const LB_GETHORIZONTALEXTENT = &H193
Public Const LB_SETHORIZONTALEXTENT = &H194
Public Const LB_SETCOLUMNWIDTH = &H195
Public Const LB_ADDFILE = &H196
Public Const LB_SETTOPINDEX = &H197
Public Const LB_GETITEMRECT = &H198
Public Const LB_GETITEMDATA = &H199
Public Const LB_SETITEMDATA = &H19A
Public Const LB_SELITEMRANGE = &H19B
Public Const LB_SETANCHORINDEX = &H19C
Public Const LB_GETANCHORINDEX = &H19D
Public Const LB_SETCARETINDEX = &H19E
Public Const LB_GETCARETINDEX = &H19F
Public Const LB_SETITEMHEIGHT = &H1A0
Public Const LB_GETITEMHEIGHT = &H1A1
Public Const LB_FINDSTRINGEXACT = &H1A2
Public Const LB_SETLOCALE = &H1A5
Public Const LB_GETLOCALE = &H1A6
Public Const LB_SETCOUNT = &H1A7
Public Const LB_INITSTORAGE = &H1A8
Public Const LB_ITEMFROMPOINT = &H1A9
Public Const LB_MSGMAX = &H1B0

'// DlgDirList constants
Public Const DDL_READWRITE = &H0
Public Const DDL_READONLY = &H1
Public Const DDL_HIDDEN = &H2
Public Const DDL_SYSTEM = &H4
Public Const DDL_DIRECTORY = &H10
Public Const DDL_ARCHIVE = &H20

Public Const DDL_POSTMSGS = &H2000
Public Const DDL_DRIVES = &H4000
Public Const DDL_EXCLUSIVE = &H8000



'// Combo Box return Values
Public Const CB_OKAY = 0
Public Const CB_ERR = (-1)
Public Const CB_ERRSPACE = (-2)

'// Combo Box Notification Codes
Public Const CBN_ERRSPACE = (-1)
Public Const CBN_SELCHANGE = 1
Public Const CBN_DBLCLK = 2
Public Const CBN_SETFOCUS = 3
Public Const CBN_KILLFOCUS = 4
Public Const CBN_EDITCHANGE = 5
Public Const CBN_EDITUPDATE = 6
Public Const CBN_DROPDOWN = 7
Public Const CBN_CLOSEUP = 8
Public Const CBN_SELENDOK = 9
Public Const CBN_SELENDCANCEL = 10

'//Combo Box messages
Public Const CB_GETEDITSEL = &H140
Public Const CB_LIMITTEXT = &H141
Public Const CB_SETEDITSEL = &H142
Public Const CB_ADDSTRING = &H143
Public Const CB_DELETESTRING = &H144
Public Const CB_DIR = &H145
Public Const CB_GETCOUNT = &H146
Public Const CB_GETCURSEL = &H147
Public Const CB_GETLBTEXT = &H148
Public Const CB_GETLBTEXTLEN = &H149
Public Const CB_INSERTSTRING = &H14A
Public Const CB_RESETCONTENT = &H14B
Public Const CB_FINDSTRING = &H14C
Public Const CB_SELECTSTRING = &H14D
Public Const CB_SETCURSEL = &H14E
Public Const CB_SHOWDROPDOWN = &H14F
Public Const CB_GETITEMDATA = &H150
Public Const CB_SETITEMDATA = &H151
Public Const CB_GETDROPPEDCONTROLRECT = &H152
Public Const CB_SETITEMHEIGHT = &H153
Public Const CB_GETITEMHEIGHT = &H154
Public Const CB_SETEXTENDEDUI = &H155
Public Const CB_GETEXTENDEDUI = &H156
Public Const CB_GETDROPPEDSTATE = &H157
Public Const CB_FINDSTRINGEXACT = &H158
Public Const CB_SETLOCALE = &H159
Public Const CB_GETLOCALE = &H15A
Public Const CB_GETTOPINDEX = &H15B
Public Const CB_SETTOPINDEX = &H15C
Public Const CB_GETHORIZONTALEXTENT = &H15D
Public Const CB_SETHORIZONTALEXTENT = &H15E
Public Const CB_GETDROPPEDWIDTH = &H15F
Public Const CB_SETDROPPEDWIDTH = &H160
Public Const CB_INITSTORAGE = &H161
Public Const CB_MSGMAX = &H162

'// Edit Control Notification Codes
Public Const EN_SETFOCUS = &H100
Public Const EN_KILLFOCUS = &H200
Public Const EN_CHANGE = &H300
Public Const EN_UPDATE = &H400
Public Const EN_ERRSPACE = &H500
Public Const EN_MAXTEXT = &H501
Public Const EN_HSCROLL = &H601
Public Const EN_VSCROLL = &H602

'// Edit Control Messages
Public Const EM_GETSEL = &HB0
Public Const EM_SETSEL = &HB1
Public Const EM_GETRECT = &HB2
Public Const EM_SETRECT = &HB3
Public Const EM_SETRECTNP = &HB4
Public Const EM_SCROLL = &HB5
Public Const EM_LINESCROLL = &HB6
Public Const EM_SCROLLCARET = &HB7
Public Const EM_GETMODIFY = &HB8
Public Const EM_SETMODIFY = &HB9
Public Const EM_GETLINECOUNT = &HBA
Public Const EM_LINEINDEX = &HBB
Public Const EM_SETHANDLE = &HBC
Public Const EM_GETHANDLE = &HBD
Public Const EM_GETTHUMB = &HBE
Public Const EM_LINELENGTH = &HC1
Public Const EM_REPLACESEL = &HC2
Public Const EM_GETLINE = &HC4
Public Const EM_LIMITTEXT = &HC5
Public Const EM_CANUNDO = &HC6
Public Const EM_UNDO = &HC7
Public Const EM_FMTLINES = &HC8
Public Const EM_LINEFROMCHAR = &HC9
Public Const EM_SETTABSTOPS = &HCB
Public Const EM_SETPASSWORDCHAR = &HCC
Public Const EM_EMPTYUNDOBUFFER = &HCD
Public Const EM_GETFIRSTVISIBLELINE = &HCE
Public Const EM_SETREADONLY = &HCF
Public Const EM_SETWORDBREAKPROC = &HD0
Public Const EM_GETWORDBREAKPROC = &HD1
Public Const EM_GETPASSWORDCHAR = &HD2
Public Const EM_SETMARGINS = &HD3
    '// Edit control EM_SETMARGIN parameters
    Public Const EC_LEFTMARGIN = &H1
    Public Const EC_RIGHTMARGIN = &H2
    Public Const EC_USEFONTINFO = &HFFFF
Public Const EM_GETMARGINS = &HD4
Public Const EM_SETLIMITTEXT = EM_LIMITTEXT          '// win40 Name change
Public Const EM_GETLIMITTEXT = &HD5
Public Const EM_POSFROMCHAR = &HD6
Public Const EM_CHARFROMPOS = &HD7

'// Static control notification messages
Public Const STN_CLICKED = 0
Public Const STN_DBLCLK = 1
Public Const STN_ENABLE = 2
Public Const STN_DISABLE = 3

'// Static Control Mesages
Public Const STM_SETICON = &H170
Public Const STM_GETICON = &H171
Public Const STM_SETIMAGE = &H172
Public Const STM_GETIMAGE = &H173
Public Const STM_MSGMAX = &H174

'// User Button Notification Codes
Public Const BN_CLICKED = 0
Public Const BN_PAINT = 1
Public Const BN_HILITE = 2
Public Const BN_UNHILITE = 3
Public Const BN_DISABLE = 4
Public Const BN_DOUBLECLICKED = 5
Public Const BN_PUSHED = BN_HILITE
Public Const BN_UNPUSHED = BN_UNHILITE
Public Const BN_DBLCLK = BN_DOUBLECLICKED
Public Const BN_SETFOCUS = 6
Public Const BN_KILLFOCUS = 7

'// Button Control Messages
Public Const BM_GETCHECK = &HF0
Public Const BM_SETCHECK = &HF1
Public Const BM_GETSTATE = &HF2
Public Const BM_SETSTATE = &HF3
Public Const BM_SETSTYLE = &HF4
Public Const BM_CLICK = &HF5
Public Const BM_GETIMAGE = &HF6
Public Const BM_SETIMAGE = &HF7
Public Const BST_UNCHECKED = &H0
Public Const BST_CHECKED = &H1
Public Const BST_INDETERMINATE = &H2
Public Const BST_PUSHED = &H4
Public Const BST_FOCUS = &H8

'// Scroll bar messages
Public Const SBM_SETPOS = &HE0                    '// not in win3.1 */
Public Const SBM_GETPOS = &HE1                    '// not in win3.1 */
Public Const SBM_SETRANGE = &HE2                  '// not in win3.1 */
Public Const SBM_SETRANGEREDRAW = &HE6            '// not in win3.1 */
Public Const SBM_GETRANGE = &HE3                  '// not in win3.1 */
Public Const SBM_ENABLE_ARROWS = &HE4             '// not in win3.1 */
Public Const SBM_SETSCROLLINFO = &HE9
Public Const SBM_GETSCROLLINFO = &HEA
    Public Const SIF_RANGE = &H1
    Public Const SIF_PAGE = &H2
    Public Const SIF_POS = &H4
    Public Const SIF_DISABLENOSCROLL = &H8
    Public Const SIF_TRACKPOS = &H10
    Public Const SIF_ALL = (SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS)

    '// Scrollbar commands
    Public Const SB_LINEUP = 0
    Public Const SB_LINELEFT = 0
    Public Const SB_LINEDOWN = 1
    Public Const SB_LINERIGHT = 1
    Public Const SB_PAGEUP = 2
    Public Const SB_PAGELEFT = 2
    Public Const SB_PAGEDOWN = 3
    Public Const SB_PAGERIGHT = 3
    Public Const SB_THUMBPOSITION = 4
    Public Const SB_THUMBTRACK = 5
    Public Const SB_TOP = 6
    Public Const SB_LEFT = 6
    Public Const SB_BOTTOM = 7
    Public Const SB_RIGHT = 7
    Public Const SB_ENDSCROLL = 8


'// EnableScrollBar() flags
Public Const ESB_ENABLE_BOTH = &H0
Public Const ESB_DISABLE_BOTH = &H3
Public Const ESB_DISABLE_LEFT = &H1
Public Const ESB_DISABLE_RIGHT = &H2
Public Const ESB_DISABLE_UP = &H1
Public Const ESB_DISABLE_DOWN = &H2
Public Const ESB_DISABLE_LTUP = ESB_DISABLE_LEFT
Public Const ESB_DISABLE_RTDN = ESB_DISABLE_RIGHT

'
'   Returns an ANSI or Unicode pointer to a string.
'   Note that StrPtr() is equivalent to the wide
'   version of CharNext
'
Public Function StrToPtr(sz As String) As Long

    If Len(sz) Then
#If UNICODE Then
        StrToPtr = StrPtr(sz)
#Else
        StrToPtr = CharNext(sz) - 1
#End If
    End If

End Function

'
'   Dereferences an ANSI or Unicode string pointer
'   and returns a normal VB BSTR
'
Public Function PtrToStr(ByVal lpsz As Long) As String

    Dim sOut As String
    Dim lLen As Long
    
    lLen = lstrlen(lpsz)
    
    If (lLen > 0) Then
        sOut = String$(lLen, vbNullChar)
        Call CopyMemory(ByVal sOut, ByVal lpsz, lLen)
        PtrToStr = sOut
    End If

End Function

Public Function HIBYTE(ByVal wParam As Integer) As Integer
    HIBYTE = (Abs(wParam) \ &H100) And &HFF&
End Function

Public Function LOBYTE(ByVal wParam As Integer) As Integer
    LOBYTE = Abs(wParam) And &HFF&
End Function

Public Function MAKEWORD(ByVal wLow As Integer, ByVal wHigh As Integer) As Integer
    If wHigh And &H80 Then
        MAKEWORD = (((wHigh And &H7F) * 256) + wLow) Or &H8000
    Else
        MAKEWORD = (wHigh * 256) + wLow
    End If
    'MAKEWORD = wLow Or (&H8000 * wHigh)
    '#define MAKEWORD ((WORD) (((BYTE) (a)) | ((WORD) ((BYTE) (b))) << 8))
End Function


Public Function Max(ByVal param1 As Long, ByVal param2 As Long) As Long
    If param1 > param2 Then Max = param1 Else Max = param2
End Function

Public Function Min(ByVal param1 As Long, ByVal param2 As Long) As Long
    If param1 < param2 Then Min = param1 Else Min = param2
End Function

Public Function HIWORD(ByVal dwValue As Long) As Long
    CopyMemory HIWORD, ByVal VarPtr(dwValue) + 2, 2
End Function
  
Public Function LOWORD(ByVal dwValue As Long) As Long
    CopyMemory LOWORD, dwValue, 2
End Function

Public Function MAKELONG(ByVal wLow As Long, ByVal wHi As Long) As Long

    If (wHi And &H8000&) Then
        MAKELONG = (((wHi And &H7FFF&) * 65536) Or (wLow And &HFFFF&)) Or &H80000000
    Else
        MAKELONG = LOWORD(wLow) Or (&H10000 * LOWORD(wHi))
        'MAKELONG = ((wHi * 65535) + wLow)
    End If

End Function

Public Function MAKEINTRESOURCE(ByVal lID As Long) As String

    MAKEINTRESOURCE = "#" & CStr(MAKELONG(lID, 0))

End Function

'
'   Workaround for the unary AddressOf operator
'
Public Function GetProcAddress(ByVal hProc As Long) As Long
    
    GetProcAddress = hProc

End Function


#End If


