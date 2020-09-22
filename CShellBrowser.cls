VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CShellBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------------------
'   Name:           CShellBrowser (CSHELLBROWSER.CLS)
'   Type:           Utility class
'   Description:    Encapsulates the SHBrowseForFolder API
'
'   Author:         Klaus H. Probst [kprobst@vbbox.com]
'   URL:            http://www.vbbox.com/
'   Copyright:      This work is copyright © 1998-2001, Klaus H. Probst
'   Usage:          You may use this code as you see fit, provided that you assume all
'                   responsibilities for doing so.
'   Distribution:   If you intend to distribute the file(s) that make up this sample to
'                   any WWW site, online service, electronic bulletin board system (BBS),
'                   CD or any other electronic or physical media, you must notify me in
'                   advance to obtain my express permission.
'
'
'   Notes:
'
'   Dependencies:
'
'       (sub)       modCBShBrowseForFolder, modCBDeclares
'
'--------------------------------------------------------------------------------------------
Option Explicit
DefLng A-Z

#If Not CB_NO_LOCAL_SYMBOLS Then

Private Const WM_USER = &H400

Private Type BROWSEINFO
    hwndOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long             '// Flags that control the return stuff
    lpfn As Long
    lParam As Long              '// extra info that's passed back in callbacks
    iImage As Long              '// output var: where to return the Image index.
End Type

Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long

Private Declare Function GetProcessHeap Lib "kernel32" () As Long
Private Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long


'// Browsing for directory.
Private Const BIF_RETURNONLYFSDIRS = &H1         '// For finding a folder to start document searching
Private Const BIF_DONTGOBELOWDOMAIN = &H2        '// For starting the Find Computer
Private Const BIF_STATUSTEXT = &H4               '// Top of the dialog has 2 lines of text for BROWSEINFO.lpszTitle and one line if
                                                '// this flag is set.  Passing the message BFFM_SETSTATUSTEXTA to the hwnd can set the
                                                '// rest of the text.  This is not used with BIF_USENEWUI and BROWSEINFO.lpszTitle gets
                                                '// all three lines of text.
Private Const BIF_RETURNFSANCESTORS = &H8
Private Const BIF_EDITBOX = &H10                 '// Add an editbox to the dialog
Private Const BIF_VALIDATE = &H20                '// insist on valid result (or CANCEL)

Private Const BIF_NEWDIALOGSTYLE = &H40          '// Use the new dialog layout with the ability to resize
                                                '// Caller needs to call OleInitialize() before using this API

Private Const BIF_USENEWUI = (BIF_NEWDIALOGSTYLE Or BIF_EDITBOX)

Private Const BIF_BROWSEINCLUDEURLS = &H80       '// Allow URLs to be displayed or entered. (Requires BIF_USENEWUI)

Private Const BIF_BROWSEFORCOMPUTER = &H1000     '// Browsing for Computers.
Private Const BIF_BROWSEFORPRINTER = &H2000      '// Browsing for Printers
Private Const BIF_BROWSEINCLUDEFILES = &H4000    '// Browsing for Everything
Private Const BIF_SHAREABLE = &H8000             '// sharable resources displayed (remote shares, requires BIF_USENEWUI)

'// message from browser
Private Const BFFM_INITIALIZED = 1
Private Const BFFM_SELCHANGED = 2
Private Const BFFM_VALIDATEFAILED = 3           '// lParam:szPath ret:1(cont),0(EndDialog)

'// messages to browser
Private Const BFFM_SETSTATUSTEXT = (WM_USER + 100)
Private Const BFFM_ENABLEOK = (WM_USER + 101)
Private Const BFFM_SETSELECTION = (WM_USER + 102)

#End If

Public Enum BIFFlags
    bifReturnOnlyFSDirs = BIF_RETURNONLYFSDIRS
    bifDontGoBelowDomain = BIF_DONTGOBELOWDOMAIN
    bifStatusText = BIF_STATUSTEXT
    bifReturnFSAncestors = BIF_RETURNFSANCESTORS
    bifEditBox = BIF_EDITBOX
    bifValidate = BIF_VALIDATE
    bifNewDialogStyle = BIF_NEWDIALOGSTYLE
    bifUseNewUI = BIF_USENEWUI
    bifBrowseIncludeURLs = BIF_BROWSEINCLUDEURLS
    bifBrowseForComputer = BIF_BROWSEFORCOMPUTER
    bifBrowseForPrinter = BIF_BROWSEFORPRINTER
    bifBrowseIncludeFiles = BIF_BROWSEINCLUDEFILES
    bifShareable = BIF_SHAREABLE
End Enum
Friend Function BrowseForFolder(ByRef CallbackHandler As IBFFCallback, _
                                Optional ByVal hWndParent As Long = 0, _
                                Optional ByVal Flags As BIFFlags = bifReturnOnlyFSDirs, _
                                Optional ByVal sTitle As String = "") As String
    Dim pidlReturn As Long
    Dim lReturn As Long
    Dim sBuffer As String
    Dim ubi As BROWSEINFO

    ubi.ulFlags = Flags
    ubi.hwndOwner = hWndParent
    ubi.lpszTitle = sTitle
    
    '// This is a lot simpler because all handling is delegated
    '// to the callback implementation.
    ubi.lParam = ObjPtr(CallbackHandler)
    ubi.lpfn = GetProcAddress(AddressOf BFFCallback)
    
    pidlReturn = SHBrowseForFolder(ubi)
    
    If (pidlReturn <> 0) Then
        sBuffer = String$(256, 0)      '// MAX_PATH
        lReturn = SHGetPathFromIDList(pidlReturn, sBuffer)
        If lReturn <> 0 Then
            BrowseForFolder = Left$(sBuffer, InStr(sBuffer, Chr$(0)) - 1)
        
        End If
        Call CoTaskMemFree(pidlReturn)
    End If
End Function
