Attribute VB_Name = "modCBSHBrowseForFolder"
'--------------------------------------------------------------------------------------------
'   Name:           modCBSHBrowseForFolder (MODCBSHBROWSEFORFOLDER.BAS)
'   Type:           Declarations Module
'   Description:    Declarations and defines for the SHBrowseForFolder API.
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
'       (cls)       IBFFCallback, if CB_BFF_CALLBACK is #defined
'
'   Last Revision:  04/07/01
'   Revision:       1
'
'--------------------------------------------------------------------------------------------
Option Explicit
DefLng A-Z

'// NOTE: YOu'll need to define CB_BFF_CALLBACK at the project level
'// in order to do callbacks via the IBFFCallback class.

'//
'// This is a prototype of what a BFFCALLBACK procedure should look like.
'// The name itself does not matter.
'//
'// Public Function BFFCallback(ByVal hwnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long

Private Const WM_USER = &H400

Type BROWSEINFO
    hwndOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long             '// Flags that control the return stuff
    lpfn As Long
    lParam As Long              '// extra info that's passed back in callbacks
    iImage As Long              '// output var: where to return the Image index.
End Type

Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Public Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long


'// Browsing for directory.
Public Const BIF_RETURNONLYFSDIRS = &H1         '// For finding a folder to start document searching
Public Const BIF_DONTGOBELOWDOMAIN = &H2        '// For starting the Find Computer
Public Const BIF_STATUSTEXT = &H4               '// Top of the dialog has 2 lines of text for BROWSEINFO.lpszTitle and one line if
                                                '// this flag is set.  Passing the message BFFM_SETSTATUSTEXTA to the hwnd can set the
                                                '// rest of the text.  This is not used with BIF_USENEWUI and BROWSEINFO.lpszTitle gets
                                                '// all three lines of text.
Public Const BIF_RETURNFSANCESTORS = &H8
Public Const BIF_EDITBOX = &H10                 '// Add an editbox to the dialog
Public Const BIF_VALIDATE = &H20                '// insist on valid result (or CANCEL)

Public Const BIF_NEWDIALOGSTYLE = &H40          '// Use the new dialog layout with the ability to resize
                                                '// Caller needs to call OleInitialize() before using this API

Public Const BIF_USENEWUI = (BIF_NEWDIALOGSTYLE Or BIF_EDITBOX)

Public Const BIF_BROWSEINCLUDEURLS = &H80       '// Allow URLs to be displayed or entered. (Requires BIF_USENEWUI)

Public Const BIF_BROWSEFORCOMPUTER = &H1000     '// Browsing for Computers.
Public Const BIF_BROWSEFORPRINTER = &H2000      '// Browsing for Printers
Public Const BIF_BROWSEINCLUDEFILES = &H4000    '// Browsing for Everything
Public Const BIF_SHAREABLE = &H8000             '// sharable resources displayed (remote shares, requires BIF_USENEWUI)

'// message from browser
Public Const BFFM_INITIALIZED = 1
Public Const BFFM_SELCHANGED = 2
Public Const BFFM_VALIDATEFAILED = 3           '// lParam:szPath ret:1(cont),0(EndDialog)

'// messages to browser
Public Const BFFM_SETSTATUSTEXT = (WM_USER + 100)
Public Const BFFM_ENABLEOK = (WM_USER + 101)
Public Const BFFM_SETSELECTION = (WM_USER + 102)
'   A callback that dereferences the lpData arg to a IBFFCallback
'   object and delegates handling to it.
'
Public Function BFFCallback(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long

    Dim hbff As IBFFCallback
    
    If (lpData <> 0) Then
        Call CopyMemory(hbff, lpData, 4)
        If Not hbff Is Nothing Then
            Call hbff.Message(hWnd, uMsg, lParam, lpData)
            Call CopyMemory(hbff, 0&, 4)
        End If
    End If

End Function
'   A simple callback. The wrapper class allocates memory via HeapAlloc()
'   and stores a pointer to a string on lpData, which is then dereferenced
'   here and set to be the initially selected path on the dialog.
'
Public Function BFFCallbackSimple(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long

    If uMsg = BFFM_INITIALIZED Then
    
        If (lpData <> 0) Then
            If (lstrlenA(lpData) <> 0) Then
                Call SendMessage(hWnd, BFFM_SETSELECTION, 1, ByVal lpData)
            End If
        End If
    End If
    
End Function
'   Used to set the dialog's status text from the BFF callback
'
Public Sub BFFSetStatusText(ByVal hWndDialog As Long, ByVal Status As String)

    Call SendMessage(hWndDialog, BFFM_SETSTATUSTEXT, 0, ByVal Status)

End Sub
'   Used to enable or disable the dialog's OK button from the BFF callback
'
Public Sub BFFEnableOKButton(ByVal hWndDialog As Long, ByVal Enable As Boolean)

    Call SendMessage(hWndDialog, BFFM_ENABLEOK, 0, ByVal Abs(Enable))

End Sub
'   Use to set the dialog's selected path from the BFF callback
'   This method uses the standard path, but a PIDL can be used
'   as well.
'
Public Sub BFFSetPath(ByVal hWndDialog As Long, ByVal Path As String)

    Call SendMessage(hWndDialog, BFFM_SETSELECTION, 1, ByVal Path)

End Sub
