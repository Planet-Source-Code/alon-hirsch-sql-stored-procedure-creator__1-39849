VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGlobal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ***********************************************************************
'
' CLASS : clsGlobals.cls
'
' PURPOSE : Provide access to global functions / variables without using
'           real global variables. All variables are properties of this
'           class.
'
' WRITTEN BY : Alon Hirsch
'
' COMPANY : Debtpack (Pty) Ltd. - Development
'
' DATE : 12 February 2001
'
' ***********************************************************************
Option Explicit
DefInt A-Z

' ***** PRIVATE VARIABLES *****
' poperties etc.
Private m_sVersion As String

' declares for Windows API's
Private Declare Function SendMessageBynum& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function LockWindowUpdate& Lib "user32" (ByVal hWndLock As Long)
Private Declare Function SetParent& Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long)

' ***** PRIVATE CONSTANTS *****
' system menus
Private Const MF_BYPOSITION = &H400&
Private Const MF_REMOVE = &H1000&

' other
Private Const S_APP_TITLE = "SQL Stored Procedure Creator"
Public Sub Pause(iSeconds As Integer)
    ' pause execution using doevents for specified number of seconds
    Dim fStart As Single
    fStart = Timer
    Do While Timer <= fStart + iSeconds
        DoEvents
    Loop
End Sub
Public Function bLockWindowUpdate(ByVal lHwnd As Long) As Boolean
    ' try and lock / unlock a window / form when being updated
    bLockWindowUpdate = (LockWindowUpdate(lHwnd) = 1)
End Function
Public Sub CenterForm(frm As Form)
    ' Center form frm on the screen
    frm.Left = (Screen.Width - frm.Width) \ 2
    frm.Top = (Screen.Height - frm.Height) \ 2
End Sub
Public Sub RemoveSystemClose(frm As Form)
    Dim hMenu As Long
    Dim itemCount As Long
    
    ' get the handle of the system menu
    hMenu = GetSystemMenu(frm.hWnd, 0)
    ' get the number of items in the menu
    itemCount = GetMenuItemCount(hMenu)
    'remove the system menu Close menu item
    RemoveMenu hMenu, itemCount - 1, MF_REMOVE Or MF_BYPOSITION
    'remove the system menu separator line
    RemoveMenu hMenu, itemCount - 2, MF_REMOVE Or MF_BYPOSITION
End Sub
Public Property Get sAppTitle() As String
    sAppTitle = S_APP_TITLE
End Property
Public Sub SetTheParent(ByVal lhWndChild As Long, _
                             ByVal lhWndParent As Long)
    Dim lResult As Long
    lResult = SetParent(lhWndChild, lhWndParent)
End Sub
Public Property Get sVersion() As String
    sVersion = m_sVersion
End Property
Private Sub Class_Initialize()
    ' set the version
    m_sVersion = App.Major & "." & App.Minor & "." & App.Revision
End Sub
