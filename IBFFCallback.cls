VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IBFFCallback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------------------
'   Name:           IBFFCallback (IBFFCALLBACK.CLS)
'   Type:           Base class
'   Description:    Used to implement SHBrowseForFolder callbacks
'
'   Author:         Klaus H. Probst [kprobst@vbbox.com]
'   URL:            http://www.vbbox.com
'   Copyright:      This work is copyright © 1998-2001, Klaus H. Probst
'   Usage:          You may use this code as you see fit, provided that you assume all
'                   responsibilities for doing so.
'   Distribution:   If you intend to distribute the file(s) that make up this sample to
'                   any WWW site, online service, electronic bulletin board system (BBS),
'                   CD or any other electronic or physical media, you must notify me in
'                   advance to obtain my express permission.
'
'   Notes:
'                   The modCBShBrowseForFolder module contains a few helper methods
'                   that can be used to interact with the BFF dialog. Those methods
'                   require the hWndDialog argument passed to the method defined in
'                   this class.
'
'   Dependencies:
'
'       (bas)       modCBShBrowseForFolder
'       (cls)       CShellBrowser
'
'--------------------------------------------------------------------------------------------
Option Explicit
DefLng A-Z

Public Function Message(ByVal hWndDialog As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
    ' callback
End Function

