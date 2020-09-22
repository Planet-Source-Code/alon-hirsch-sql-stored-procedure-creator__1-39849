VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   6780
   ClientLeft      =   2835
   ClientTop       =   1155
   ClientWidth     =   6585
   _ExtentX        =   11615
   _ExtentY        =   11959
   _Version        =   393216
   Description     =   "SQL Stored Procedure Creator"
   DisplayName     =   "SQL Stored Procedure Creator"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "None"
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   RegExtra        =   "SOFTWARE\Syzygy\spCreator"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Create Stored Procedures for SQL Server tables"
Option Explicit

Public p_bFormDisplayed As Boolean
Public p_objVBInstance As VBIDE.VBE
Public WithEvents MenuHandler As CommandBarEvents          'command bar event handler
Attribute MenuHandler.VB_VarHelpID = -1

Private m_objMenuCommandBar As Office.CommandBarControl
Private m_frmCreator As New frmSPCreator
Sub Hide()
    
    On Error Resume Next
    
    p_bFormDisplayed = False
    m_frmCreator.Hide
   
End Sub

Sub Show()
    On Error Resume Next
    
    If m_frmCreator Is Nothing Then
        Set m_frmCreator = New frmSPCreator
    End If
    
    Set m_frmCreator.p_objVBInstance = p_objVBInstance
    Set m_frmCreator.p_objConnect = Me
    p_bFormDisplayed = True
    m_frmCreator.Show
   
End Sub

'------------------------------------------------------
'this method adds the Add-In to VB
'------------------------------------------------------
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    On Error GoTo error_handler
    
    ' instanciate the global object
    Set g_objGlobal = New clsGlobal

    'save the vb instance
    Set p_objVBInstance = Application
    
    'this is a good place to set a breakpoint and
    'test various addin objects, properties and methods
    'Debug.Print p_objVBInstance.FullName
    
    ShowSplashScreen

    If ConnectMode = ext_cm_External Then
        'Used by the wizard toolbar to start this wizard
        Me.Show
    Else
        Set m_objMenuCommandBar = AddToAddInCommandBar("S&QL Stored Procedure Creator")
        'sink the event
        Set Me.MenuHandler = p_objVBInstance.Events.CommandBarEvents(m_objMenuCommandBar)
    End If
  
    If ConnectMode = ext_cm_AfterStartup Then
        If GetSetting(App.Title, "Settings", "DisplayOnConnect", "0") = "1" Then
            'set this to display the form on connect
            Me.Show
        End If
    End If
  
    Exit Sub
    
error_handler:
    
    MsgBox Err.Description
    
End Sub
'------------------------------------------------------
'this method removes the Add-In from VB
'------------------------------------------------------
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    On Error Resume Next
    
    'delete the command bar entry
    m_objMenuCommandBar.Delete
    
    'shut down the Add-In
    If p_bFormDisplayed Then
        SaveSetting App.Title, "Settings", "DisplayOnConnect", "1"
        p_bFormDisplayed = False
    Else
        SaveSetting App.Title, "Settings", "DisplayOnConnect", "0"
    End If
    
    Unload m_frmCreator
    Set m_frmCreator = Nothing

End Sub

Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
    If GetSetting(App.Title, "Settings", "DisplayOnConnect", "0") = "1" Then
        'set this to display the form on connect
        Me.Show
    End If
End Sub

'this event fires when the menu is clicked in the IDE
Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Me.Show
End Sub

Function AddToAddInCommandBar(sCaption As String) As Office.CommandBarControl
    Dim cbMenuCommandBar As Office.CommandBarControl  'command bar object
    Dim cbMenu As Object
  
    On Error GoTo AddToAddInCommandBarErr
    
    'see if we can find the Add-Ins menu
    Set cbMenu = p_objVBInstance.CommandBars("Add-Ins")
    If cbMenu Is Nothing Then
        'not available so we fail
        Exit Function
    End If
    
    'add it to the command bar
    Set cbMenuCommandBar = cbMenu.Controls.Add(1)
    'set the caption
    cbMenuCommandBar.Caption = sCaption
    
    ' add the icon for this program to the command bar
    Clipboard.SetData m_frmCreator.imgIcon.Picture
    cbMenuCommandBar.PasteFace
    
    Set AddToAddInCommandBar = cbMenuCommandBar
    
    Exit Function
    
AddToAddInCommandBarErr:

End Function

