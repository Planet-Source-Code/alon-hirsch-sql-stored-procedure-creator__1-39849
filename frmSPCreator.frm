VERSION 5.00
Begin VB.Form frmSPCreator 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SQL Stored Procedure Creator"
   ClientHeight    =   6855
   ClientLeft      =   2760
   ClientTop       =   1275
   ClientWidth     =   6135
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSPCreator.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   6135
   Begin VB.Frame fmeOtherOptions 
      Caption         =   " Other Options "
      Height          =   1100
      Left            =   360
      TabIndex        =   21
      Top             =   5160
      Width           =   5655
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "..."
         Height          =   285
         Left            =   5160
         TabIndex        =   30
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtOutput 
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   29
         Text            =   "txtOutput"
         Top             =   480
         Width           =   3015
      End
      Begin VB.CheckBox chkUpdateDB 
         Caption         =   "Output Folder"
         Height          =   255
         Left            =   360
         TabIndex        =   23
         Top             =   480
         Width           =   1575
      End
      Begin VB.CheckBox chkDropSP 
         Caption         =   "Drop SP if it already exists"
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   800
         Width           =   2295
      End
      Begin VB.CheckBox chkOutput 
         Caption         =   "SP's return values with OUTPUT parameters"
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete the Selected Connection"
      Height          =   375
      Left            =   360
      TabIndex        =   27
      Top             =   6360
      Width           =   2535
   End
   Begin VB.Frame fmeSPs 
      Caption         =   " Which Stored Procedures to generate "
      Height          =   800
      Left            =   360
      TabIndex        =   13
      Top             =   4320
      Width           =   5655
      Begin VB.CheckBox chkSingle 
         Caption         =   "Select a single record"
         Height          =   255
         Left            =   2160
         TabIndex        =   20
         Top             =   480
         Value           =   1  'Checked
         Width           =   3255
      End
      Begin VB.CheckBox chkIdentity 
         Caption         =   "Insert with IDENTITY"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   480
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox chkFields 
         Caption         =   "Select <Fields>"
         Height          =   255
         Left            =   4080
         TabIndex        =   18
         Top             =   240
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox chkSelect 
         Caption         =   "Select *"
         Height          =   255
         Left            =   3120
         TabIndex        =   17
         Top             =   240
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chkDelete 
         Caption         =   "Delete"
         Height          =   255
         Left            =   2160
         TabIndex        =   16
         Top             =   240
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkUpdate 
         Caption         =   "Update"
         Height          =   255
         Left            =   1200
         TabIndex        =   15
         Top             =   240
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkInsert 
         Caption         =   "Insert"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Value           =   1  'Checked
         Width           =   855
      End
   End
   Begin VB.Frame fmeTables 
      Caption         =   " Tables "
      Height          =   2835
      Left            =   360
      TabIndex        =   8
      Top             =   1440
      Width           =   5655
      Begin VB.CommandButton cmdToggle 
         Caption         =   "Reverse Selection"
         Height          =   315
         Left            =   600
         TabIndex        =   12
         Top             =   2400
         Width           =   1575
      End
      Begin VB.CommandButton cmdUnSelectAll 
         Caption         =   "UnSelect All"
         Height          =   315
         Left            =   2280
         TabIndex        =   11
         Top             =   2400
         Width           =   1575
      End
      Begin VB.CommandButton cmdSelectAll 
         Caption         =   "Select All"
         Height          =   315
         Left            =   3960
         TabIndex        =   10
         Top             =   2400
         Width           =   1575
      End
      Begin VB.ListBox lstItems 
         BackColor       =   &H00C0FFFF&
         Height          =   2085
         Left            =   120
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   9
         Top             =   240
         Width           =   5415
      End
   End
   Begin VB.Frame fmeObjects 
      Caption         =   " Objects"
      Height          =   675
      Left            =   360
      TabIndex        =   4
      Top             =   720
      Width           =   5655
      Begin VB.CommandButton cmdLoad 
         Caption         =   "Load SQL Items"
         Height          =   315
         Left            =   3960
         TabIndex        =   7
         Top             =   240
         Width           =   1575
      End
      Begin VB.CheckBox chkViews 
         Alignment       =   1  'Right Justify
         Caption         =   "Views"
         Height          =   255
         Left            =   1320
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox chkTables 
         Alignment       =   1  'Right Justify
         Caption         =   "Tables"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Value           =   1  'Checked
         Width           =   855
      End
   End
   Begin VB.Frame fmeConnect 
      Caption         =   " Connect "
      Height          =   675
      Left            =   360
      TabIndex        =   0
      Top             =   0
      Width           =   5655
      Begin VB.CommandButton cmdConnect 
         Caption         =   "Connect"
         Height          =   315
         Left            =   4680
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdBuild 
         Caption         =   "Build ..."
         Height          =   315
         Left            =   3840
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
      Begin VB.ComboBox cboServer 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.PictureBox picSideBar 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3735
      Left            =   0
      ScaleHeight     =   3735
      ScaleWidth      =   255
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4800
      TabIndex        =   26
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate"
      Height          =   375
      Left            =   3480
      TabIndex        =   25
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Image imgIcon 
      BorderStyle     =   1  'Fixed Single
      Height          =   360
      Left            =   3000
      Picture         =   "frmSPCreator.frx":014A
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   360
   End
End
Attribute VB_Name = "frmSPCreator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefInt A-Z

Public p_objVBInstance As VBIDE.VBE
Public p_objConnect As Connect

' used by the Browse for Folder dialog
Implements IBFFCallback

' private enum
Private Enum blkSelects
    blkSelectAll = 0
    blkUnselectAll = 1
    blkReverseAll = 2
End Enum

' store details of SQL connections
Private Type SQLConnectionType
    sName As String * 20
    sConnection As String * 255
    sDatabase As String * 35
End Type

Private m_sLicenced As String
Private m_lItems As Long
Private m_sOutPutFolder As String

Private m_arrsConnections() As SQLConnectionType

Private m_acnCon As ADODB.Connection

Private m_objSideBar As cLogo
Private m_objRegistry As cRegistry

Private Const S_SECTION_KEY = "Software\"
Private Const S_COMPANY_KEY = "Software\Syzygy\"
Private Const S_APP_KEY = "Software\Syzygy\spCreator\"
Private Function IBFFCallback_Message(ByVal hWndDialog As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
    ' handle callbacks from the Browse for Folder call
    Dim sBuffer As String
    Dim lReturn As Long
    Dim rcMe As RECT
    
    Select Case uMsg
        Case BFFM_INITIALIZED
            'LogMsg "Got BFFM_INITIALIZED message. hWndDialog: 0x" & Hex(hWndDialog)
            
            '// Set a custom caption and move the dialog right beneath the form. Because
            '// this is a standard window handle,there's very little you cannot do with
            '// it, such as setting it on top (for modeless operation) and so on.
            Call SetWindowText(hWndDialog, "Browse for Folder")
            Call GetWindowRect(Me.hWnd, rcMe)
            Call SetWindowPos(hWndDialog, 0, rcMe.Left + 5, rcMe.Top + 5, 0, 0, SWP_NOSIZE Or SWP_NOOWNERZORDER)
        
            '// Set the initial path
            Call BFFSetPath(hWndDialog, m_sOutPutFolder)
            '// Set the initial status text (second line; different from the "title")
            '// Note that this is not applicable if the "new UI" style is set for the
            '// dialog.
            'Call BFFSetStatusText(hWndDialog, "Second status line")
        Case BFFM_SELCHANGED
            If (lParam <> 0) Then
                sBuffer = String$(256, 0)      '// MAX_PATH
                lReturn = SHGetPathFromIDList(lParam, sBuffer)
                'If lReturn <> 0 Then
                '    LogMsg "Selection changed to " & Left$(sBuffer, InStr(sBuffer, Chr$(0)) - 1)
                'End If
                '// NOTE: all your PIDL does /not/ belong to us, so we don't need to
                '// free it. <g>
            End If
        
        '// When the dialog has the edit box (new UI) and the user enters an invalid
        '// path, this is called. lParam contains the text entered. The proc should
        '// return 0 to just close the dialog, or 1 to continue.
        '// If you are *really* showing a message box, you should have to do it using APIs
        '// to specify the HWND parent to be the dialog's, not your VB form.
        Case BFFM_VALIDATEFAILED
            'LogMsg "Validate on '" & PtrToStr(lParam) & "' failed. "
            IBFFCallback_Message = 1        '// Continue
    End Select
End Function

Private Sub CheckGenerate()
    ' check to see if we can enable the generate button or not
    If lstItems.SelCount > 0 Then
        ' we have items selected - ensure that there are options to script
        If (chkInsert.Value = vbChecked) Or (chkUpdate.Value = vbChecked) Or _
            (chkDelete.Value = vbChecked) Or (chkSelect.Value = vbChecked) Or _
            (chkFields.Value = vbChecked) Or (chkIdentity.Value = vbChecked) Or _
            (chkSingle.Value = vbChecked) Then
            ' at least one of these is selected - check further
            If (chkUpdateDB.Value = vbUnchecked) And (txtOutput.Text = "") Then
                ' we are not updating the database and no output folder specified
                cmdGenerate.Enabled = False
            Else
                ' we are not updating the database, but we have an output folder
                cmdGenerate.Enabled = True
            End If
        Else
            ' no items are checked - can't generate
            cmdGenerate.Enabled = False
        End If
    Else
        ' no items selected - nothing to generate
        cmdGenerate.Enabled = False
    End If
End Sub
Private Sub Disconnect()
    ' if we have an ADO connection - disconnect
    If Not (m_acnCon Is Nothing) Then
        If m_acnCon.State = adStateOpen Then
            m_acnCon.Close
        End If
    End If
    DisConnected
End Sub

Private Sub DisConnected()
    ' when we are disconnected - these items are not enabled
    cboServer.ListIndex = -1
    cmdConnect.Enabled = False
    cmdDelete.Enabled = False
    fmeObjects.Enabled = False
    cmdLoad.Enabled = False
    fmeTables.Enabled = False
    lstItems.Clear
    fmeTables.Caption = " Tables "
    fmeSPs.Enabled = False
    cmdGenerate.Enabled = False
End Sub
Private Sub LoadDefaults()
    ' load the default values for this application
    Dim lItems As Long
    Dim lItem As Long
    Dim arrsItems() As String
    
    ' clear the server combo box
    cboServer.Clear
    
    ' destroy the array before we start
    ReDim m_arrsConnections(0)
    
    ' read stuff from the registry
    With m_objRegistry
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = S_COMPANY_KEY
        
        If Not .KeyExists Then
            ' create the key
            .CreateKey
        End If
        .SectionKey = S_APP_KEY
        If Not .KeyExists Then
            ' create the key
            .CreateKey
        End If
        
        .ValueKey = "Output Folder"
        .ValueType = REG_SZ
        m_sOutPutFolder = .Value
        txtOutput.Text = m_sOutPutFolder
        
        ' now that all the registry KEY and SECTION entries are present - we can start
        .SectionKey = S_APP_KEY & "Servers"
        If Not .KeyExists Then
            ' create the key
            .CreateKey
        End If
        
        If .EnumerateSections(arrsItems, lItems) Then
            ' all items enumerated - get the list of servers from the array
            If lItems > 0 Then
                ' we have some items in the list - we can proceed
                ReDim m_arrsConnections(1 To lItems)
                ' loop through all the sections (Servers)
                For lItem = 1 To lItems
                    m_arrsConnections(lItem).sName = Trim$(arrsItems(lItem))
                    .SectionKey = S_APP_KEY & "Servers\" & Trim$(arrsItems(lItem))
                    .ValueType = REG_SZ
                    .ValueKey = "Connection"
                    m_arrsConnections(lItem).sConnection = .Value
                    .ValueKey = "Database"
                    m_arrsConnections(lItem).sDatabase = .Value
                    cboServer.AddItem UCase$(Trim$(arrsItems(lItem))) & " - " & .Value
                Next lItem
            End If
            ' store the total number of entries
            m_lItems = lItems
        End If
        
        ' determine if we are licenced or not
        m_sLicenced = " (Unlicenced)"
    End With
End Sub
Private Sub SetFrameCaption()
    Dim sCaption As String
    
    sCaption = ""
    If chkTables.Value = vbChecked Then
        sCaption = "Tables"
    End If
    If chkViews.Value = vbChecked Then
        If sCaption = "" Then
            sCaption = "Views"
        Else
            sCaption = sCaption & " / Views"
        End If
    End If
    sCaption = " " & sCaption & " "
    fmeTables.Caption = sCaption
End Sub
Private Sub ToggleSelected(ByVal blkFlag As blkSelects)
    Dim lItems As Long
    Dim lItem As Long
    Dim lCurrent As Long
    Dim lTop As Long
    
    ' start by disabling the screen refresh
    If Not g_objGlobal.bLockWindowUpdate(lstItems.hWnd) Then
    End If
    
    ' detrmine how many items there are
    lItems = lstItems.ListCount - 1
    lCurrent = lstItems.ListIndex
    lTop = lstItems.TopIndex
    ' loop through all the items
    For lItem = 0 To lItems
        ' process each item
        Select Case blkFlag
            Case Is = blkSelectAll
                ' select this item
                lstItems.Selected(lItem) = True
            Case Is = blkUnselectAll
                ' unselect this item
                lstItems.Selected(lItem) = False
            Case Is = blkReverseAll
                ' reverse the selection of this item
                lstItems.Selected(lItem) = Not (lstItems.Selected(lItem))
        End Select
    Next lItem
    lstItems.ListIndex = lCurrent
    lstItems.TopIndex = lTop
    
    ' enable the screen updates again
    If Not g_objGlobal.bLockWindowUpdate(0) Then
    End If
End Sub

Private Sub CancelButton_Click()
    Disconnect
    p_objConnect.Hide
End Sub
Private Sub cboServer_Click()
    If cboServer.ListIndex >= 0 Then
        ' we have an items selected - enabled the connect
        cmdConnect.Enabled = True
        cmdDelete.Enabled = True
    Else
        ' no item selected - can't connect
        cmdConnect.Enabled = False
        cmdDelete.Enabled = False
    End If
End Sub


Private Sub chkDelete_Click()
    CheckGenerate
End Sub
Private Sub chkFields_Click()
    CheckGenerate
End Sub

Private Sub chkIdentity_Click()
    CheckGenerate
End Sub


Private Sub chkInsert_Click()
    CheckGenerate
End Sub

Private Sub chkSelect_Click()
    CheckGenerate
End Sub

Private Sub chkSingle_Click()
    CheckGenerate
End Sub

Private Sub chkTables_Click()
    cmdLoad.Enabled = ((chkTables.Value = vbChecked) Or (chkViews.Value = vbChecked))
    SetFrameCaption
End Sub

Private Sub chkUpdate_Click()
    CheckGenerate
End Sub

Private Sub chkUpdateDB_Click()
    If chkUpdateDB.Value = vbChecked Then
        ' we must update the database
        chkUpdateDB.Caption = "Update Database"
        txtOutput.Visible = False
        cmdBrowse.Visible = False
        cmdBrowse.Enabled = False
    Else
        ' we must output to file
        chkUpdateDB.Caption = "Output Folder"
        txtOutput.Visible = True
        cmdBrowse.Visible = True
        cmdBrowse.Enabled = True
    End If
    CheckGenerate
End Sub

Private Sub chkViews_Click()
    cmdLoad.Enabled = ((chkTables.Value = vbChecked) Or (chkViews.Value = vbChecked))
    SetFrameCaption
End Sub

Private Sub cmdBrowse_Click()
    ' browse for a folder for the output files
    Dim sFolder As String
    Dim objBrowse As CShellBrowser
    
    Set objBrowse = New CShellBrowser
    
    sFolder = objBrowse.BrowseForFolder(Me, Me.hWnd, bifBrowseIncludeFiles Or bifUseNewUI, "Select Folder for output Stored Procedures")
    
    If sFolder <> "" Then
        ' a folder was selected
        m_sOutPutFolder = sFolder
        txtOutput.Text = m_sOutPutFolder
    End If
    
    ' and now clean up a bit
    Set objBrowse = Nothing
End Sub

Private Sub cmdBuild_Click()
    ' invoke OLE DB connection builders
    Dim sConnection As String
    Dim sName As String
    Dim sDatabase As String
    Dim lItems As Long
    Dim lPos As Long
    Dim lPos1 As Long
    
    Dim acnConn As ADODB.Connection
    Dim objDataLink As MSDASC.DataLinks
    
    On Local Error GoTo ErrHandler
    
    Set acnConn = New ADODB.Connection
    Set objDataLink = New MSDASC.DataLinks

    acnConn = objDataLink.PromptNew
    sConnection = acnConn.ConnectionString
    
    If sConnection <> "" Then
        ' it is a valid connection (hopefully) - asdk the user to name the connection
        lPos = InStr(1, sConnection, "Initial Catalog=", vbTextCompare)
        If lPos <> 0 Then
            ' we have a valid connection string - strip the database name
            lPos1 = InStr(lPos, sConnection, ";", vbTextCompare)
            If lPos1 <> 0 Then
                ' so far so good - get the database name
                sDatabase = Mid$(sConnection, lPos + 16, lPos1 - lPos - 16)
                frmConnection.p_sConnection = sConnection
                frmConnection.p_sDatabase = sDatabase
                frmConnection.Show vbModal, Me
                sName = Trim$(UCase$(frmConnection.p_sName))
                If sName <> "" Then
                    ' we have entered a name - use it
                    'If LBound(m_arrsConnections) = 0 Then
                    If m_lItems = 0 Then
                        m_lItems = 1
                        ReDim m_arrsConnections(1 To m_lItems)
                    Else
                        'lItems = UBound(m_arrsConnections) + 1
                        m_lItems = m_lItems + 1
                        ReDim Preserve m_arrsConnections(1 To m_lItems)
                    End If
                    
                    m_arrsConnections(m_lItems).sConnection = sConnection
                    m_arrsConnections(m_lItems).sDatabase = sDatabase
                    m_arrsConnections(m_lItems).sName = sName
                    Set frmConnection = Nothing
                    
                    ' now add it to the combo box
                    cboServer.AddItem sName & " - " & sDatabase
                    
                    ' and save the details to the registry
                    m_objRegistry.SectionKey = S_APP_KEY & "\Servers\" & sName
                    If Not m_objRegistry.KeyExists Then
                        ' create the key
                        m_objRegistry.CreateKey
                    End If
                    ' now store the details
                    m_objRegistry.ValueType = REG_SZ
                    m_objRegistry.ValueKey = "Connection"
                    m_objRegistry.Value = sConnection
                    m_objRegistry.ValueKey = "Database"
                    m_objRegistry.Value = sDatabase
                End If
            End If
        End If
    End If
    
    Exit Sub
ErrHandler:
    ' skip known error conditions; otherwise, report error
    ' (e.g. skip ODBC builder 'action cancelled' error)
    If Err.Number = 91 Or Err.Number = -2147217842 Then
        Exit Sub
    Else
        MsgBox "Error: " & Err.Description
    End If
End Sub
Private Sub cmdConnect_Click()
    Dim sConnection As String
    Dim sName As String
    Dim sDatabase As String
    Dim sMsg As String
    Dim sItem As String
    Dim lErrors As Long
    Dim lItem As Long
    Dim bConnected As Boolean
    
    ' ensure that we do not an open connection
    If m_acnCon Is Nothing Then
        ' we do not yet have a connection object - create one
        Set m_acnCon = New ADODB.Connection
        With m_acnCon
            .CommandTimeout = 15
            .ConnectionTimeout = 15
            .CursorLocation = adUseServer
        End With
    End If
    If m_acnCon.State = adStateOpen Then
        ' the connection is open - close it
        m_acnCon.Close
    End If
    
    ' determine the database and connection name
    sItem = cboServer.List(cboServer.ListIndex)
    lItem = InStrRev(sItem, "-")
    sDatabase = Trim$(Mid$(sItem, lItem + 2))
    sName = Trim$(Left$(sItem, lItem - 2))
    
    ' find the selected item in the array of connections
    sConnection = ""
    For lItem = 1 To m_lItems
        ' find the matching record in the array
        If Trim$(m_arrsConnections(lItem).sName) = sName Then
            ' the name is the same - check the database name
            If StrComp(Trim$(m_arrsConnections(lItem).sDatabase), sDatabase, vbTextCompare) = 0 Then
                ' we have found a match - exit the loop and use the connection string
                sConnection = Trim$(m_arrsConnections(lItem).sConnection)
                Exit For
            End If
        End If
    Next lItem
    
    ' now ensuure that we have a connection string
    If sConnection <> "" Then
        ' we have a connection string - connect
        m_acnCon.ConnectionString = sConnection
        Me.MousePointer = vbHourglass
        On Local Error GoTo ConnectionError
        m_acnCon.Open
        On Local Error GoTo 0
        Me.MousePointer = vbDefault
        fmeObjects.Enabled = True
        chkTables_Click
    Else
        ' no connection string - report error
        MsgBox "The connection string is missing"
    End If
EndOfConnect:
    Exit Sub
'--------------------------------------------------------------------------------------------
ConnectionError:    ' there was an error connecting to the data source / database
'--------------------------------------------------------------------------------------------
    sMsg = "There was an error connecting the data source :" & vbCrLf
    lErrors = m_acnCon.Errors.Count
'MsgBox "Errors : " & lErrors
'    For lItem = 1 To lErrors
'        sMsg = sMsg & "   ADO Error (" & lItem & ") : " & m_acnCon.Errors(lItem).Number & vbCrLf
'        sMsg = sMsg & "     Description : " & m_acnCon.Errors(lItem).Description & vbCrLf
'        sMsg = sMsg & "     SQL State : " & m_acnCon.Errors(lItem).SQLState & vbCrLf
'        sMsg = sMsg & "     Native Error : " & m_acnCon.Errors(lItem).NativeError & vbCrLf
'    Next lItem
    sMsg = sMsg & "   Error : " & Err.Number & vbCrLf
    sMsg = sMsg & "     Description : " & Err.Description & vbCrLf
    
    sMsg = sMsg & vbCrLf
    If m_acnCon.State = adStateOpen Then
        sMsg = sMsg & "You are connected to the Server / Database."
    Else
        sMsg = sMsg & "You are not connected to the Server / Database"
    End If
    bConnected = (m_acnCon.State = adStateOpen)
    Me.MousePointer = vbDefault
    MsgBox sMsg, vbInformation + vbOKOnly, g_objGlobal.sAppTitle
    If bConnected Then
        Resume Next
    Else
        Resume EndOfConnect
    End If
End Sub

Private Sub cmdDelete_Click()
    ' This will remove the selected connection from the list and registry
    Dim sMsg As String
    Dim sItem As String
    Dim sName As String
    Dim sDatabase As String
    Dim lPos As Long
    Dim lItems As Long
    Dim lItem As Long
    Dim iChoice As Integer
    Dim arrsServers() As String
    
    ' make sure the user wants to do this
    sMsg = "You are about to delete a Database Connection" & vbCrLf
    sMsg = sMsg & "from this program." & vbCrLf
    sMsg = sMsg & "This setting stores all the information required" & vbCrLf
    sMsg = sMsg & "in order to connect to the specified database, including" & vbCrLf
    sMsg = sMsg & "the user name and password." & vbCrLf
    sMsg = sMsg & "Once deleted, you will have to recreate the connection" & vbCrLf
    sMsg = sMsg & "in order to connect again to this database." & vbCrLf & vbCrLf
    sMsg = "Are you sure you want to delete this connection ?"
    iChoice = MsgBox(sMsg, vbQuestion + vbYesNo + vbDefaultButton2, g_objGlobal.sAppTitle)
    
    ' check what was chosen
    If iChoice = vbNo Then
        ' do not delete
        Exit Sub
    End If
    
    ' the user wan't it gone - delete it from the registry and then load defaults again
    sItem = cboServer.List(cboServer.ListIndex)
    lPos = InStrRev(sItem, "-")
    sName = Trim$(Left$(sItem, lPos - 2))
    sDatabase = Trim$(Mid$(sItem, lPos + 2))
MsgBox sName
MsgBox sDatabase
    With m_objRegistry
        .SectionKey = S_APP_KEY & "Servers"
        ' enumerate all the server names
        If .EnumerateSections(arrsServers, lItems) Then
            ' we have a list of the servers
            If lItems <> 0 Then
                ' we have a list of items - check them all
                For lItem = 1 To lItems
                    ' check each one to see if it is the connection we are looking for
                    If arrsServers(lItem) = sName Then
                        ' this is the name we are looking for - check the database
                        .SectionKey = S_APP_KEY & "Servers\" & arrsServers(lItem)
                        .ValueType = REG_SZ
                        .ValueKey = "Database"
                        If .Value = sDatabase Then
                            ' we have found our match - delete this section
                            .DeleteKey
                            Exit For
                        End If
                    End If
                Next lItem
            End If
        End If
    End With
    
    ' load the defaults
    LoadDefaults
End Sub

Private Sub cmdGenerate_Click()
    Me.Enabled = False
    
    If Not g_objGlobal.bLockWindowUpdate(lstItems.hWnd) Then
    End If
    
    ' save the output folder to the registry if one is selected
    If chkUpdateDB.Value = vbUnchecked Then
        ' we are to output the results and not save them to the DB
        With m_objRegistry
            .SectionKey = S_APP_KEY
            .ValueType = REG_SZ
            .ValueKey = "Output Folder"
            .Value = m_sOutPutFolder
        End With
    End If
    
    GenerateAllStoredProcs m_acnCon, lstItems, (chkInsert.Value = vbChecked), _
                                                (chkUpdate.Value = vbChecked), _
                                                (chkDelete.Value = vbChecked), _
                                                (chkSelect.Value = vbChecked), _
                                                (chkFields.Value = vbChecked), _
                                                (chkIdentity.Value = vbChecked), _
                                                (chkSingle.Value = vbChecked), _
                                                (chkOutput.Value = vbChecked), _
                                                (chkUpdateDB.Value = vbChecked), _
                                                (chkDropSP.Value = vbChecked), _
                                                m_sOutPutFolder

    ' move to the top of the list
    lstItems.ListIndex = 0
    
    If Not g_objGlobal.bLockWindowUpdate(0) Then
    End If
        
    Me.Enabled = True
    Me.SetFocus
End Sub

Private Sub cmdLoad_Click()
    ' This will load all the tables / views etc. as per the options
    Dim sItem As String
    Dim lItem As Long
    Dim arsData As ADODB.Recordset
    
    ' enable the tables frame and the scripting options
    cmdLoad.Enabled = False
    fmeTables.Enabled = True
    fmeSPs.Enabled = True
    
    ' clear the entire list of items
    lstItems.Clear
    
    ' disable screen refresh while we are adding items
    If Not g_objGlobal.bLockWindowUpdate(lstItems.hWnd) Then
    End If
    Screen.MousePointer = vbHourglass
   
    ' load the tables
    If (chkTables.Value = vbChecked) Then
        ' only tables
        Set arsData = m_acnCon.OpenSchema(adSchemaTables, Array(Empty, Empty, Empty, "TABLE"))
        GoSub LoadItemsIntoList
    End If
    
    ' load the views
    If (chkTables.Value = vbChecked) Then
        ' only views
        Set arsData = m_acnCon.OpenSchema(adSchemaTables, Array(Empty, Empty, Empty, "VIEW"))
        GoSub LoadItemsIntoList
    End If
    
    ' enable the refresh again
    If Not g_objGlobal.bLockWindowUpdate(0) Then
    End If
    
    If lstItems.ListCount <> 0 Then
        ' there are items in the list - set focus to thelist
        lstItems.SetFocus
    Else
        ' inform the user that there are no items
        sItem = ""
        If chkTables.Value = vbChecked Then
            sItem = "Tables"
        End If
        If chkViews.Value = vbChecked Then
            If sItem = "" Then
                ' no tables checked
                sItem = "Views"
            Else
                ' there are no tables either
                sItem = sItem & " or Views"
            End If
        End If
        sItem = "There are no " & sItem & " in this database."
        MsgBox sItem, vbInformation + vbOKOnly, g_objGlobal.sAppTitle
    End If
    
    Screen.MousePointer = vbDefault
    Set arsData = Nothing
    cmdLoad.Enabled = True
    Exit Sub
'--------------------------------------------------------------------------------------------
LoadItemsIntoList:  ' load the items from the recordset into the list
'--------------------------------------------------------------------------------------------
    If Not (arsData.EOF And arsData.BOF) Then
        Do While Not arsData.EOF
            ' load each item
            sItem = arsData("Table_Name") & "   (" & arsData("Table_Type") & ")"
            lstItems.AddItem sItem
            arsData.MoveNext
            DoEvents
        Loop
        arsData.Close
    End If
    Return
End Sub
Private Sub cmdSelectAll_Click()
    ToggleSelected blkSelectAll
End Sub

Private Sub cmdToggle_Click()
    ToggleSelected blkReverseAll
End Sub

Private Sub cmdUnSelectAll_Click()
    ToggleSelected blkUnselectAll
End Sub


Private Sub Form_Load()
    Set m_objSideBar = New cLogo
    Set m_objRegistry = New cRegistry
        
    ' display the sidebar
    With m_objSideBar
        .DrawingObject = picSideBar
        .Caption = " " & App.LegalCopyright & " " & App.CompanyName
        .StartColor = vbBlack
        .EndColor = &H8000000F
        .Draw
    End With
    
    ' load the application defaults
    LoadDefaults
    
    ' prevent the user from connecting until they select a connection
    DisConnected
    
    ' update the version number
    Me.Caption = "SQL SP Creator - version " & g_objGlobal.sVersion & m_sLicenced
    
    ' do whatever needs to be done when a form is loaded
    WhenFormIsLoaded Me
End Sub
Private Sub Form_Resize()
    picSideBar.Top = 0
    picSideBar.Height = Me.ScaleHeight
    m_objSideBar.Draw
End Sub


Private Sub Form_Unload(Cancel As Integer)
    ' we are terminating the program
    Set m_objSideBar = Nothing
    Set m_objRegistry = Nothing
    Set g_objGlobal = Nothing
    
    ' close any open connections
    If Not (m_acnCon Is Nothing) Then
        If m_acnCon.State = adStateOpen Then
            m_acnCon.Close
        End If
        ' and destroy it
        Set m_acnCon = Nothing
    End If
End Sub

Private Sub imgIcon_DblClick()
    frmSplash.p_bOK = True
    frmSplash.Show vbModal, Me
    Set frmSplash = Nothing
End Sub


Private Sub lstItems_Click()
    CheckGenerate
End Sub


