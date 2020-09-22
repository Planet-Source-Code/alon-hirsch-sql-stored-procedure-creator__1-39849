VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About ..."
   ClientHeight    =   3930
   ClientLeft      =   2445
   ClientTop       =   2325
   ClientWidth     =   6480
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0FFFF&
      Height          =   1335
      Left            =   360
      ScaleHeight     =   1275
      ScaleWidth      =   5955
      TabIndex        =   7
      Top             =   2040
      Width           =   6015
      Begin VB.Label lblDescription 
         BackStyle       =   0  'Transparent
         Caption         =   "lblDescription"
         Height          =   1180
         Left            =   50
         TabIndex        =   8
         Top             =   50
         Width           =   5850
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Timer tmrClose 
      Enabled         =   0   'False
      Interval        =   3500
      Left            =   6000
      Top             =   120
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
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   255
   End
   Begin VB.Label lblCopyright 
      Caption         =   "lblCopyright"
      Height          =   255
      Left            =   480
      TabIndex        =   9
      Top             =   3480
      Width           =   4575
   End
   Begin VB.Label lblComments 
      Alignment       =   2  'Center
      Caption         =   "Comments"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   1680
      Width           =   6375
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   1440
      Width           =   6375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "SQL Stored Procedure Creator"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   1080
      Width           =   6375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Presents ..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Width           =   6375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Syzygy Computer Services cc"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   6375
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefInt A-Z

Private m_objSideBar As cLogo
Public p_bOK As Boolean

Private Sub cmdOK_Click()
    Unload Me
End Sub


Private Sub Form_Load()
    Dim sText As String
    
    Set m_objSideBar = New cLogo
    
    If p_bOK Then
        ' display the OK button
        cmdOK.Enabled = True
        cmdOK.Visible = True
    Else
        cmdOK.Enabled = False
        cmdOK.Visible = False
        tmrClose.Enabled = True
    End If
    
    lblVersion.Caption = "Version " & g_objGlobal.sVersion
    lblComments.Caption = App.Comments
    
    ' set the program description
    sText = "This program was designed as a VB6 add-in in order to " & vbCrLf
    sText = sText & "facilitate the creation of basic / default SQL Stored Procedures" & vbCrLf
    sText = sText & "for SQL Server / VB projects." & vbCrLf & vbCrLf
    sText = sText & "This project was designed to 'make my life easier' while developing" & vbCrLf
    sText = sText & "Visual Basic / SQL Server projects."
    lblDescription.Caption = sText
    
    ' and the copyright stuff
    lblCopyright = App.LegalCopyright & " " & App.CompanyName
    
    ' display the sidebar
    With m_objSideBar
        .DrawingObject = picSideBar
        .Caption = " " & App.CompanyName
        .StartColor = vbBlack
        .EndColor = &H8000000F
        .Draw
    End With
    
    WhenFormIsLoaded Me
End Sub
Private Sub Form_Resize()
    picSideBar.Top = 0
    picSideBar.Height = Me.ScaleHeight
    m_objSideBar.Draw
End Sub
Private Sub tmrClose_Timer()
    Me.Enabled = False
    Unload Me
End Sub


