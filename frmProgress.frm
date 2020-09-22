VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProgress 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2100
   ClientLeft      =   2985
   ClientTop       =   2745
   ClientWidth     =   4935
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ProgressBar prgScripts 
      Height          =   320
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar prgTables 
      Height          =   320
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   4800
      Y1              =   1100
      Y2              =   1100
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   4800
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label lblScriptPerc 
      Alignment       =   2  'Center
      Caption         =   "lblScriptPerc"
      Height          =   255
      Left            =   1800
      TabIndex        =   9
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label lblScriptTotal 
      Alignment       =   1  'Right Justify
      Caption         =   "lblScriptTotal"
      Height          =   255
      Left            =   3600
      TabIndex        =   8
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label lblScriptFrom 
      Caption         =   "lblScriptFrom"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label lblScript 
      Caption         =   "lblScript"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   4575
   End
   Begin VB.Label lblTablePerc 
      Alignment       =   2  'Center
      Caption         =   "lblTablePerc"
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label lblTableTotal 
      Alignment       =   1  'Right Justify
      Caption         =   "lblTableTotal"
      Height          =   255
      Left            =   3600
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label lblTableFrom 
      Caption         =   "lblTableFrom"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label lblTable 
      Caption         =   "lblTable"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    WhenFormIsLoaded Me
End Sub


