VERSION 5.00
Begin VB.Form frmConnection 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Server Connection Information"
   ClientHeight    =   3855
   ClientLeft      =   3675
   ClientTop       =   2355
   ClientWidth     =   3855
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConnection.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdUse 
      Caption         =   "&Use"
      Height          =   375
      Left            =   1440
      TabIndex        =   7
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox txtConnection 
      BackColor       =   &H00C0C0C0&
      Height          =   2085
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Text            =   "frmConnection.frx":014A
      Top             =   1200
      Width           =   3615
   End
   Begin VB.TextBox txtDatabase 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   3
      Text            =   "txtDatabase"
      Top             =   480
      Width           =   2775
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   960
      MaxLength       =   25
      TabIndex        =   1
      Text            =   "txtName"
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "ADO Connection String"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   3615
   End
   Begin VB.Label Label2 
      Caption         =   "Database"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefInt A-Z

Public p_sName As String
Public p_sDatabase As String
Public p_sConnection As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdUse_Click()
    p_sName = UCase$(Trim$(txtName.Text))
    cmdCancel.Value = True
End Sub
Private Sub Form_Load()
    txtName.Text = ""
    txtDatabase.Text = p_sDatabase
    txtConnection.Text = p_sConnection
    cmdUse.Enabled = False
    WhenFormIsLoaded Me
End Sub

Private Sub txtName_Change()
    cmdUse.Enabled = (Trim$(txtName.Text) <> "")
End Sub
