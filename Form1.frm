VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "-->TrayIcon Example -->Trimbitas Sorin --> www.nekhbet.tk"
   ClientHeight    =   1050
   ClientLeft      =   1635
   ClientTop       =   1545
   ClientWidth     =   6915
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1050
   ScaleWidth      =   6915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Exit sample"
      Height          =   615
      Left            =   5280
      TabIndex        =   3
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Show ToolTip"
      Height          =   615
      Left            =   3600
      TabIndex        =   2
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show Tray Icon"
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Hide Tray Icon"
      Height          =   615
      Left            =   1920
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu Example1 
         Caption         =   "Example1"
      End
      Begin VB.Menu Example2 
         Caption         =   "Example2"
      End
      Begin VB.Menu Example3 
         Caption         =   "Example3"
      End
      Begin VB.Menu mnuSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub Command1_Click()
  'add the form's icon to the tray
  TrayCode.AddToTray Me, mnuMain
End Sub

Private Sub Command2_Click()
  'remove the icon from the tray
  TrayCode.RemoveFromTray
End Sub

Private Sub Command3_Click()
  'show a tooltip
  TrayCode.AddToTrayToolTip Me, mnuMain, "This is just a test message!", "Code taken from www.nekhbet.tk", 2
End Sub

Private Sub Command4_Click()
  'close the application
  Unload Me
End Sub

Private Sub Example1_Click()
  'put here some code
  MsgBox "example1"
End Sub

Private Sub Example2_Click()
  'put here some code
  MsgBox "example2"
End Sub

Private Sub Example3_Click()
 'put here some code
  MsgBox "example3"
End Sub

Private Sub Form_Load()
  'you MUST call this function ONLY ONCE in the form_load if
  'you want to use TrayCode module
  Call TrayCode.InitializeTrayModule
End Sub

Private Sub Form_Unload(Cancel As Integer)
  'when the form is about to unload remove the tray icon
  TrayCode.RemoveFromTray
  DoEvents
End Sub

Private Sub mnuExit_Click()
  'close the application
  Unload Me
End Sub
