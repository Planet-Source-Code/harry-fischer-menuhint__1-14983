VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frm_main 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MenuHint Test"
   ClientHeight    =   4095
   ClientLeft      =   6060
   ClientTop       =   6015
   ClientWidth     =   8565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   8565
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   3780
      Width           =   8565
      _ExtentX        =   15108
      _ExtentY        =   556
      Style           =   1
      ShowTips        =   0   'False
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Menu MenuFile 
      Caption         =   "&File"
      Begin VB.Menu MenuFileOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu MenuFileSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu MenuFileClose 
         Caption         =   "&Close Commands"
         Begin VB.Menu MenuFileCloseCurrWin 
            Caption         =   "Current Window"
         End
         Begin VB.Menu MenuFileCloseAllWin 
            Caption         =   "All Windows"
         End
      End
      Begin VB.Menu MenuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu MenuFileExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frm_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

  gHWnd = Me.hwnd
  pHook Me.hwnd

End Sub

Private Sub Form_Unload(Cancel As Integer)

  pUnhook Me.hwnd

End Sub
