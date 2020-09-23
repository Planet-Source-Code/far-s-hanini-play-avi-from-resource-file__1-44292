VERSION 5.00
Begin VB.Form Animate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "File Download"
   ClientHeight    =   2025
   ClientLeft      =   4665
   ClientTop       =   3990
   ClientWidth     =   5565
   Icon            =   "frmAnimate.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Enabled         =   0   'False
      Height          =   345
      Left            =   4335
      TabIndex        =   3
      Top             =   1500
      Width           =   1080
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "setup.zip from 987.456.133.222"
      Height          =   255
      Left            =   135
      TabIndex        =   2
      Top             =   1110
      Width           =   3975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Saving:"
      Height          =   255
      Left            =   135
      TabIndex        =   1
      Top             =   870
      Width           =   1890
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Estimated Time Left: 1 hr 4 min 30 sec (400 kb copied)"
      Height          =   255
      Left            =   135
      TabIndex        =   0
      Top             =   1770
      Width           =   3930
   End
End
Attribute VB_Name = "Animate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Animate As New CAnimate32



Private Sub Form_Load()

On Error Resume Next
 
If Dir(App.Path & "\FILECOPY.AVI") = "" Then
MsgBox "Unable to find AVI"
Unload Me
End If

Animate.Create Me.hwnd, App.Path & "\FILECOPY.AVI", -5, 0, 300, 50
 
Animate.AnimatePlay
 
Me.Show
 
 
 
End Sub

