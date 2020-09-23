VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAnimate32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Type tagInitCommonControlsEx
    lngSize As Long
    lngICC As Long
End Type
 
Private Declare Function InitCommonControlsEx Lib "Comctl32.dll" _
 (iccex As tagInitCommonControlsEx) As Boolean

Private Const ICC_ANIMATE_CLASS = &H80
Private Const ANIMATE_CLASS = "SysAnimate32"
Private Const WM_USER = &H400&
Private Const ACS_CENTER = &H1&
Private Const ACS_TRANSPARENT = &H2&
Private Const ACS_AUTOPLAY = &H4&
Private Const ACM_OPEN = WM_USER + 100
Private Const ACM_PLAY = WM_USER + 101
Private Const ACM_STOP = WM_USER + 102
Private Const WS_EX_TRANSPARENT = &H20&
Private AnimateHwnd As Long
Private StaticWin As Long
 

Public Sub Create(ByVal hwnd As Long, ByVal StrAVI As String, ByVal Left As Long, ByVal Top As Long, ByVal width As Long, ByVal height As Long)
    
    StaticWin = CreateWindowEx(WS_EX_TRANSPARENT, _
    ANIMATE_CLASS, _
    "", _
     &H50000007, _
    Left, Top, width, height, _
    hwnd, 0&, App.hInstance, ByVal 0&)
    
    
    AnimateHwnd = CreateWindowEx(WS_EX_TRANSPARENT, _
    ANIMATE_CLASS, _
    "", _
     &H50000007, _
    Left, Top, width, height, _
    StaticWin, 0&, App.hInstance, ByVal 0&)
    
    
    SendMessage AnimateHwnd, ACM_OPEN Or ACS_AUTOPLAY, 0&, ByVal StrAVI
  

End Sub

Public Sub AnimatePlay()
    SendMessage AnimateHwnd, ACM_PLAY, -1, 0
End Sub

Public Sub AnimateStop()
    SendMessage AnimateHwnd, ACM_STOP, 0, 0
End Sub

Public Sub Destroy()
    AnimateStop
    DestroyWindow AnimateHwnd
    DestroyWindow StaticWin
     
End Sub


Public Sub AutoPlay()
    SendMessage AnimateHwnd, ACS_AUTOPLAY, -1, 0
End Sub

Private Sub Class_Initialize()
 Dim iccex As tagInitCommonControlsEx
    With iccex
        .lngSize = LenB(iccex)
        .lngICC = ICC_ANIMATE_CLASS
    End With
    Call InitCommonControlsEx(iccex)
     
End Sub
