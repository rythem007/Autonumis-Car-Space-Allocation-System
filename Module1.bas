Attribute VB_Name = "Module1"
Public Const i As Integer = 0
Dim a As Integer

Public Const ws_child As Long = &H40000000
Public Const ws_visible As Long = &H10000000

Global Const WM_USER = 1024
Global Const wm_cap_driver_connect = WM_USER + 10
Global Const wm_cap_set_preview = WM_USER + 50
Global Const WM_CAP_SET_PREVIEWRATE = WM_USER + 52
Global Const WM_CAP_DRIVER_DISCONNECT As Long = WM_USER + 11
Public Const WM_CAP_DLG_VIDEOFORMAT As Long = WM_USER + 41
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function capCreateCaptureWindow Lib "avicap32.dll" Alias "capCreateCaptureWindowA" (ByVal a As String, ByVal b As Long, ByVal c As Integer, ByVal d As Integer, ByVal e As Integer, ByVal f As Integer, ByVal g As Long, ByVal h As Integer) As Long





Public Function abc()
    a = i + 1
    abc = a
End Function

