Attribute VB_Name = "mSubclass"
'Copyright - Stephen Teilhet
Option Explicit

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const WM_SYSKEYDOWN As Long = &H104
Private Const VK_MENU As Long = &H12
Private Const WM_SYSKEYUP As Long = &H105

Public Const WM_GETMINMAXINFO As Long = &H24
Public Const WM_WINDOWPOSCHANGED As Long = &H47
Public Const WM_SIZING = &H214

Public CSubClsApp As cSubclass

Public Function NewWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
      Dim RectStruct As RECT
      Dim OrigRectStruct As RECT
      
      Select Case uMsg
        Case WM_ACTIVATE
            SendMessage frmMain.Hw, WM_ACTIVATE, 2, 0
        Case WM_WINDOWPOSCHANGED
            frmMain.Wm_ResizeMe
      End Select
      
      NewWndProc = CallWindowProc(CSubClsApp.OrigWndProc, hwnd, uMsg, wParam, lParam)
End Function

