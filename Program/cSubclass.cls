VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSubclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Copyright - Stephen Teilhet
Option Explicit
Private Declare Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWLP_WNDPROC = -4
Private m_lOrigWndProc As Long
Private m_hwnd As Long

Public Property Get OrigWndProc() As Long
      OrigWndProc = m_lOrigWndProc
End Property
Public Property Let hwnd(Handle As Long)
      m_hwnd = Handle
End Property
Private Sub Class_Initialize()
      m_lOrigWndProc = 0
      m_hwnd = 0
End Sub
Public Sub EnableSubclass()
      If m_lOrigWndProc <> 0 Then Exit Sub
      m_lOrigWndProc = SetWindowLongPtr(m_hwnd, GWLP_WNDPROC, AddressOf mSubclass.NewWndProc)
End Sub
Public Sub DisableSubclass()
      If m_lOrigWndProc > 0 Then
            SetWindowLongPtr m_hwnd, GWLP_WNDPROC, m_lOrigWndProc
            m_lOrigWndProc = 0
      End If
End Sub
Private Sub Class_Terminate()
      Call DisableSubclass
End Sub
