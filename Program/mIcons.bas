Attribute VB_Name = "mIcons"
'From VBaccelerator
Option Explicit

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function LoadImageAsString Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long

Private Const SM_CXICON = 11
Private Const SM_CYICON = 12
Private Const SM_CXSMICON = 49
Private Const SM_CYSMICON = 50
Private Const LR_DEFAULTCOLOR = &H0
Private Const LR_MONOCHROME = &H1
Private Const LR_COLOR = &H2
Private Const LR_COPYRETURNORG = &H4
Private Const LR_COPYDELETEORG = &H8
Private Const LR_LOADFROMFILE = &H10
Private Const LR_LOADTRANSPARENT = &H20
Private Const LR_DEFAULTSIZE = &H40
Private Const LR_VGACOLOR = &H80
Private Const LR_LOADMAP3DCOLORS = &H1000
Private Const LR_CREATEDIBSECTION = &H2000
Private Const LR_COPYFROMRESOURCE = &H4000
Private Const LR_SHARED = &H8000&
Private Const IMAGE_ICON = 1
Private Const WM_SETICON = &H80
Private Const ICON_SMALL = 0
Private Const ICON_BIG = 1
Private Const GW_OWNER = 4

Public Sub SetIcon(ByVal hwnd As Long, ByVal sIconResName As String, Optional ByVal bSetAsAppIcon As Boolean = True)
Dim lhWndTop As Long
Dim lhWnd As Long
Dim cx As Long
Dim cy As Long
Dim hIconLarge As Long
Dim hIconSmall As Long
      
   If (bSetAsAppIcon) Then
      ' Find VB's hidden parent window:
      lhWnd = hwnd
      lhWndTop = lhWnd
      Do While Not (lhWnd = 0)
         lhWnd = GetWindow(lhWnd, GW_OWNER)
         If Not (lhWnd = 0) Then
            lhWndTop = lhWnd
         End If
      Loop
   End If
   
   cx = GetSystemMetrics(SM_CXICON)
   cy = GetSystemMetrics(SM_CYICON)
   hIconLarge = LoadImageAsString( _
         App.hInstance, sIconResName, _
         IMAGE_ICON, _
         cx, cy, _
         LR_SHARED)
   If (bSetAsAppIcon) Then
      SendMessageLong lhWndTop, WM_SETICON, ICON_BIG, hIconLarge
   End If
   SendMessageLong hwnd, WM_SETICON, ICON_BIG, hIconLarge
   
   cx = GetSystemMetrics(SM_CXSMICON)
   cy = GetSystemMetrics(SM_CYSMICON)
   hIconSmall = LoadImageAsString( _
         App.hInstance, sIconResName, _
         IMAGE_ICON, _
         cx, cy, _
         LR_SHARED)
   If (bSetAsAppIcon) Then
      SendMessageLong lhWndTop, WM_SETICON, ICON_SMALL, hIconSmall
   End If
   SendMessageLong hwnd, WM_SETICON, ICON_SMALL, hIconSmall
End Sub


