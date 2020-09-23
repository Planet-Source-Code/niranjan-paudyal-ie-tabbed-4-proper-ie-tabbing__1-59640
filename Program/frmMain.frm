VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Internet Explorer - No page"
   ClientHeight    =   6750
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   9900
   FontTransparent =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   450
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   660
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin ComCtl3.CoolBar CB 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9900
      _ExtentX        =   17463
      _ExtentY        =   767
      BandCount       =   1
      FixedOrder      =   -1  'True
      BandBorders     =   0   'False
      VariantHeight   =   0   'False
      OLEDropMode     =   1
      _CBWidth        =   9900
      _CBHeight       =   435
      _Version        =   "6.0.8169"
      Child1          =   "TS"
      MinWidth1       =   569
      MinHeight1      =   25
      Width1          =   656
      NewRow1         =   0   'False
      BandStyle1      =   1
      AllowVertical1  =   0   'False
      Begin ComctlLib.TabStrip TS 
         Height          =   375
         Left            =   30
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   30
         Width           =   9840
         _ExtentX        =   17357
         _ExtentY        =   661
         TabWidthStyle   =   2
         _Version        =   327682
         BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
            NumTabs         =   1
            BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   ""
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin VB.Timer Tfunctions 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   480
   End
   Begin VB.PictureBox pParent 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   1320
      ScaleHeight     =   615
      ScaleWidth      =   615
      TabIndex        =   2
      Top             =   480
      Width           =   615
   End
   Begin ComctlLib.ImageList IMGl 
      Left            =   600
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   19
      ImageHeight     =   19
      MaskColor       =   -2147483633
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":728A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mTabs 
      Caption         =   "Tabs"
      Visible         =   0   'False
      Begin VB.Menu mNew 
         Caption         =   "New"
      End
      Begin VB.Menu mClose 
         Caption         =   "Close"
      End
      Begin VB.Menu mOptionsSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mCaot 
         Caption         =   "Close all other tabs"
      End
      Begin VB.Menu mOptionsSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mOptions 
         Caption         =   "Options"
         Begin VB.Menu mRunInBackground 
            Caption         =   "Run in background"
            Checked         =   -1  'True
         End
         Begin VB.Menu mRunAtStartup 
            Caption         =   "Run at startup"
         End
         Begin VB.Menu mShowUnopenedTabIcon 
            Caption         =   "Show unopened tab icon"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mOptionsSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mHelp 
         Caption         =   "Help"
      End
      Begin VB.Menu mAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mOptionsSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mExit1 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Please compile the program before running!
'I warn in advance, please compile before running!

'When you run the program, it will not show itself unless there are IE windows open.
'Otherwise, launch an IE window and it will be tabbed.

'Enjoy.
'Niranjan Paudyal

Option Explicit
Public Hw As Long                       'This is the hWnd of the currently visible IE window

Private LastR As RECT                   'Last size of this form
Private LastIE As RECT                  'Last size of HW
Private ForceClose As Boolean           'To force close the App regardless of the the options!
Private InternetExtraTitle As String    'From the REG, extra text added to the end of the title bar (normally 'Mocrosoft Internet Explorer"
Private LastWindowCaption As String     'The title of Hw (current internet window) minus InternetExtraTitle
Private LastTabPopup As Long            'The last tab that the rightbutton was clicked on (Used for the popup menu)
Private TotalWindowsTrapped As Long     'How many windows have been trapped by the program? (gives a total number of windows trapped since program ran, not the current number trapped)
Private IEpath As String                'IE path (where installed on computer)

Private Const WindowClass = "IEFrame"   'Class of the window to bring in.

Private Function FitTabs()
    Dim TmW As Long, TnW As Long, TMaxW
    'Function will resize the Tab width according to the current with of the
    'form and the number of tabs
    If TS.Tabs.Count > 0 Then
        TmW = 300
        TMaxW = 355 * Screen.TwipsPerPixelX
        TnW = (Me.ScaleWidth * Screen.TwipsPerPixelX - 200) / TS.Tabs.Count
        If TnW < TmW Then TnW = TmW
        If TnW > TMaxW Then TnW = TMaxW
        TS.TabFixedWidth = TnW
    End If
End Function

Private Function SaveSettings() As Boolean
    'Function will save the program settings settings in the App directory
    'If it can not do so, it will not produce any error message here!
    '-just returns false.
    On Error GoTo 1
    Open App.Path & "\Settings.txt" For Output As #1
    Print #1, mRunInBackground.Checked 'Run in background?
    Print #1, mShowUnopenedTabIcon.Checked 'Show icons for tabs that have not been visited?
    Print #1, mRunAtStartup.Checked 'Load at windows startup?
    Print #1, IsZoomed(Me.hwnd)
    Print #1, Me.Left
    Print #1, Me.Top
    Print #1, Me.Width
    Print #1, Me.Height
    
    SaveSettings = True 'Success.
1:  Close #1
End Function

Private Sub FindIEWindows()
    'This procedure looks for new IE windows and adds it to the form (frmMain)
    'It creates a new tab and resizes them accordingly
    Dim H As Long, x As Long
    Dim WC As String
    Dim DontAdd As Boolean
    
    H = FindWindow(WindowClass, vbNullString) 'Find the first IE window.
    Do
        If H <> 0 Then 'If a window was found then...
            'Get the class name of the window
            WC = String(Len(WindowClass) + 1, " ")
            GetClassName H, WC, Len(WC)
            WC = Mid(WC, 1, Len(WC) - 1)
            If WC = WindowClass Then 'If the window was of the class that is being searched for...
                For x = 1 To TS.Tabs.Count 'Make sure it has not been added already
                    If H = CLng(Mid(TS.Tabs(x).Key, 2)) Then
                        DontAdd = True
                        Exit For
                    End If
                Next x
                
                If Not DontAdd Then 'If it is ok to add this form then
                    Tfunctions.Enabled = False 'Disable the timer to prevent serching for windows while a new window is being added.
                    While IsWindowVisible(Me.hwnd) = 0
                        ShowWindow Me.hwnd, SW_SHOW 'if the main window is not visible then show it!
                    Wend
                    LockWindowUpdate GetDesktopWindow 'Stop desktop redraw (hides flicker of a new window).
                    Call BringInIEWindow(H) 'This will bring in the IE window to this form
                    TotalWindowsTrapped = TotalWindowsTrapped + 1 'Update windows counter
                    TS.Tabs.Add , CStr("i" & H), "" 'Add the tab (at the end of all the other tabs)
                    Call FitTabs 'Call to best fit all the tabs in the tab strip.
                    If TS.Tabs.Count = 1 Then 'If this is the only IE window then...
                        ActivateTab 'Activate it.
                    Else
                        If mShowUnopenedTabIcon.Checked Then TS.Tabs(TS.Tabs.Count).Image = 1 'Otherwise, dont activate it, but show the Un-visited icon on the tab if the user has set it on.
                    End If
                    AppActivate Me.Caption 'Make this form the focus.
                    If IsIconic(Me.hwnd) Then 'If this window is minimized, then focus it and flash it.
                        ShowWindow Me.hwnd, 9
                        FlashWindow Me.hwnd, 2
                    End If
                    LockWindowUpdate 0 'Unloack desktop window to allow redrawing.
                    Tfunctions.Enabled = True 'Switch on the timer to deal with windows.
                Else
                    DontAdd = False 'reset to allow searching for new windows.
                End If
            End If
            H = GetNextWindow(H, GW_HWNDNEXT) 'Find the next window
        Else
            Exit Do
        End If
    Loop
End Sub

Private Function BringInIEWindow(Wnd As Long)
    'This procedure adds the IE window found by FindIEWindows sub
    Dim C As Long, OldT As Long, RetVal As Long
    Dim R As RECT
    
    SetWindowStyle Wnd, False, WS_CHILD, True, False 'Make the IE window a child window
    If TS.Tabs.Count = 0 Then
        SetParent Wnd, Me.hwnd 'If this is the only IE window (therefor, first!) then display it on the form
    Else
        SetParent Wnd, pParent.hwnd 'If not the first, then add put into the picture box.
    End If
End Function
Public Sub Wm_ResizeMe()
    'This sub is called form mFunctions module when a change in window size is detected
    Dim R As RECT
    
    GetWindowRect Me.hwnd, R
    'Make sure that the window size has been changed (not just moved)
    If (R.Right - R.Left) = (LastR.Right - LastR.Left) And (R.Bottom - R.Top) = (LastR.Bottom - LastR.Top) Then Exit Sub
    'If all is fine, resize the IE window to fit the form (frmMain)
    SetWindowPos Hw, 0, 0, CB.Height, Me.ScaleWidth, Me.ScaleHeight - CB.Height, SWP_NOOWNERZORDER Or SWP_NOACTIVATE
    GetWindowRect Me.hwnd, LastR
    FitTabs
End Sub

Private Sub Form_Initialize()
    Dim I As Long
    Dim V As String

    'Check previous instance
    If App.PrevInstance = True Then
        End 'Close this instance
    Else
        'Check Windows version
        If OSVersion >= 2004 Then 'if Win 2000 or later then...
            IEpath = InternetExplorerPath 'Check that IE is installed
            If IEpath = "" Then
                MsgBox "Microsoft Internet Explorer could not be found in this computer. Please make sure that it is installed.", vbCritical, "IE not found."
                End
            Else
                If IEVersionShort < 5 Then 'Check the version of IE (make sure its greater then v5
                    MsgBox "In order for the tabbed interface to work, Internet Explorer must be of version above 5." & vbNewLine & "You are currently running :" & IEVersionLong, vbCritical, "Incorrect IE version."
                    End
                Else
                    'All is fine, Skin the controls according to XP skin (in 2000, this is ignored)
                    InitCommonControls
                End If
            End If
        Else
            MsgBox "In order for the tabbed interface to work, version of Windows has to be 2000 or later.", vbCritical, "Windows version not supported."
            End
        End If
    End If
    
    'Get rid of pre-existing tabs
    If TS.Tabs.Count > 0 Then
        For I = 1 To TS.Tabs.Count
            TS.Tabs.Remove (I)
        Next I
    End If
    
    pParent.Move -90, -90, 40, 40 'Hide the picture box that traps the IE windows
    
    'This gets the extra string added to every page title in IE.
    InternetExtraTitle = ReadRegistry(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer\Main\", "Window Title")
    If InternetExtraTitle = "\error\" Then InternetExtraTitle = "Microsoft Internet Explorer" 'If error is returned, then it must be the default title.
    InternetExtraTitle = " - " & InternetExtraTitle
    
    TS.ImageList = IMGl 'Assign an Image list to the tab control
    
    'Set up classing to to detect window changes
    Set CSubClsApp = New cSubclass
    CSubClsApp.hwnd = Me.hwnd
    Call CSubClsApp.EnableSubclass
    
    'Load settings. If the settings file is not found, then it is ignored and default settings are then used
    On Error GoTo 1
    Open App.Path & "\Settings.txt" For Input As #1
    Line Input #1, V
    mRunInBackground.Checked = V
    Line Input #1, V
    mShowUnopenedTabIcon.Checked = V
    Line Input #1, V
    mRunAtStartup.Checked = V
    
    'Window size and position
    Line Input #1, V
    If V <> 1 Then 'If window was closed while it was not maximized
        Line Input #1, V
        Me.Left = V
        Line Input #1, V
        Me.Top = V
        Line Input #1, V
        Me.Width = V
        Line Input #1, V
        Me.Height = V
    Else
        'If maximized then show the window in the Max state, then hide it again
        LockWindowUpdate GetDesktopWindow 'Hide flicker
            ShowWindow Me.hwnd, 3
            ShowWindow Me.hwnd, 0
        LockWindowUpdate 0
        'Go through tht rest of the file
        Line Input #1, V
        Line Input #1, V
        Line Input #1, V
        Line Input #1, V
    End If

    
1:
    Close #1
    'Get everything rolling.
    Tfunctions.Enabled = True
End Sub

Private Sub Form_Load()
    SetIcon Me.hwnd, "AAA" 'Assign the high colour icon form resource to the form.
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim x As Long, OldT As Long, RetVal  As Long
    
    Tfunctions.Enabled = False 'Disable timer so it doesn't search for windows
1:
    LockWindowUpdate Me.hwnd 'hides flicker of closing windows
        For x = 1 To TS.Tabs.Count  'Co through all tabs
            If IsWindow(CLng(Mid(TS.Tabs(x).Key, 2))) <> 0 Then 'If window is there then
                 CloseHwnd Mid(TS.Tabs(x).Key, 2)   'send the close window message
            End If
        Next
        
        'Next, we check to make sure that the windows have indeed closed
        OldT = GetTickCount
        For x = 1 To TS.Tabs.Count  'Go through all tabs
2:
            DoEvents
            If IsWindow(Mid(TS.Tabs(x).Key, 2)) <> 0 Then   'If the window is there then...
                If GetTickCount - OldT > 2000 Then
                    'If it has been there for 2000ms then
                    LockWindowUpdate 0
                    'Show message to give option of waiting, force closing or just ending all windows
                    RetVal = MsgBox("One or more IE windows could not be closed. They have already been given 2 seconds." & vbNewLine & "If you want to attempt closing them again, please press Retry." & vbNewLine & "If you want to wait and see if they close by themself, press Ignore." & vbNewLine & "If you would like to force them to close, then press Abort.", vbAbortRetryIgnore, "Unable to close IE windows")
                    Select Case RetVal
                        Case vbRetry
                            GoTo 1
                        Case vbIgnore
                            OldT = GetTickCount
                            LockWindowUpdate GetDesktopWindow
                            GoTo 2
                        Case vbAbort
                            Call CSubClsApp.DisableSubclass
                            Set CSubClsApp = Nothing
                            End
                    End Select
                Else
                    'Continue looking at this window
                    GoTo 2
                End If
            Else
                'Move onto the next window, reset timer
                OldT = GetTickCount
            End If
        Next

        'If running in background, then just hide frmMain
        If mRunInBackground.Checked And (ForceClose = False) Then
            ShowWindow Me.hwnd, SW_HIDE
            Tfunctions.Enabled = True
            Cancel = 1
        Else
            'Otherwise, save settings and unload
            SaveSettings
            Call CSubClsApp.DisableSubclass
            Set CSubClsApp = Nothing
        End If
    LockWindowUpdate 0
End Sub

Private Sub mAbout_Click()
    frmAbout.Show 1
End Sub

Private Sub mCaot_Click() 'Close all other tabs
    Dim I As Long
    Tfunctions.Enabled = False
    'Go through all tabs, close IE window if not in the selected tab
    For I = 1 To TS.Tabs.Count
        If I <> LastTabPopup Then
            CloseHwnd Mid(TS.Tabs(I).Key, 2) 'send the close window message
        End If
    Next I
    Tfunctions.Enabled = True
End Sub

Private Sub mClose_Click()
    CloseHwnd Mid(TS.Tabs(LastTabPopup).Key, 2) 'send the close window message to the selected tab window
End Sub

Private Sub mExit1_Click()
    'forcefully closes the program (from background as well)
    If mRunInBackground.Checked = True Then
        If MsgBox("Please note that by pressing 'Yes', you will close all Internet Explorer processes and Tabbed IE will no longer run in the background." & vbNewLine & "If you will like to close all the tabs and keep the Tabbed IE program running in the background, you must close by using the 'X' button on the window border." & vbNewLine & "Are you sure you wish to continue?", vbYesNo Or vbQuestion, "Exit from background?") = vbNo Then
            Exit Sub
        End If
    End If
    
    ForceClose = True
    Unload Me
    ForceClose = False
End Sub
Private Sub mExit2_Click()
    mExit1_Click
End Sub

Private Sub mHelp_Click()
    frmHelp.Show 1
End Sub

Private Sub mNew_Click()
    Dim Ret As Long, XX As Long
    'Open IE using the command line
    Ret = ShellExecute(hwnd, "open", "iexplore.exe", "", "", SW_HIDE)
    'If that doesn't work, try opening by giving the path
    If Ret <= 32 Then
        Ret = ShellExecute(hwnd, "open", IEpath & "\iexplore.exe", "", "", SW_HIDE)
        If Ret <= 32 Then
            'Otherwise, try to open a blank document
            If ShellExecute(hwnd, "open", "about:blank", "", "", SW_HIDE) <= 32 Then
                'If all else fails then...
                MsgBox "Could not open a new window.", vbExclamation, "Unknown problem."
            End If
        End If
    End If
End Sub

Private Sub mRestore_Click()
    mNew_Click
End Sub

Private Sub mRunAtStartup_Click()
    'This adds or removes the start up entery in the registery
    mRunAtStartup.Checked = Not mRunAtStartup.Checked
    If mRunAtStartup.Checked Then
        Call WriteRegistry(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run\", "Tabbed IE-Niranjan Paudyal", ValString, App.Path & "\" & App.EXEName & ".exe")
    Else
        Call DeleteValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run\", "Tabbed IE-Niranjan Paudyal")
    End If
    SaveSettings
End Sub
Private Sub mRunInBackground_Click()
    mRunInBackground.Checked = Not mRunInBackground.Checked
    SaveSettings
End Sub
Private Sub mShowUnopenedTabIcon_Click()
    mShowUnopenedTabIcon.Checked = Not mShowUnopenedTabIcon.Checked
    SaveSettings
End Sub

Private Sub Tfunctions_Timer()
    'go through each tabs making sure that the IE windows do exist
    Dim TPC As Boolean, IsCurrHWND As Long
    Dim XX As Long, NewHW As Long
    Dim WT As String
    Dim R As RECT
    Dim C As Long
        
        'This is to check if the window needs to be resized..
        'This checks if the browser window has resized
        If IsWindow(Hw) Then
            GetWindowRect Hw, R
            If ((R.Bottom - R.Top) <> (LastIE.Bottom - LastIE.Top) Or (R.Right - R.Left) <> (LastIE.Right - LastIE.Left)) Then
                ActivateTab
                GetWindowRect Hw, LastIE
            End If
        End If
2:
        
    If IsWindow(Hw) Then
        'Check the latest title of Hw is displayed as a window caption
        WT = n_GetWindowText(Hw)
        If LastWindowCaption <> WT Then
            Me.Caption = WT
            LastWindowCaption = WT
        End If
    Else
    'If the current window is not there anymore (closed)
        If TS.Tabs.Count = 0 Then
        'If there are no more windows left then...
            If LastWindowCaption <> "Internet Explorer - No page" Then
                LastWindowCaption = "Internet Explorer - No page"
                Me.Caption = "Internet Explorer - No page"
                If TotalWindowsTrapped > 0 Then
                    If mRunInBackground.Checked Then
                        While IsWindowVisible(Me.hwnd)
                            ShowWindow Me.hwnd, SW_HIDE
                        Wend
                    Else
                        Unload Me
                        Exit Sub
                    End If
                End If
            End If
            
        End If
    End If
    
1:
    
    Dim M As Long
    For XX = 1 To TS.Tabs.Count
        NewHW = CLng(Mid(TS.Tabs(XX).Key, 2))
        If NewHW = Hw Then IsCurrHWND = XX
        If IsWindow(NewHW) = 0 Then

            'Remove tabs for any non-existing windows
            If NewHW = Hw Then IsCurrHWND = XX
            TS.Tabs.Remove (XX)
            TPC = True
            GoTo 1
        Else
            'If the window is still there, make sure the tab caption is its latest title...
            WT = n_GetWindowText(NewHW)
            If Len(WT) >= Len(InternetExtraTitle) Then WT = Mid(WT, 1, Len(WT) - Len(InternetExtraTitle))
            WT = RemoveUnderscore(WT)
            'WT = IIf(Len(WT) > 20, Mid(WT, 1, 30) & "...", WT)
            If TS.Tabs(XX).Caption <> WT Then
                TS.Tabs(XX).Caption = WT
                TS.Tabs(XX).ToolTipText = n_GetWindowText(NewHW)
            End If
            
        End If
    Next XX
    
    If TPC Then
        
        FitTabs
        'LockWindowUpdate 0
        If IsCurrHWND <> 0 Then
            If TS.Tabs.Count > 0 Then
                If IsCurrHWND = 1 Then
                    XX = 1
                Else
                    XX = IsCurrHWND
                End If
                If XX <= 0 Then XX = 1
                If XX > TS.Tabs.Count Then XX = TS.Tabs.Count
                TS.Tabs(XX).Selected = True
                ActivateTab    ' If a tab was removed then go to a new tab
            End If
        Else
            TS.Tabs(CStr("i" & Hw)).Selected = True
            ActivateTab    ' If a tab was removed then go to a new tab
        End If

    End If
    
    FindIEWindows   'Search for new IE windows
End Sub


Private Sub ActivateTab()
    Dim NewHW As Long, C As Long
    If TS.Tabs.Count = 0 Then Exit Sub 'if there are no tabs, then do nothing!
                
    Screen.MousePointer = vbHourglass
                
    Tfunctions.Enabled = False
    NewHW = Mid(TS.SelectedItem.Key, 2) 'the window associated with the tab is stored as its Key
                
    LockWindowUpdate GetDesktopWindow 'Reduce flicker
                    
    TS.SelectedItem.Image = 0 'Remove any tab icons
    SetWindowStyle NewHW, False, WS_CAPTION, False, False     'Hide the caption (window title)
    SetWindowStyle NewHW, False, WS_THICKFRAME, False, False  'Get rid of the borders (so as to lock resizing)
    SetWindowStyle NewHW, False, WS_DLGFRAME, False, False
    C = FindWindowEx(NewHW, 0, "WorkerW", "")
    C = FindWindowEx(C, 0, "ReBarWindow32", "")
    SetWindowStyle C, False, WS_BORDER, False, False 'This has to be done sometimes (e.g after comming back from full screen)
                    
    SendMessage NewHW, WM_NCACTIVATE, 1, 0 'Activate the clicked tab window
    SendMessage NewHW, WM_ACTIVATE, 2, 0
    SendMessage NewHW, WM_ACTIVATE, 2, 0
    If NewHW <> Hw Then SetParent Hw, pParent.hwnd 'set the IE windows parent as frmMain (not the picture box)
    If GetParent(NewHW) <> Me.hwnd Then SetParent NewHW, Me.hwnd
    ShowWindow NewHW, 3 'Maximize the window
    SetWindowPos NewHW, 0, 0, CB.Height, Me.ScaleWidth, Me.ScaleHeight - CB.Height, 0
                    
    Hw = NewHW
    LockWindowUpdate 0
    Tfunctions.Enabled = True
    Screen.MousePointer = vbNormal
End Sub

Private Sub TS_Click()
    ActivateTab
End Sub

Private Sub TS_MouseDown(button As Integer, Shift As Integer, x As Single, y As Single)
    Dim Indexx As Long, HwX As Long
    Dim Ret As String
    'This sub pops up the right click menu
    Indexx = 1 + (x \ TS.TabFixedWidth)
    If TS.Tabs.Count >= Indexx Then
        HwX = CLng(Mid(TS.Tabs(Indexx).Key, 2))
        
        Select Case button
            Case 1
            Case 2 'Right button pressed
                'The contents of the menu depends on, how many tabs are there.
                mClose.Visible = True
                LastTabPopup = Indexx
                mOptionsSep1.Visible = (TS.Tabs.Count <> 1)
                mCaot.Visible = (TS.Tabs.Count <> 1)
                mRunAtStartup.Checked = (ReadRegistry(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run\", "Tabbed IE-Niranjan Paudyal") = App.Path & "\" & App.EXEName & ".exe")
                PopupMenu mTabs 'Show the menu
            Case Else
                CloseHwnd Mid(TS.Tabs(Indexx).Key, 2) 'send the close window message (this is the middle or any other button)
        End Select

    End If
End Sub


Private Sub CB_MouseDown(button As Integer, Shift As Integer, x As Single, y As Single)
    If button = 2 Then
        'Right click pressed else where apart from the tabs
        mClose.Visible = False
        mOptionsSep1.Visible = False
        mCaot.Visible = False
        PopupMenu mTabs
    End If
End Sub

Private Sub TS_OLEDragDrop(Data As ComctlLib.DataObject, Effect As Long, button As Integer, Shift As Integer, x As Single, y As Single)
    'Drag and drop link code
    If Data.GetFormat(1) = True Then
        Dim XX As Long, Indexx As Long, HwX As Long
        Dim URL As String
        Indexx = 1 + (x \ TS.TabFixedWidth)
        URL = Data.GetData(1)
        If TS.Tabs.Count >= Indexx Then
            HwX = CLng(Mid(TS.Tabs(Indexx).Key, 2))
                    
            XX = FindWindowEx(HwX, 0, "WorkerW", "")
            XX = FindWindowEx(XX, 0, "ReBarWindow32", "")
            XX = FindWindowEx(XX, 0, "ComboBoxEx32", "")
            XX = FindWindowEx(XX, 0, "ComboBox", "")
            XX = FindWindowEx(XX, 0, "Edit", "")
            If XX = 0 Then
                MsgBox "Could not open link on this tab.", vbExclamation, "Link problem."
            Else
                SendMessage XX, WM_SETTEXT, 0, ByVal URL
                SendMessage XX, WM_SETFOCUS, Hw, 0
                SendMessage XX, WM_KEYDOWN, VK_RETURN, 0
                SendMessage XX, WM_KEYUP, VK_RETURN, 0
            End If
        End If
    End If
End Sub
