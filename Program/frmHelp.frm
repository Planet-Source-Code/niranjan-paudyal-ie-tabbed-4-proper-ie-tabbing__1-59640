VERSION 5.00
Begin VB.Form frmHelp 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Help - Tabbed IE: Niranjan Paudyal"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   337
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   673
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cOK 
      Caption         =   "Ok"
      Height          =   375
      Left            =   8400
      TabIndex        =   3
      Top             =   4500
      Width           =   1455
   End
   Begin VB.PictureBox pB 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4095
      Left            =   120
      ScaleHeight     =   4095
      ScaleWidth      =   9855
      TabIndex        =   0
      Top             =   120
      Width           =   9855
      Begin VB.VScrollBar VS 
         Height          =   3855
         Left            =   9600
         TabIndex        =   2
         Top             =   120
         Width           =   255
      End
      Begin VB.PictureBox pH 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   7440
         Left            =   0
         Picture         =   "frmHelp.frx":0000
         ScaleHeight     =   496
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   636
         TabIndex        =   1
         Top             =   0
         Width           =   9540
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000D&
      BorderWidth     =   3
      X1              =   8
      X2              =   664
      Y1              =   288
      Y2              =   288
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(copy to clipboard)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   195
      Left            =   3360
      MousePointer    =   99  'Custom
      TabIndex        =   6
      ToolTipText     =   "Copy e-mail address to clipboard"
      Top             =   4575
      Width           =   1350
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "nirpaudyal@hotmail.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   960
      MousePointer    =   99  'Custom
      TabIndex        =   5
      ToolTipText     =   "E-mail me now!"
      Top             =   4545
      Width           =   2415
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "e-mail:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   4
      Top             =   4545
      Width           =   660
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000010&
      FillColor       =   &H80000016&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   4440
      Width           =   9855
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Label2.MouseIcon = LoadResPicture("CURSOR_0", 2)
    Label3.MouseIcon = LoadResPicture("CURSOR_0", 2)
    pH.Move 0, 0
    VS.Move pB.ScaleWidth - VS.Width, 0, VS.Width, pB.ScaleHeight
    
    VS.Min = 0
    VS.Max = pH.Height - pB.ScaleHeight
    VS.SmallChange = 20
    VS.LargeChange = 60
End Sub

Private Sub Label2_Click()
    If ShellExecute(Me.hwnd, "open", "mailto:nirpaudyal@hotmail.com", "", 0, SW_SHOW) <= 32 Then
        Label3_Click
        MsgBox "Unable to run your default mail client. The e-mail address has been copied to the clipboard.", vbExclamation, "Problem with e-mail client"
    End If
End Sub

Private Sub Label3_Click()
    Clipboard.Clear
    Clipboard.SetText "nirpaudyal@hotmail.com"
End Sub

Private Sub VS_Change()
    pH.Move 0, -VS.Value
End Sub

Private Sub VS_Scroll()
    VS_Change
End Sub
