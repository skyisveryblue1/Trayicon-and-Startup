VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00FAEBEB&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Demo for TrayIcon and Register/Unregister Startup"
   ClientHeight    =   1935
   ClientLeft      =   6570
   ClientTop       =   2280
   ClientWidth     =   7935
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   7935
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   6240
      ScaleHeight     =   795
      ScaleWidth      =   1155
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton cmdRegisterToStartup 
      Appearance      =   0  'Flat
      Caption         =   "Register To Startup"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   240
      Width           =   2295
   End
   Begin VB.CommandButton cmdRemoveFromStartup 
      Appearance      =   0  'Flat
      Caption         =   "Remove From Startup"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   840
      Width           =   2295
   End
   Begin VB.CommandButton cmdQuit 
      Appearance      =   0  'Flat
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub cmdRegisterToStartup_Click()
    RegisterStartup
End Sub

Private Sub cmdRemoveFromStartup_Click()
    RemoveStartup
End Sub

Private Sub Form_Load()
    'Setup initial Tray Icon
    T.cbSize = Len(T)
    T.hWnd = Picture1.hWnd
    T.uId = 1&
    T.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    T.ucallbackMessage = WM_MOUSEMOVE
    T.hIcon = Me.Icon
    T.szTip = "Recent" & Chr$(0)
    Shell_NotifyIcon NIM_ADD, T
   
    'Hide this form
    Me.Hide
End Sub


Private Sub Form_Unload(Cancel As Integer)
    'Unload this form. Important: always end with "unload me".
    T.cbSize = Len(T)
    T.hWnd = Picture1.hWnd
    T.uId = 1&
    Shell_NotifyIcon NIM_DELETE, T
    End
End Sub


Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Static State As Boolean
 Static Popped As Boolean
 Static Msg As Long
 
    Msg = X / Screen.TwipsPerPixelX
    If Popped = False Then
        Popped = True
        Select Case Msg
            Case WM_LBUTTONDOWN
                frmMain.Show
        End Select
        'OK to popup again
        Popped = False
    End If
End Sub
