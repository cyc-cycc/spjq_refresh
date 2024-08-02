VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "关于我的应用程序"
   ClientHeight    =   2925
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5910
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   5  'Size
   ScaleHeight     =   2018.886
   ScaleMode       =   0  'User
   ScaleWidth      =   5549.797
   Begin VB.TextBox Text1 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   3240
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      TabStop         =   0   'False
      Text            =   "frmAbout.frx":10CA
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   345
      Left            =   2040
      MousePointer    =   1  'Arrow
      TabIndex        =   0
      Top             =   1920
      Width           =   1020
   End
   Begin VB.Label Label3 
      Caption         =   "本程序遵循 GPL-3.0 协议"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   5655
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   5521.625
      Y1              =   1656.522
      Y2              =   1656.522
   End
   Begin VB.Line Line1 
      DrawMode        =   1  'Blackness
      X1              =   2929.842
      X2              =   2929.842
      Y1              =   82.826
      Y2              =   1490.87
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "支持系统:Windows7sp1 ~ Windows11"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      MousePointer    =   1  'Arrow
      TabIndex        =   5
      Top             =   1560
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "made by CYC"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      MousePointer    =   1  'Arrow
      TabIndex        =   4
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "重制版的刷屏机器，拥有更多的功能、更美观的界面和更强的稳定性。用于日常刷屏"
      ForeColor       =   &H00000000&
      Height          =   690
      Left            =   120
      MousePointer    =   1  'Arrow
      TabIndex        =   1
      Top             =   840
      Width           =   2925
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "应用程序标题"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   120
      MousePointer    =   1  'Arrow
      TabIndex        =   2
      Top             =   120
      Width           =   1845
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "版本"
      Height          =   225
      Left            =   120
      MousePointer    =   1  'Arrow
      TabIndex        =   3
      Top             =   480
      Width           =   1275
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    Me.Text1.Text = "运行目录：" & App.Path
    Dim rtn As Long
    rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes hwnd, 0, aero, LWA_ALPHA
    If gif = True Then
        If gif_frmabout <> True Then
            Call rgn_gif(Me, 25, 3, 3, 3)
            gif_frmabout = True
        End If
    End If
    Call rgnform(Me, rou, rou, 25, 3, 3, 3)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    gif_frmabout = False
    SaveSetting App.ProductName, "location", "left_frmabout", Me.left
    SaveSetting App.ProductName, "location", "top_frmabout", Me.top
    DeleteObject outrgn
End Sub
Private Sub cmdOK_Click()
    If gif = True Then
        aero_tmp = aero
        For i = 1 To aero_tmp / 9
            SetLayeredWindowAttributes hwnd, 0, aero_tmp, LWA_ALPHA
            aero_tmp = aero_tmp - 8
            Call Sleep(1)
        Next
    End If
    Unload Me
End Sub
Private Sub Form_Load()
    Me.Caption = "关于 " & App.ProductName
    Me.lblVersion.Caption = "版本 " & App.Major & "." & App.Minor & "." & App.Revision
    Me.lblTitle.Caption = App.ProductName
    If sfly = 1 And sfl = "all" Then
        Me.left = GetSetting(App.ProductName, "location", "left_frmabout", Screen.Width / 2 - Me.Width / 2)
        Me.top = GetSetting(App.ProductName, "location", "top_frmabout", Screen.Height / 2 - Me.Height / 2)
    Else
        Me.left = Screen.Width / 2 - Me.Width / 2
        Me.top = Screen.Height / 2 - Me.Height / 2
    End If
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    sx = X
    sy = Y
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Me.left = Me.left + (X - sx)
        Me.top = Me.top + (Y - sy)
    End If
End Sub
