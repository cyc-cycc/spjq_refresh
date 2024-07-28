VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   3495
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   5  'Size
   ScaleHeight     =   2790
   ScaleWidth      =   3495
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "退出"
      Height          =   375
      Left            =   120
      MousePointer    =   1  'Arrow
      TabIndex        =   10
      Top             =   2280
      Width           =   3255
   End
   Begin VB.CommandButton command2 
      BackColor       =   &H80000004&
      Caption         =   "开始刷屏（请手动复制刷屏内容）"
      Height          =   495
      Left            =   120
      MousePointer    =   1  'Arrow
      TabIndex        =   9
      Top             =   1680
      Width           =   3255
   End
   Begin VB.Frame sz 
      BackColor       =   &H80000004&
      Caption         =   "数据设置"
      ForeColor       =   &H80000007&
      Height          =   1455
      Left            =   120
      MousePointer    =   1  'Arrow
      TabIndex        =   11
      Top             =   120
      Width           =   3255
      Begin VB.CommandButton Command4 
         Caption         =   "-"
         Height          =   255
         Left            =   2280
         TabIndex        =   8
         Top             =   1080
         Width           =   375
      End
      Begin VB.CommandButton Command3 
         Caption         =   "+"
         Height          =   255
         Left            =   1800
         TabIndex        =   7
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   960
         MaxLength       =   4
         MousePointer    =   3  'I-Beam
         TabIndex        =   6
         Text            =   "50"
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox wai 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   960
         MaxLength       =   4
         MousePointer    =   3  'I-Beam
         TabIndex        =   3
         Text            =   "0"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox num 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   960
         MousePointer    =   3  'I-Beam
         TabIndex        =   0
         Text            =   "0"
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton Command13 
         BackColor       =   &H80000004&
         Caption         =   "-"
         Height          =   255
         Left            =   2280
         TabIndex        =   5
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton Command12 
         BackColor       =   &H80000004&
         Caption         =   "+"
         Height          =   255
         Left            =   1800
         TabIndex        =   4
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H80000004&
         Caption         =   "-"
         Height          =   255
         Left            =   2280
         TabIndex        =   2
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H80000004&
         Caption         =   "+"
         Height          =   255
         Left            =   1800
         TabIndex        =   1
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "毫秒"
         Height          =   255
         Left            =   2760
         TabIndex        =   18
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label5 
         Caption         =   "刷屏延迟："
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         Caption         =   "次"
         Height          =   255
         Left            =   2880
         TabIndex        =   16
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         Caption         =   "秒"
         Height          =   255
         Left            =   2880
         TabIndex        =   15
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000004&
         Caption         =   "刷屏间隔："
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000004&
         Caption         =   "刷屏次数："
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000A&
      X1              =   0
      X2              =   4440
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000A&
      Caption         =   "V0.0.0"
      Height          =   255
      Left            =   120
      MousePointer    =   1  'Arrow
      TabIndex        =   14
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Menu tool 
      Caption         =   "工具(&T)"
      Begin VB.Menu others 
         Caption         =   "其它刷屏模式(&O)"
         Begin VB.Menu random_gap 
            Caption         =   "随机间隔"
            Shortcut        =   ^{F1}
         End
         Begin VB.Menu fastest 
            Caption         =   "高速刷屏"
            Shortcut        =   ^{F2}
         End
      End
      Begin VB.Menu del 
         Caption         =   "删除所有保存的数据(&D)"
         Shortcut        =   {F5}
      End
      Begin VB.Menu word 
         Caption         =   "文本框(&W)"
         Shortcut        =   {F6}
      End
   End
   Begin VB.Menu help 
      Caption         =   "帮助(&H)"
      Begin VB.Menu about 
         Caption         =   "关于(&A)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu set 
         Caption         =   "设置(&S)"
         Shortcut        =   {F2}
      End
      Begin VB.Menu bar0 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu helpw 
         Caption         =   "帮助文档(&T)"
         Visible         =   0   'False
      End
      Begin VB.Menu bar1 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "退出(&E)"
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu test 
      Caption         =   "未完成版本，不要编译和外传"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub about_Click()
    frmAbout.Show
End Sub
Private Sub Command1_Click()
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
Private Sub Command10_Click()
    Me.num.Text = Me.num.Text + 1
    If Me.num.Text <= 0 Then
        Me.num.Text = 1
    End If
End Sub
Private Sub Command11_Click()
    Me.num.Text = Me.num.Text - 1
    If Me.num.Text <= 0 Then
        Me.num.Text = 1
    End If
End Sub
Private Sub Command12_Click()
    Me.wai.Text = Me.wai.Text + 0.1
    If Me.wai.Text < 0 Then
        Me.wai.Text = 0
    End If
End Sub
Private Sub Command13_Click()
    Me.wai.Text = Me.wai.Text - 0.1
    If Me.wai.Text < 0 Then
        Me.wai.Text = 0
    End If
End Sub
Private Sub Command3_Click()
    Me.Text2.Text = Me.Text2.Text + 10
    If Me.Text2.Text < 0 Then
        Me.Text2.Text = 0
    End If
End Sub
Private Sub Command4_Click()
    Me.Text2.Text = Me.Text2.Text - 10
    If Me.Text2.Text < 0 Then
        Me.Text2.Text = 0
    End If
End Sub
Private Sub del_Click()
    ans = MsgBox("应用程序将被关闭！", 4096 + 64 + vbYesNo + vbDefaultButton2, "引导")
    If ans = vbYes Then
        DeleteSetting App.ProductName
        End
    End If
End Sub
Private Sub exit_Click()
    Unload Me
End Sub
Private Sub fastest_Click()
    gif_form1 = False
    Me.Hide
    quickest.Show
End Sub
Private Sub Form_Activate()
    Dim rtn As Long
    rtn = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes hwnd, 0, aero, LWA_ALPHA
    If gif = True Then
        If gif_form1 <> True Then
            Call rgn_gif(Me, 25, 3, 3, 3)
            gif_form1 = True
        End If
    End If
    Call rgnform(Me, rou, rou, 25, 3, 3, 3)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    gif_form1 = False
    SaveSetting App.ProductName, "settings", "left", Me.left
    SaveSetting App.ProductName, "settings", "top", Me.top
    SaveSetting App.ProductName, "settings", "aero", aero
    SaveSetting App.ProductName, "settings", "round", rou
    SaveSetting App.ProductName, "settings", "splash", spl
    SaveSetting App.ProductName, "settings", "splash_time", splt
    SaveSetting App.ProductName, "settings", "save_form_location_yn", sfly
    SaveSetting App.ProductName, "settings", "save_form_location", sfl
    SaveSetting App.ProductName, "settings", "gif", gif
    SaveSetting App.ProductName, "number", "num", Me.num.Text
    SaveSetting App.ProductName, "number", "wai", Me.wai.Text
    SaveSetting App.ProductName, "number", "wai_enter", Me.Text2.Text
    DeleteObject outrgn
    End
End Sub
Private Sub Command2_Click()
    gif_form1 = False
    Me.Hide
    Form2.Show
End Sub
Private Sub Form_Load()
    Me.num.Text = GetSetting(App.ProductName, "number", "num", "0")
    Me.wai.Text = GetSetting(App.ProductName, "number", "wai", "0")
    Me.Text2.Text = GetSetting(App.ProductName, "number", "wai_enter", "50")
    If sfly = 1 Then
        Me.left = GetSetting(App.ProductName, "settings", "left", Screen.Width / 2 - Me.Width / 2)
        Me.top = GetSetting(App.ProductName, "settings", "top", Screen.Height / 2 - Me.Height / 2)
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
        Form1.left = Form1.left + (X - sx)
        Form1.top = Form1.top + (Y - sy)
    End If
End Sub
Private Sub num_LostFocus()
    Me.num.Text = Fix(Me.num.Text)
    If Me.num.Text <= 0 Then
        Me.num.Text = 1
    End If
End Sub
Private Sub random_gap_Click()
    gif_form1 = False
    Me.Hide
    gap.Show
End Sub
Private Sub set_Click()
    settings.Show
End Sub
Private Sub Text2_LostFocus()
    If Me.Text2.Text < 0 Then
        Me.Text2.Text = 0
    End If
End Sub
Private Sub wai_LostFocus()
    If Me.wai.Text < 0 Then
        Me.wai.Text = 0
    End If
End Sub
Private Sub word_Click()
    wordbox.Show
End Sub
