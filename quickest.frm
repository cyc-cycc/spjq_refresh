VERSION 5.00
Begin VB.Form quickest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "高速刷屏"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3495
   Icon            =   "quickest.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   5  'Size
   ScaleHeight     =   3150
   ScaleWidth      =   3495
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command4 
      Caption         =   "开始刷屏（请手动复制刷屏内容）"
      Height          =   495
      Left            =   120
      MousePointer    =   1  'Arrow
      TabIndex        =   3
      Top             =   2040
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      Caption         =   "数据设置"
      Height          =   855
      Left            =   120
      MousePointer    =   1  'Arrow
      TabIndex        =   5
      Top             =   120
      Width           =   3255
      Begin VB.TextBox num 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   960
         MaxLength       =   5
         MousePointer    =   3  'I-Beam
         TabIndex        =   0
         Text            =   "0"
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H80000004&
         Caption         =   "-"
         Height          =   255
         Left            =   2280
         MousePointer    =   1  'Arrow
         TabIndex        =   2
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H80000004&
         Caption         =   "+"
         Height          =   255
         Left            =   1800
         MousePointer    =   1  'Arrow
         TabIndex        =   1
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         Caption         =   "次"
         Height          =   255
         Left            =   2880
         MousePointer    =   1  'Arrow
         TabIndex        =   7
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000004&
         Caption         =   "刷屏次数："
         Height          =   255
         Left            =   120
         MousePointer    =   1  'Arrow
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "关闭"
      Height          =   375
      Left            =   120
      MousePointer    =   1  'Arrow
      TabIndex        =   4
      Top             =   2640
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   $"quickest.frx":10CA
      ForeColor       =   &H000000FF&
      Height          =   795
      Left            =   120
      MousePointer    =   1  'Arrow
      TabIndex        =   8
      Top             =   1080
      Width           =   3240
   End
End
Attribute VB_Name = "quickest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
    Me.num.Text = Me.num.Text + 1
    If Me.num.Text <= 0 Then
        Me.num.Text = 1
    End If
End Sub
Private Sub Command3_Click()
    Me.num.Text = Me.num.Text - 1
    If Me.num.Text <= 0 Then
        Me.num.Text = 1
    End If
End Sub
Private Sub Command4_Click()
    gif_quickest = False
    Me.Hide
    special_quickest.Show
End Sub
Private Sub Form_Activate()
    Dim rtn As Long
    rtn = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes hwnd, 0, aero, LWA_ALPHA
    If gif = True Then
        If gif_quickest <> True Then
            Call rgn_gif(Me, 25, 3, 3, 3)
            gif_quickest = True
        End If
    End If
    Call rgnform(Me, rou, rou, 25, 3, 3, 3)
End Sub
Private Sub Form_Load()
    If sfly = 1 And sfl = "all" Then
        Me.left = GetSetting(App.ProductName, "location", "left_quickest", Screen.Width / 2 - Me.Width / 2)
        Me.top = GetSetting(App.ProductName, "location", "top_quickest", Screen.Height / 2 - Me.Height / 2)
    Else
        Me.left = Screen.Width / 2 - Me.Width / 2
        Me.top = Screen.Height / 2 - Me.Height / 2
    End If
    Me.num.Text = Form1.num.Text
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
Private Sub Form_Unload(Cancel As Integer)
    gif_quickest = False
    Form1.Show
    SaveSetting App.ProductName, "location", "left_quickest", Me.left
    SaveSetting App.ProductName, "location", "top_quickest", Me.top
    DeleteObject outrgn
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
Private Sub num_LostFocus()
    Me.num.Text = Fix(Me.num.Text)
    If Me.num.Text <= 0 Then
        Me.num.Text = 1
    End If
End Sub
