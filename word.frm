VERSION 5.00
Begin VB.Form wordbox 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÎÄ±¾¿ò"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4680
   Icon            =   "word.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   5  'Size
   ScaleHeight     =   3615
   ScaleWidth      =   4680
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Height          =   150
      Left            =   0
      MaskColor       =   &H8000000A&
      MousePointer    =   1  'Arrow
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ÍË³ö"
      Height          =   375
      Left            =   3480
      MousePointer    =   1  'Arrow
      TabIndex        =   1
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   2895
      Left            =   120
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "wordbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub
Private Sub Command4_Click()
    ²Êµ°.Show
End Sub
Private Sub Form_Activate()
    Dim rtn As Long
    rtn = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes hwnd, 0, aero, LWA_ALPHA
    If gif = True Then
        If gif_wordbox <> True Then
            Call rgn_gif(Me, 25, 3, 3, 3)
            gif_wordbox = True
        End If
    End If
    Call rgnform(Me, rou, rou, 25, 3, 3, 3)
End Sub
Private Sub Form_Load()
    If sfly = 1 And sfl = "all" Then
        Me.left = GetSetting(App.ProductName, "location", "left_wordbox", Screen.Width / 2 - Me.Width / 2)
        Me.top = GetSetting(App.ProductName, "location", "top_wordbox", Screen.Height / 2 - Me.Height / 2)
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
Private Sub Form_Unload(Cancel As Integer)
    gif_wordbox = False
    SaveSetting App.ProductName, "location", "left_wordbox", Me.left
    SaveSetting App.ProductName, "location", "top_wordbox", Me.top
    DeleteObject outrgn
End Sub
