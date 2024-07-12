VERSION 5.00
Begin VB.Form ²Êµ° 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "²Êµ°"
   ClientHeight    =   1245
   ClientLeft      =   13350
   ClientTop       =   5235
   ClientWidth     =   3720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   3720
   StartUpPosition =   1  'ËùÓÐÕßÖÐÐÄ
   Begin VB.CommandButton Command1 
      Caption         =   "ÍË³ö"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton ²Êµ° 
      Caption         =   "ÉÍ×÷ÕßÒ»°ÍÕÆ"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      MousePointer    =   1  'Arrow
      TabIndex        =   2
      Top             =   720
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "²âÊÔ´úÂë.frx":0000
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "²Êµ°"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub
Private Sub Form_Activate()
    Dim rtn As Long
    rtn = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes hwnd, 0, aero, LWA_ALPHA
    Call rgnform(Me, rou, rou, 25, 3, 3, 3)
End Sub
Private Sub ²Êµ°_Click()
    ºÚÆÁ.Show
    Unload Me
End Sub
