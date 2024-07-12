Attribute VB_Name = "函数声明"
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Const WS_EX_LAYERED = &H80000
Public Const GWL_EXSTYLE = (-20)
Public Const LWA_ALPHA = &H2
Public Const LWA_COLORKEY = &H1
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Sub rgnform(ByVal frmbox As Form, ByVal fw As Long, ByVal fh As Long, ByVal top As String, ByVal left As String, ByVal bottom As String, ByVal right As String)
    Dim w As Long, h As Long
    w = frmbox.ScaleX(frmbox.Width, vbTwips, vbPixels) - right
    h = frmbox.ScaleY(frmbox.Height, vbTwips, vbPixels) - bottom
    outrgn = CreateRoundRectRgn(left, top, w, h, fw, fh)
    Call SetWindowRgn(frmbox.hwnd, outrgn, True)
End Sub
Public Sub Main()
    App.Title = "spjq_refresh"
    SaveSetting App.ProductName, "settings", "path", App.Path
    aero = GetSetting(App.ProductName, "settings", "aero", 210)
    rou = GetSetting(App.ProductName, "settings", "round", 20)
    spl = GetSetting(App.ProductName, "settings", "splash", 1)
    splt = GetSetting(App.ProductName, "settings", "splash_time", 2000)
    sfly = GetSetting(App.ProductName, "settings", "save_form_location_yn", 1)
    sfl = GetSetting(App.ProductName, "settings", "save_form_location", "o_main")
    gif = GetSetting(App.ProductName, "settings", "gif", True)
    If spl = 1 Then
        frmSplash.Show
    Else
        Form1.Show
    End If
End Sub
Public Sub sp(ByVal number As String, ByVal wait As String, ByVal wai_enter As String, ByVal way As String)
    n = number
    w = wait
    w = w * 1000
    ent = wai_enter
    MsgBox "请点击要刷屏的文本框", 4096, "引导"
    If way = 1 Then
        For i = 1 To n
            SendKeys "^v"
            Call Sleep(ent)
            SendKeys "{ENTER}"
            Call Sleep(w)
        Next
    ElseIf way = 2 Then
        For i = 1 To n
            SendKeys "^v"
            Call Sleep(ent)
            SendKeys "^{ENTER}"
            Call Sleep(w)
        Next
    ElseIf way = 3 Then
        For i = 1 To n
            SendKeys "^v"
            Call Sleep(ent)
            SendKeys "%s"
            Call Sleep(w)
        Next
    End If
    Call Sleep(500)
End Sub
Public Sub rgn_gif(ByVal frm As Form, ByVal top As String, ByVal left As String, ByVal bottom As String, ByVal right As String)
    frm.Refresh
    If frm.Width / 2 > frm.Height / 2 Then
        b = frm.Height / 25
    Else
        b = frm.Width / 30
    End If
    a = rou * 3
    For i = 1 To b / 1
        Call rgnform(frm, a, a, b + top, b + left, b + bottom, b + right)
        Sleep (0)
        b = b - 1
    Next
    For i = 1 To (a - rou) / 1
        Call rgnform(frm, a, a, b + top, b + left, b + bottom, b + right)
        Sleep (0)
        a = a - 1
    Next
    Call rgnform(frm, rou, rou, top, left, bottom, right)
End Sub
