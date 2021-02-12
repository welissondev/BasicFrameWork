Attribute VB_Name = "WindowsApi"
'@Folder "Frabasic"
Option Explicit

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Const GWL_STYLE1 = (-16)
Const GWL_STYLE2 = (-20)
Const WS_CAPTION = &HC00000
Const SWP_FRAMECHANGED = &H20
Const WS_EX_LAYERED = &H80000
Const LWA_ALPHA = &H2&

Public hwnd As Long

'//Código para versão do windows 64 bits
#If VBA7 Then
    Private Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, _
    ByVal lpWindowName As String) As Long
    
    Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hwnd As Long, _
    lpRect As RECT) As Long
    
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal _
    hwnd As Long, ByVal nIndex As Long) As Long
    
    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal _
    hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    
    Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal CX As Long, _
    ByVal CY As Long, ByVal wFlags As Long) As Long
    
    Private Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, _
    ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
    
    Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    
'//Código para versão do windows 32 bits
#Else
    Private Declare Function FindWindowA Lib "user32" (ByVal lpClassName As String, _
    ByVal lpWindowName As String) As Long
    
    Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, _
    lpRect As RECT) As Long
    
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal _
    hWnd As Long, ByVal nIndex As Long) As Long
    
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal _
    hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    
    Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, _
    ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) As Long
    
    Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, _
    ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
    
    Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Public Sub RemoveFormHead(stCaption As String, sbxVisible As Boolean)
        
    Dim vrWin As RECT
    Dim style As Long
    Dim lhWnd As Long
    
    lhWnd = FindWindowA(vbNullString, stCaption)
    GetWindowRect lhWnd, vrWin
    style = GetWindowLong(lhWnd, GWL_STYLE1)
    
    If sbxVisible Then
        SetWindowLong lhWnd, GWL_STYLE1, style Or WS_CAPTION
    Else
        SetWindowLong lhWnd, GWL_STYLE1, style And Not WS_CAPTION
    End If
    
    SetWindowPos lhWnd, 0, vrWin.Left, vrWin.Top, vrWin.Right - vrWin.Left, _
    vrWin.Bottom - vrWin.Top, SWP_FRAMECHANGED
        
End Sub


Public Sub MakeTransparent(frm As Object, TransparentValue As Integer)

    Dim bytOpacity As Byte
    
    'Control the opacity setting.
    bytOpacity = TransparentValue
    
    hwnd = FindWindowA("ThunderDFrame", frm.Caption)
    Call SetWindowLong(hwnd, GWL_STYLE2, GetWindowLong(hwnd, GWL_STYLE2) Or WS_EX_LAYERED)
    Call SetLayeredWindowAttributes(hwnd, 0, bytOpacity, LWA_ALPHA)

End Sub

Public Sub ExecuteSleep(Time As Integer)
    Sleep Time
End Sub
