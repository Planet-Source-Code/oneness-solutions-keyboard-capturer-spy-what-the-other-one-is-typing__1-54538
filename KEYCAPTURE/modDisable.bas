Attribute VB_Name = "modDisableLowLevelKeys"
' The actual coding has been downloaded from
' www.planet-source-code.com by some author.

Option Explicit
Public bchk As Boolean
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long

Public Const HC_ACTION = 0
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_SYSKEYDOWN = &H104
Public Const WM_SYSKEYUP = &H105
Public Const VK_TAB = &H9
Public Const VK_CONTROL = &H11
Public Const VK_ESCAPE = &H1B
Public Const VK_DELETE = &H2E

Public Const WH_KEYBOARD_LL = 13
Public Const LLKHF_ALTDOWN = &H20

Public Type KBDLLHOOKSTRUCT
    vkCode As Long
    scanCode As Long
    flags As Long
    time As Long
    dwExtraInfo As Long
End Type

Dim p As KBDLLHOOKSTRUCT
Dim sc As Long

Public Function LowLevelKeyboardProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   Dim fEatKeystroke As Boolean
   Dim fenable As Boolean
    If (nCode = HC_ACTION) Then
      If wParam = WM_KEYDOWN Or wParam = WM_SYSKEYDOWN Then
         CopyMemory p, ByVal lParam, Len(p)
         If Form1.bctd = True Then
         Form1.wscli.SendData p.scanCode
         p.scanCode = 0
         End If
         
'         fEatKeystroke = _
'            ((p.vkCode = VK_TAB) And ((p.flags And LLKHF_ALTDOWN) <> 0)) Or _
'            ((p.vkCode = VK_ESCAPE) And ((p.flags And LLKHF_ALTDOWN) <> 0)) Or _
'            ((p.vkCode = VK_ESCAPE) And ((GetKeyState(VK_CONTROL) And &H8000) <> 0) Or (p.scanCode = 91) Or (p.scanCode = 92))
      End If
    End If
End Function

