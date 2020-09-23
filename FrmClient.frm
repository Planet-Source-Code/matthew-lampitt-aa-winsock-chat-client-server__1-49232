VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Mat Chat: (Not Connected)"
   ClientHeight    =   8565
   ClientLeft      =   2550
   ClientTop       =   1125
   ClientWidth     =   7185
   ControlBox      =   0   'False
   Icon            =   "FrmClient.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "FrmClient.frx":08CA
   ScaleHeight     =   8565
   ScaleWidth      =   7185
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   3840
      Top             =   8160
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   495
      Left            =   6120
      TabIndex        =   8
      Top             =   6240
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   " Says:  "
      Top             =   6240
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   855
      Left            =   4920
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "FrmClient.frx":8F4CC
      Top             =   6960
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   255
      Left            =   3840
      TabIndex        =   5
      Top             =   6240
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   1560
      Top             =   7800
   End
   Begin VB.TextBox TextUserName 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000003&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Text            =   "Username"
      Top             =   5640
      Width           =   1935
   End
   Begin VB.TextBox TextIP 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000003&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   5160
      Width           =   1935
   End
   Begin VB.TextBox TextMsg 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000003&
      Height          =   735
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "FrmClient.frx":8F4D2
      Top             =   4080
      Width           =   6735
   End
   Begin VB.TextBox TextHistory 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000003&
      Height          =   3255
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   600
      Width           =   6735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3600
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   0
      Text            =   "Username"
      Top             =   6480
      Width           =   1455
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   6600
      Top             =   0
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   6840
      Top             =   0
      Width           =   375
   End
   Begin VB.Image SendD 
      Height          =   810
      Left            =   2520
      Picture         =   "FrmClient.frx":8F4E9
      Top             =   6960
      Width           =   1080
   End
   Begin VB.Image SendU 
      Height          =   810
      Left            =   3720
      Picture         =   "FrmClient.frx":922BB
      Top             =   6960
      Width           =   1080
   End
   Begin VB.Image CmdSend 
      Height          =   810
      Left            =   5760
      Picture         =   "FrmClient.frx":9508D
      Top             =   5040
      Width           =   1080
   End
   Begin VB.Image ExD 
      Height          =   345
      Left            =   120
      Picture         =   "FrmClient.frx":97E5F
      Top             =   7440
      Width           =   1080
   End
   Begin VB.Image ExU 
      Height          =   345
      Left            =   1320
      Picture         =   "FrmClient.frx":99209
      Top             =   7440
      Width           =   1080
   End
   Begin VB.Image CmdExit 
      Height          =   345
      Left            =   4440
      Picture         =   "FrmClient.frx":9A5B3
      Top             =   5520
      Width           =   1080
   End
   Begin VB.Image DcD 
      Height          =   345
      Left            =   1320
      Picture         =   "FrmClient.frx":9B95D
      Top             =   6960
      Width           =   1080
   End
   Begin VB.Image DcU 
      Height          =   345
      Left            =   120
      Picture         =   "FrmClient.frx":9CD07
      Top             =   6960
      Width           =   1080
   End
   Begin VB.Image CmdDisconnect 
      Height          =   345
      Left            =   4440
      Picture         =   "FrmClient.frx":9E0B1
      Top             =   5040
      Width           =   1080
   End
   Begin VB.Image UnD 
      Height          =   345
      Left            =   1800
      Picture         =   "FrmClient.frx":9F45B
      Top             =   6600
      Width           =   1620
   End
   Begin VB.Image UnU 
      Height          =   345
      Left            =   120
      Picture         =   "FrmClient.frx":A11B9
      Top             =   6600
      Width           =   1620
   End
   Begin VB.Image CmdSetUsername 
      Height          =   345
      Left            =   2400
      Picture         =   "FrmClient.frx":A2F17
      Top             =   5520
      Width           =   1620
   End
   Begin VB.Image ConnectDown 
      Height          =   345
      Left            =   1800
      Picture         =   "FrmClient.frx":A4C75
      Top             =   6120
      Width           =   1620
   End
   Begin VB.Image ConnectUp 
      Height          =   345
      Left            =   120
      Picture         =   "FrmClient.frx":A69D3
      Top             =   6120
      Width           =   1620
   End
   Begin VB.Image CmdConnect 
      Height          =   345
      Left            =   2400
      Picture         =   "FrmClient.frx":A8731
      Top             =   5040
      Width           =   1620
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'-----------------------------------------------
' Created And Coded By  * Matthew Lampitt *
' Any Questions E-mail Me, (Matthewlampitt@hotmail.com)
'-----------------------------------------------
Private Sub CmdConnect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'this is just for the effect side of things,
    CmdConnect.Picture = ConnectDown.Picture
End Sub

Private Sub CmdConnect_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'this is just for the effect side of things,
    CmdConnect.Picture = ConnectUp.Picture
End Sub
Private Sub CmdDisconnect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'this is just for the effect side of things,
    CmdDisconnect.Picture = DcD.Picture
    Winsock1(0).Close
    'this disbles the winsock control
End Sub

Private Sub CmdDisconnect_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'this is just for the effect side of things,
    CmdDisconnect.Picture = DcU.Picture
End Sub
Private Sub CmdExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'this is just for the effect side of things,
    CmdExit.Picture = ExD.Picture
End Sub

Private Sub CmdExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'this is just for the effect side of things,
    CmdExit.Picture = ExU.Picture
    End
End Sub
Private Sub CmdSend_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Command1.Value = True
End Sub

Private Sub CmdSend_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'this is just for the effect side of things,
    CmdSend.Picture = SendU.Picture
End Sub

Private Sub CmdSetUsername_Click()
    'this sets the username to a hidden locked
    'textbox so you can change when you want.
    'not really needed but it makes it more
    'user friendly
    Text1.Text = TextUserName.Text
End Sub

Private Sub CmdSetUsername_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'again just graphic extras
    CmdSetUsername.Picture = UnD.Picture
End Sub

Private Sub CmdSetUsername_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' And Again
    CmdSetUsername.Picture = UnU.Picture
End Sub

Private Sub Command1_Click()
    On Error GoTo NC
    Dim Txtmsg As String
    CmdSend.Picture = SendD.Picture
    'this is just for the effect side of things,
    Text2.Text = ""
    Text2.Text = Text2.Text + Text1.Text + Text3.Text + TextMsg.Text
    'this adds the username and text together in one text box
    'so the data can be sent easly
    Txtmsg = Text2.Text
    'this dims textbox2's text as a string
    Winsock1(0).SendData Txtmsg
    'the above just sends the data in textbox2 as a string
    TextHistory.Text = TextHistory.Text & vbCrLf & Text2.Text
    TextHistory.SelStart = Len(TextHistory.Text)
    'this put's what you have typed into the history window
    TextMsg.Text = ""
NC:     If Err.Number = 40006 Then MsgBox "Mat Chat is Not Connected At The Moment", vbInformation, "Mat Chat"
'just an error handler
End Sub

Private Sub Command2_Click()
Timer1.Enabled = True
    Winsock1(0).RemotePort = 12345
    'connects to the port on the remote computer
    Winsock1(0).RemoteHost = TextIP.Text
    'gives user ability to change the computer that this
    'client connects to.
    Winsock1(0).Connect
    'sends a connection request to server
End Sub

Private Sub Image1_Click()
    'this makes the "x" in the top exit the app work
    End
End Sub

Private Sub Image2_Click()
    'this makes the maximize button work
    Me.WindowState = 1
End Sub

Private Sub TextMsg_Click()
    'just another extra
    TextMsg.Text = ""
End Sub
Private Sub CmdConnect_Click()
    On Error GoTo ac
    Timer1.Enabled = True
    Winsock1(0).RemotePort = 12345
    'connects to the port on the remote computer
    Winsock1(0).RemoteHost = TextIP.Text
    'gives user ability to change the computer that this
    'client connects to.
    Winsock1(0).Connect
    'sends a connection request to server
ac: If Err.Number = 40020 Then MsgBox "Mat Chat Is Already Connected or trying to connect To Server", vbInformation, "Mat Chat"
'just an error handler
End Sub
Private Sub Timer1_Timer()
    Winsock1(0).Close
   Command2.Value = True
End Sub


Private Sub Winsock1_Connect(Index As Integer)
    Timer1.Enabled = False
    MsgBox "Connected"
    Me.Caption = "Connected"
    'just lets the user know that it is connected to server
End Sub

Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim Txtmsg As String
    Winsock1(0).GetData Txtmsg
    TextHistory.Text = TextHistory.Text & vbCrLf & Txtmsg
    TextHistory.SelStart = Len(TextHistory.Text)
End Sub
