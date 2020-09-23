VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Mat Chat (Not Connected)"
   ClientHeight    =   8430
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7185
   Icon            =   "FrmServer.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "FrmServer.frx":08CA
   ScaleHeight     =   8430
   ScaleWidth      =   7185
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   3120
      Top             =   6120
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   495
      Left            =   1440
      TabIndex        =   6
      Top             =   7920
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Height          =   735
      Left            =   5400
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "FrmServer.frx":8F4CC
      Top             =   6120
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   1095
      Left            =   4920
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "FrmServer.frx":8F4D2
      Top             =   6840
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "Username"
      Top             =   6480
      Width           =   1815
   End
   Begin VB.TextBox TextHistory 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000003&
      Height          =   3255
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   600
      Width           =   6735
   End
   Begin VB.TextBox TextMsg 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000003&
      Height          =   735
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "FrmServer.frx":8F4D8
      Top             =   4080
      Width           =   6735
   End
   Begin VB.TextBox TextUserName 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000003&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Text            =   "Username"
      Top             =   5400
      Width           =   1935
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   375
   End
   Begin VB.Image CmdConnect 
      Height          =   345
      Left            =   2400
      Picture         =   "FrmServer.frx":8F4F1
      Top             =   5040
      Width           =   1080
   End
   Begin VB.Image ConnectUp 
      Height          =   345
      Left            =   120
      Picture         =   "FrmServer.frx":9089B
      Top             =   6120
      Width           =   1080
   End
   Begin VB.Image ConnectDown 
      Height          =   345
      Left            =   1320
      Picture         =   "FrmServer.frx":91C45
      Top             =   6120
      Width           =   1080
   End
   Begin VB.Image CmdSetUsername 
      Height          =   345
      Left            =   2400
      Picture         =   "FrmServer.frx":92FEF
      Top             =   5520
      Width           =   1620
   End
   Begin VB.Image UnU 
      Height          =   345
      Left            =   120
      Picture         =   "FrmServer.frx":94D4D
      Top             =   6600
      Width           =   1620
   End
   Begin VB.Image UnD 
      Height          =   345
      Left            =   1800
      Picture         =   "FrmServer.frx":96AAB
      Top             =   6600
      Width           =   1620
   End
   Begin VB.Image CmdDisconnect 
      Height          =   345
      Left            =   4440
      Picture         =   "FrmServer.frx":98809
      Top             =   5040
      Width           =   1080
   End
   Begin VB.Image DcU 
      Height          =   345
      Left            =   120
      Picture         =   "FrmServer.frx":99BB3
      Top             =   6960
      Width           =   1080
   End
   Begin VB.Image DcD 
      Height          =   345
      Left            =   1320
      Picture         =   "FrmServer.frx":9AF5D
      Top             =   6960
      Width           =   1080
   End
   Begin VB.Image CmdExit 
      Height          =   345
      Left            =   4440
      Picture         =   "FrmServer.frx":9C307
      Top             =   5520
      Width           =   1080
   End
   Begin VB.Image ExU 
      Height          =   345
      Left            =   1320
      Picture         =   "FrmServer.frx":9D6B1
      Top             =   7440
      Width           =   1080
   End
   Begin VB.Image ExD 
      Height          =   345
      Left            =   120
      Picture         =   "FrmServer.frx":9EA5B
      Top             =   7440
      Width           =   1080
   End
   Begin VB.Image CmdSend 
      Height          =   810
      Left            =   6000
      Picture         =   "FrmServer.frx":9FE05
      Top             =   5040
      Width           =   1080
   End
   Begin VB.Image SendU 
      Height          =   810
      Left            =   3720
      Picture         =   "FrmServer.frx":A2BD7
      Top             =   6960
      Width           =   1080
   End
   Begin VB.Image SendD 
      Height          =   810
      Left            =   2520
      Picture         =   "FrmServer.frx":A59A9
      Top             =   6960
      Width           =   1080
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   6840
      Top             =   0
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   6600
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************
'------------------------------------------------------
'
'Created And Coded By  * Matthew Lampitt *
'Any Questions E-mail Me, (Matthewlampitt@hotmail.com)
'
'------------------------------------------------------
'******************************************************
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long


Public Sub OpenCD()
    Dim res As Long, returnstring As String * 127
    res = mciSendString("set CDAudio door open", returnstring, 127, 0)
End Sub

Public Sub CloseCD()
    Dim res As Long, returnstring As String * 127
    res = mciSendString("set CDAudio door closed", returnstring, 127, 0)
End Sub
Private Sub CmdConnect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo ac
    'this is just for the effect side of things,
    CmdConnect.Picture = ConnectDown.Picture
    'Now we start with the Simple Winsock Coding
    Winsock1(0).LocalPort = 12345
    'give it a high numbered port because you can almost
    'be posative that it is not in use
    Winsock1(0).Listen
    'how simple is that, winsock is now listening for a
    'connection request
    Me.Caption = "Mat Chat (Waiting For Connection)"
    'this lets the users no what state its in
ac: If Err.Number = 40020 Then MsgBox "Mat Chat Is Already Connected or trying to connect To Server", vbInformation, "Mat Chat"
'just an error handler
End Sub

Private Sub CmdConnect_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'this is just for the effect side of things,
    CmdConnect.Picture = ConnectUp.Picture
End Sub
Private Sub CmdDisconnect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'this is just for the effect side of things,
    CmdDisconnect.Picture = DcD.Picture
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
On Error GoTo Dcon
    Dim Txtmsg As String
    'this is just for the effect side of things,
    CmdSend.Picture = SendD.Picture
    Text2.Text = ""
    Text2.Text = Text2.Text + TextUserName.Text + " Says:  " + TextMsg.Text
    Txtmsg = Text2.Text
    'adds the user name and the word "Says: " to the
    'message
    Winsock1(0).SendData Txtmsg
    'this sends the data in text2 as a string
    TextHistory.Text = TextHistory.Text & vbCrLf & Text2.Text
    TextHistory.SelStart = Len(TextHistory.Text)
    'this put's what you have typed into the history window
    TextMsg.Text = ""
Dcon: If Err.Number = 40006 Then If Err.Number = 40006 Then MsgBox "Mat Chat is Not Connected At The Moment", vbInformation, "Mat Chat"
End Sub

Private Sub Image1_Click()
    'this makes the "x" in the top exit the app
    End
End Sub

Private Sub Image2_Click()
    'this makes the maximize button work
    Me.WindowState = 1
End Sub

Private Sub Image3_Click()
    TextMsg.Text = Winsock1(0).LocalIP
End Sub

Private Sub TextMsg_Click()
    'just another extra
    TextMsg.Text = ""
End Sub


Private Sub Winsock1_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    'this is used when the client is is a way asking,
    'to be connected to the server (This One)
    Winsock1(0).Close
    'close the winsock before you connect, Dont Ask Me why though, haha
    DoEvents 'just get it to do the below events
    Winsock1(0).Accept requestID
    'this just acceps and makes the connection, also ir requests to id
    Me.Caption = "Mat Chat (Connected)"
    'just to let the user know the state of play
    MsgBox "Connected"
End Sub

Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim Txtmsg As String
    Winsock1(0).GetData Txtmsg
    TextHistory.Text = TextHistory.Text & vbCrLf & Txtmsg
    TextHistory.SelStart = Len(TextHistory.Text)
    Text3.Text = Txtmsg
    'the above just recieves the data and puts it on a new line
    'of the chat window
End Sub
