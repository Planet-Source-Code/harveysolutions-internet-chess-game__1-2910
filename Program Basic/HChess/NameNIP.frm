VERSION 5.00
Begin VB.Form NameNIP 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Essential Information"
   ClientHeight    =   2355
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4035
   LinkTopic       =   "Form1"
   ScaleHeight     =   157
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   269
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000002&
      ForeColor       =   &H80000009&
      Height          =   375
      Left            =   150
      TabIndex        =   1
      Top             =   1200
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000002&
      ForeColor       =   &H80000009&
      Height          =   375
      Left            =   150
      TabIndex        =   0
      Top             =   480
      Width           =   3735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nick Name"
      Height          =   255
      Index           =   1
      Left            =   150
      TabIndex        =   4
      Top             =   930
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Computer Name or IP"
      Height          =   255
      Index           =   0
      Left            =   150
      TabIndex        =   3
      Top             =   150
      Width           =   2295
   End
End
Attribute VB_Name = "NameNIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
waitforplayer = False
HostName = Text1.Text
TryConnect = False
NickName = Text2.Text
frmChess.Label3.Caption = NickName & " Vs "
If PlayOffline Then
Unload Me
VisitorName = Text1.Text
NickName = Text2.Text
frmChess.InitGame
Else
 If JoueurHote Then
  If frmChess.SockClient.State <> 7 Then frmChess.SockClient.Close
  frmChess.SockClient.Protocol = sckTCPProtocol
  frmChess.SockClient.LocalPort = 1004
  frmChess.SockClient.Listen
  'frmInfoG.Label1.Caption = "Wait For Other Player"
  waitforplayer = True
  Unload Me
  frmtryConnect.GetinitWaitting
  frmtryConnect.Show 1
  Exit Sub
Else
  'frmInfoG.Label1.Caption = "Try to Connect"
  Unload Me
  Connect
  frmtryConnect.GetinitWaitting
  frmtryConnect.Show 1
  '
End If
End If
End Sub
Public Sub Connect()
'On Error Resume Next
  TryConnect = True
  frmChess.SockClient.Close
  frmChess.SockClient.Protocol = sckTCPProtocol
  frmChess.SockClient.RemotePort = 1004
  frmChess.SockClient.RemoteHost = HostName
  frmChess.SockClient.Connect
End Sub
Private Sub Form_Load()

OutSquare 0, 0, ScaleWidth - 1, ScaleHeight, Me
InSquare Text1.Left, Text1.Top, Text1.Width - 1, Text1.Height, Me
InSquare Text2.Left, Text2.Top, Text2.Width - 1, Text2.Height, Me
InSquare Command1.Left, Command1.Top, Command1.Width - 1, Command1.Height, Me
If PlayOffline Then
 Label1(0).Caption = "Nickname Visitor"
 HostName = "Visitor"
Else
 Label1(0).Caption = "Computer Name or IP"
End If
Text1.Text = HostName
Text2.Text = NickName
Left = frmChess.Left + ((frmChess.Width - Width) / 2)
Top = frmChess.Top + ((frmChess.Height - Height) / 2)
End Sub

