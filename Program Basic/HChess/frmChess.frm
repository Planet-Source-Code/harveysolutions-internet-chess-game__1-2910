VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{5336AD54-C994-11D2-B7D6-444553540000}#11.0#0"; "HCHESSBOARDP.OCX"
Begin VB.Form frmChess 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chess Board"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   473
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   465
   Begin VB.Timer Timer4 
      Left            =   7440
      Top             =   6360
   End
   Begin VB.Timer Timer3 
      Left            =   7440
      Top             =   5880
   End
   Begin VB.Timer Timer2 
      Left            =   7440
      Top             =   5400
   End
   Begin VB.Timer Timer1 
      Left            =   7440
      Top             =   4920
   End
   Begin MSWinsockLib.Winsock SockClient 
      Left            =   7440
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin PicClip.PictureClip PictureClip1 
      Left            =   7320
      Top             =   3960
      _ExtentX        =   6985
      _ExtentY        =   556
      _Version        =   393216
      Picture         =   "frmChess.frx":0000
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   315
      Picture         =   "frmChess.frx":414A
      ScaleHeight     =   525
      ScaleWidth      =   6330
      TabIndex        =   9
      Top             =   480
      Width           =   6330
      Begin VB.Image Image1 
         Height          =   255
         Index           =   5
         Left            =   5880
         ToolTipText     =   "Help"
         Top             =   120
         Width           =   375
      End
      Begin VB.Image Image1 
         Height          =   255
         Index           =   4
         Left            =   5160
         ToolTipText     =   "Chat Window"
         Top             =   120
         Width           =   375
      End
      Begin VB.Image Image1 
         Height          =   255
         Index           =   3
         Left            =   4680
         ToolTipText     =   "Info Game Window"
         Top             =   120
         Width           =   375
      End
      Begin VB.Image Image1 
         Height          =   255
         Index           =   2
         Left            =   3960
         ToolTipText     =   "Music"
         Top             =   120
         Width           =   375
      End
      Begin VB.Image Image1 
         Height          =   255
         Index           =   1
         Left            =   3480
         ToolTipText     =   "No Sound"
         Top             =   120
         Width           =   375
      End
      Begin VB.Image Image1 
         Height          =   255
         Index           =   0
         Left            =   120
         ToolTipText     =   "New Game"
         Top             =   120
         Width           =   375
      End
   End
   Begin HChessBoardP.HChessBoard HChessBoard1 
      Height          =   6360
      Left            =   300
      TabIndex        =   11
      Top             =   480
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   11218
      DiffBoard_Y     =   11
      DiffBoard_X     =   11
      BoardWidth      =   424
      BoardHeight     =   424
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   7320
      Picture         =   "frmChess.frx":4E6A
      Top             =   3480
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Guest Vs Opponent"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   330
      Left            =   600
      TabIndex        =   10
      Top             =   75
      Width           =   5295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   75
      TabIndex        =   8
      Top             =   6120
      Width           =   135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   75
      TabIndex        =   7
      Top             =   5400
      Width           =   135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   75
      TabIndex        =   6
      Top             =   4680
      Width           =   135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   75
      TabIndex        =   5
      Top             =   3960
      Width           =   135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   75
      TabIndex        =   4
      Top             =   3120
      Width           =   135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   75
      TabIndex        =   3
      Top             =   2400
      Width           =   135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   75
      TabIndex        =   2
      Top             =   1680
      Width           =   135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   75
      TabIndex        =   1
      Top             =   960
      Width           =   135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "A            B          C          D          E          F          G           H"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   6855
      Width           =   5775
   End
End
Attribute VB_Name = "frmChess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim HoldIndex
Dim OPTSelected As Boolean
Dim piecejouer As String
Dim hr, min, sec As Integer
Dim hrs, mins, secs As String
Dim tempdebcoup
Dim waitforplayer As Boolean
Dim nbcoupmoi As Long
Dim nbcoup As Long
Dim montour As Boolean
Dim entraitement As Boolean
Dim sendok As Boolean

Public Sub recommendeG()
On Error Resume Next
ChessBoard1.EraseBoard
If PlayOffline Then
   If Connected Then SockClient.SendData "#020code#findepart"
Else
Dim str1
  If JoueurHote Then
   str1 = "false"
  Else
   str1 = "true"
  End If
If Connected Then SockClient.SendData "#020code#recomgame" & str1
End If
'initpartie
End Sub



Private Sub Form_Initialize()
strArg = Command()
If strArg <> "" Then CheckArg strArg
End Sub

Private Sub Form_Unload(Cancel As Integer)
HChessBoard1.ClearGraphicBuffer
Unload frmchat1
Unload frmInfoG
End
End Sub

Private Sub HChessBoard1_EventStatus()
'frmInfoG.Label1.Caption = HChessBoard1.StatusString
End Sub

Private Sub SockClient_ConnectionRequest(ByVal requestID As Long)
On Error Resume Next
Dim rep
rep = MsgBox("You have a Visitor Request, Do You want to Accept ?", vbYesNoCancel, "Connection !")
If rep = vbCancel Or rep = vbNo Then Exit Sub
If SockClient.State <> sckClosed Then
  SockClient.Close
End If
 SockClient.Accept requestID
 SockClient.SendData "#020code#okconnect" & NickName
 Timer3.Interval = 1000
 DoEvents
End Sub
Private Sub SockClient_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim strdata As String
SockClient.GetData strdata
If Mid(strdata, 1, 9) = "#010text#" Then
  frmchat1.List1.AddItem VisitorName & " Say >> " & Mid(strdata, 10, Len(strdata) - 9)
  frmchat1.List1.ListIndex = frmchat1.List1.ListCount - 1
Else
Dim code1, code2
 code1 = Mid(strdata, 1, 14)
 code2 = Mid(strdata, 1, 18)
 Select Case code1
 Case "#010#code:move": joujeuadv strdata, "move"
 Case "#010#code:rock": joujeuadv strdata, "rock"
 Case "#010#code:quen": joujeuadv strdata, "quen"
 Case "#010#code:ches": joujeuadv strdata, "ches"
End Select
Select Case code2
 Case "#020code#wantnewga": WantaNewGame strdata
 Case "#020code#kfornewga": Unload frmWaitting: nouvelleP strdata
 Case "#020code#okconnect": okconnect strdata
 Case "#020code#receiveok": nouvelleP strdata
 Case "#020code#refunewga": RefuseNewGame
 Case "#020code#quitegame": quitthegame
 
 'Case "#020code#movepitou": joujeuadv strdata
 End Select
End If
End Sub
Private Sub quitthegame()
MsgBox VisitorName & " Has left the game !"
HChessBoard1.EraseBoard
PartiEnCour = False
End Sub
Private Sub RefuseNewGame()
MsgBox "Your Opponent has refused your proposal !"
Unload frmWaitting
End Sub
Private Sub WantaNewGame(strdata)
Dim rep, strtemp
rep = MsgBox(VisitorName & " have a new proposal do you want to accept ?", vbYesNo, "New Game")
If rep = vbYes Then
 SockClient.SendData "#020code#kfornewga"
 strtemp = Mid(strdata, 19, Len(strdata) - 18)
 If strtemp = "No" Then
  JoueurHote = True
  HChessBoard1.Host = True
 Else
  JoueurHote = False
  HChessBoard1.Host = False
 End If
  HChessBoard1.PlayOffline = False
  PlayOffline = False
  Connected = True
  HChessBoard1.InitGame
Else
 SockClient.SendData "#020code#refunewga"
 HChessBoard1.EraseBoard
End If
End Sub
Private Sub nouvelleP(strdata)
On Error Resume Next
VisitorName = Mid(strdata, 19, Len(strdata) - 18)
'Label1.Caption = Label1.Caption & " " & Mid(strdata, 19, Len(strdata) - 18)
Connected = True
PartiEnCour = True
Timer3.Interval = 1000
'List1.Clear
'If HChessBoard1.CanIPlay Then frmInfoG.Label1.Caption = "Your Turn" Else frmInfoG.Label1.Caption = "Your Turn"
InitGame
End Sub
Private Sub okconnect(strdt)
On Error Resume Next
nouvelleP strdt
SockClient.SendData "#020code#receiveok" & NickName
DoEvents
End Sub
Private Sub findepartie(msg1)
On Error Resume Next
'Timer5.Interval = 0
'Timer2.Interval = 0
'DoEvents
'Label2.Caption = "Start a New Game !"
'List2.Clear
'Command2.Caption = "New Game"
'ChessBoard1.EraseBoard
'PartiEnCour = False
'MsgBox msg1
End Sub
Private Sub joujeuadv(strdat, cod)
On Error Resume Next
Dim strtemp, code
code = Mid(strdat, 1, 14)
Dim p1, p2, p3, p4
p1 = Val(InStr(1, strdat, "-", vbBinaryCompare))
p2 = Val(InStr(p1 + 1, strdat, "-", vbBinaryCompare))
p3 = Val(InStr(p2 + 1, strdat, "-", vbBinaryCompare))
p4 = Val(InStr(p3 + 1, strdat, "-", vbBinaryCompare))
Dim val1, val2, val3, val4
val1 = Val(Mid(strdat, 15, p1 - 15)): val2 = Val(Mid(strdat, p1 + 1, p2 - 15)): val3 = Val(Mid(strdat, p2 + 1, p3 - 15)): val4 = Val(Mid(strdat, p3 + 1, Len(strdat) - 2 - (p2))) 'lit la position du move
val1 = 7 - val1: val2 = 7 - val2: val3 = 7 - val3: val4 = 7 - val4
strtemp = Str(val1) & Chr(val2 + 65) & " To " & Str(val3) & Chr(val4 + 65)
'List2.List(List2.ListCount - 1) = "I: " & (nbcoup - nbcoupmoi) & " - " & nomoposant & " -T- : " & hrs & ":" & mins & ":" & secs & " -M- : " & strtemp
nbcoup = nbcoup + 1
nbcoupmoi = nbcoupmoi + 1
hrs = "": mins = "": secs = ""
hr = 0: min = 0: sec = 0
'List2.AddItem "I: " & nbcoupmoi & " - " & NickName & " -T- : " & hrs & ":" & mins & ":" & secs
'List2.ListIndex = List2.ListCount - 1
HChessBoard1.MoveThePlayer2Piece val1, val2, val3, val4, cod
DoEvents
waitforplayer = False
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
If min = 59 Then min = 0: hr = hr + 1: hrs = trans(hr): mins = trans(min)
If sec = 59 Then sec = 0: min = min + 1: mins = trans(min)
sec = sec + 1
secs = trans(sec)
'If Not waitforplayer Then
 ' frmInfoG.List1.List(List2.ListCount - 1) = "I: " & nbcoupmoi & " - " & NickName & " -T- : " & hrs & ":" & mins & ":" & secs
'Else
'  frmInfoG.List1.List(List2.ListCount - 1) = "I: " & (nbcoup - nbcoupmoi) & " - " & nomoposant & " -T- : " & hrs & ":" & mins & ":" & secs
'End If
End Sub
Private Sub Timer3_Timer()
checkdisconnect
End Sub
Private Sub checkdisconnect()
On Error Resume Next
If SockClient.State <> 7 And Connected Then
  SockClient.Close
  MsgBox "You are Disconnected !"
  PartiEnCour = False
  Unload frmWaitting
  Timer2.Interval = 0
  Timer3.Interval = 0
  Connected = False
  HChessBoard1.EraseBoard
  PartiEnCour = False
  'frmInfoG.Label1.Caption = "Start a New Game !"
  'frmInfoG.List1.Clear: List2.Clear
End If
End Sub

Private Function trans(var1) As String
On Error Resume Next
If var1 = 0 Then
 trans = "00"
Else
If var1 < 10 Then
  trans = "0" & Trim(Str(var1))
Else
  trans = Trim(Str(var1))
End If
End If
End Function


Public Sub InitGame()
On Error Resume Next
hrs = "": mins = "": secs = ""
hr = 0: min = 0: sec = 0
nbcoup = 1:
If PlayOffline Then
  HChessBoard1.PlayOffline = True
  HChessBoard1.Host = True
  Command2.Caption = "New Game"
  'frmInfoG.Label1.Caption = "Offline game !"
Else
  HChessBoard1.PlayOffline = False
  If JoueurHote Then
    nbcoupmoi = 1
    waitforplayer = False
    HChessBoard1.Host = True
   ' frmInfoG.Label1.Caption = "Your Turn !"
  Else
    nbcoupmoi = 0
    HChessBoard1.Host = False
    waitforplayer = True
    'frmInfoG.Label1.Caption = "Wait Your Turn !"
  End If
  Timer2.Interval = 1000
 End If
Label3.Caption = Label3.Caption & VisitorName
HChessBoard1.Sound = True
HChessBoard1.MoveString = ""
HChessBoard1.CreateGraphicBuffer
HChessBoard1.InitGame
'frmInfoG.List1.Clear
PartiEnCour = True
End Sub
Private Sub Form_Load()
'On Error Resume Next
Dim strArg As String
OutSquare 0, 0, ScaleWidth - 1, ScaleHeight, Me
InSquare HChessBoard1.Left, HChessBoard1.Top, HChessBoard1.Width - 1, HChessBoard1.Height, Me
PaintPicture Image2.Picture, 15, 5
PaintPicture Image2.Picture, ScaleWidth - 35, 5
JoueurHote = True
HostName = "Computer Name/IP"
NickName = "Guest"
VisitorName = "Opponent"
SockClient.Protocol = sckTCPProtocol
SockClient.RemoteHost = "pc2"
SockClient.RemotePort = 1004
SockClient.LocalPort = 1004
HoldIndex = -1
currentdirectory = CurDir("")
PictureClip1.Cols = 12
For i = 0 To 5
 Image1(i).Picture = PictureClip1.GraphicCell(i)
 Next i
If fromServer Then currentdirectory = currentdirectory & "\VChess"
'frmChat1.Show
'frmChat1.Left = 500
HChessBoard1.PickUpSoundFile = currentdirectory & "\Sounds\PickUp.wav"
HChessBoard1.PutDownSoundFile = currentdirectory & "\Sounds\WoodThunk.wav"
HChessBoard1.StartGameSoundFile = currentdirectory & "\Sounds\Opening.wav"
HChessBoard1.ChessSoundFile = "c:\windows\Media\canyon.mid" 'currentdirectory & "\Sounds\Rockem.mid"
HChessBoard1.MoveNotAllowedSoundFile = currentdirectory & "\Sounds\Orchestra.wav"
Set HChessBoard1.PiecePicture = LoadPicture(currentdirectory & "\Images\PieceBR.bmp")
Set HChessBoard1.BoardPicture = LoadPicture(currentdirectory & "\Images\BlueMarbBoard.bmp")
'HChessBoard1.PickUpSoundFile = "e:\Program Basic\HChessGameProj\Sounds\PickUp.wav"
'HChessBoard1.PutDownSoundFile = "e:\Program Basic\HChessGameProj\Sounds\WoodThunk.wav"
'HChessBoard1.StartGameSoundFile = "e:\Program Basic\HChessGameProj\Sounds\Opening.wav"
'HChessBoard1.ChessSoundFile = "e:\Program Basic\HChessGameProj\Sounds\Orchestra.wav"
'HChessBoard1.MoveNotAllowedSoundFile = "e:\Program Basic\HChessGameProj\Sounds\Orchestra.wav"
'Set HChessBoard1.PiecePicture = LoadPicture("e:\Program Basic\HChessGameProj\Images\PieceBR.bmp")
'Set HChessBoard1.BoardPicture = LoadPicture("e:\Program Basic\HChessGameProj\Images\BlueMarbBoard.bmp")
DoEvents
End Sub
Private Sub CheckArg(strdata)
fromServer = True
Me.Visible = True
Dim pos, pos2, nick, myname, hisname, caseT, hisIP
pos = InStr(10, Trim(strdata), "|", vbBinaryCompare)
myname = Mid(Trim(strdata), 10, pos - 10)
pos2 = InStr(pos + 1, Trim(strdata), "|", vbBinaryCompare)
hisname = Mid(Trim(strdata), pos + 1, pos2 - pos1)
VisitorName = hisname
VisitorName = myname
HChessBoard1.PlayOffline = False
PlayOffline = False
caseT = Mid(strdata, 1, 9)
If caseT = "|HostUsr|" Then
  HChessBoard1.Host = True
  JoueurHote = True
  frmChess.SockClient.Close
  frmChess.SockClient.Protocol = sckTCPProtocol
  frmChess.SockClient.LocalPort = 1004
  frmChess.SockClient.Listen
  'frmInfoG.Label1.Caption = "Wait For Other Player"
  waitforplayer = True
  frmtryConnect.GetinitWaitting
  DoEvents
  frmtryConnect.Show 1
ElseIf caseT = "|HostTsr|" Then
  hisIP = Mid(Trim(strdata), pos2 + 1, Len(strdata))
  HChessBoard1.Host = False
  JoueurHote = False
  HostName = hisIP
  SockClient.RemotePort = 1004
  SockClient.RemoteHost = HostName
  'frmInfoG.Label1.Caption = "Try to Connect"
  NameNIP.Connect
  DoEvents
  frmtryConnect.GetinitWaitting
  DoEvents
  frmtryConnect.Show 1
End If
'MsgBox ArgStr
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture1.Visible = Not Picture1.Visible
End Sub

Private Sub HChessBoard1_PieceMoved()
'On Error Resume Next
 If Not HChessBoard1.CanIPlay Then
 If Connected Then
  'MsgBox "j'ai bouger :" & HChessBoard1.MoveString & "    :   " & HChessBoard1.CanIPlay
  Dim strtemp, code, p1, p2, p3, p4
  p1 = Val(InStr(1, HChessBoard1.MoveString, "-", vbBinaryCompare))
  p2 = Val(InStr(p1 + 1, HChessBoard1.MoveString, "-", vbBinaryCompare))
  p3 = Val(InStr(p2 + 1, HChessBoard1.MoveString, "-", vbBinaryCompare))
  p4 = Val(InStr(p3 + 1, HChessBoard1.MoveString, "-", vbBinaryCompare))
  Dim val1, val2, val3, val4
  val1 = Val(Mid(HChessBoard1.MoveString, 15, p1 - 15)): val2 = Val(Mid(HChessBoard1.MoveString, p1 + 1, p2 - 15)): val3 = Val(Mid(HChessBoard1.MoveString, p2 + 1, p3 - 15)): val4 = Val(Mid(HChessBoard1.MoveString, p3 + 1, Len(HChessBoard1.MoveString) - 2 - (p2))) 'lit la position du move
  strtemp = Str(val1) & Chr(val2 + 65) & " To " & Str(val3) & Chr(val4 + 65)
  SockClient.SendData HChessBoard1.MoveString
  DoEvents
  'frmInfoG.List1.List(frmInfoG.List1.ListCount - 1) = "I: " & nbcoupmoi & " - " & NickName & " -T- : " & hrs & ":" & mins & ":" & secs & " -M- : " & strtemp
  nbcoup = nbcoup + 1
  hrs = "": mins = "": secs = ""
  hr = 0: min = 0: sec = 0
  'frmInfoG.List1.AddItem "I: " & (nbcoup - nbcoupmoi) & " - " & nomoposant & " -T- : " & hrs & ":" & mins & ":" & secs & " -M- : " & strtemp
  'frmInfoG.List1.ListIndex = frmInfoG.List1.ListCount - 1
  'frmInfoG.Label1.Caption = "Wait Your Turn"
  waitforplayer = True
  End If
 Else
  waitforplayer = False
 End If
End Sub

Private Sub OpenFile()
    Dim sFile As String
    With dlgCommonDialog
        'To Do
        'set the flags and attributes of the
        'common dialog control
        .Filter = "All Files (*.*)|*.*"
        .ShowOpen
        If Len(.filename) = 0 Then
            Exit Sub
        End If
        sFile = .filename
    End With
    'To Do
    'process the opened file
End Sub

Private Sub Image1_Click(Index As Integer)
Select Case Index
Case 0: NewGame
Case 1: HChessBoard1.Sound = Not HChessBoard1.Sound 'son
Case 2: HChessBoard1.Music = Not HChessBoard1.Music 'music
Case 3: MsgBox "This Option is not part of the Demo" 'frmInfoG.Visible = Not frmInfoG.Visible 'info
Case 4: frmchat1.Visible = Not frmchat1.Visible 'chat
Case 5: frmAbout.Show 1
End Select
End Sub
Private Sub NewGame()
Dim rep
If PartiEnCour Then
 rep = MsgBox("You Have a Game Currently on going do you still want to continue ?", vbYesNoCancel, "New Game")
 If rep = vbYes Then
     frmnewgame.Show 1
 End If
Else
frmnewgame.Show 1
End If
End Sub
Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index <> HoldIndex Then
OPTSelected = True
 If HoldIndex = -1 Then HoldIndex = 0
 Image1(HoldIndex).Picture = PictureClip1.GraphicCell(HoldIndex)
 Image1(Index).Picture = PictureClip1.GraphicCell(Index + 6)
End If
HoldIndex = Index
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If OPTSelected Then Image1(HoldIndex).Picture = PictureClip1.GraphicCell(HoldIndex): OPTSelected = False: HoldIndex = -1
End Sub

