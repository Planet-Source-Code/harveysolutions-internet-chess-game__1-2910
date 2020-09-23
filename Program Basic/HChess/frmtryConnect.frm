VERSION 5.00
Begin VB.Form frmtryConnect 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Try To Connect to :"
   ClientHeight    =   810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5475
   LinkTopic       =   "Form1"
   ScaleHeight     =   54
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   365
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   120
      Top             =   720
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   225
      Left            =   240
      Shape           =   3  'Circle
      Top             =   315
      Width           =   225
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      X1              =   16
      X2              =   344
      Y1              =   28
      Y2              =   28
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   150
      TabIndex        =   0
      Top             =   240
      Width           =   5175
   End
End
Attribute VB_Name = "frmtryConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim TryConnect As Boolean
Dim ind
Dim tempconnect As Integer
Private Sub Form_Load()
OutSquare 0, 0, ScaleWidth - 1, ScaleHeight, Me
InSquare Label1.Left, Label1.Top, Label1.Width - 1, Label1.Height, Me
GetinitWaitting
End Sub
Public Sub GetinitWaitting()
If Not TryConnect Then
Caption = "Wait for Opponent..."
Else
Caption = "Try to Connect to :" & HostName
End If
ind = 4
Timer1.Interval = 10
Shape1.Left = 22
tempconnect = 0
Left = frmChess.Left + ((frmChess.Width - Width) / 2)
Top = frmChess.Top + ((frmChess.Height - Height) / 2)
End Sub
Private Sub Form_Terminate()
Timer1.Interval = 0
'frmChess.SockClient.Close
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
If tempconnect = 1000 And frmChess.SockClient.State <> 7 Then
 MsgBox "Can't establish connection"
 Close
 Timer1.Interval = 0: tempconnect = 0: Connected = False
 frmChess.SockClient.Close
 'Label2.Caption = "Start a New Game !"
 Unload Me
 Exit Sub
Else
If tempconnect = 800 And TryConnect Then
If frmChess.SockClient.State <> 7 Then NameNIP.Connect
Else
If tempconnect = 600 And TryConnect Then
If frmChess.SockClient.State <> 7 Then NameNIP.Connect
Else
If tempconnect = 400 And TryConnect Then
If frmChess.SockClient.State <> 7 Then NameNIP.Connect
Else
If tempconnect = 200 And TryConnect Then
If frmChess.SockClient.State <> 7 Then NameNIP.Connect
Else
If frmChess.SockClient.State = 7 Then
Timer1.Interval = 0: tempconnect = 0
Unload Me
Exit Sub
End If
End If
End If
End If
End If
End If
tempconnect = tempconnect + 1
If Shape1.Left >= 338 Then
ind = -4
Else
If Shape1.Left <= 22 Then
ind = 4
End If
End If
Shape1.Left = Shape1.Left + ind
End Sub
