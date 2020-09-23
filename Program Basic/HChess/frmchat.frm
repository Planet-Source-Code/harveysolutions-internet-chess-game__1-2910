VERSION 5.00
Begin VB.Form frmchat1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chess Chat"
   ClientHeight    =   3195
   ClientLeft      =   540
   ClientTop       =   495
   ClientWidth     =   5670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   378
   Begin VB.ListBox List1 
      BackColor       =   &H80000002&
      ForeColor       =   &H80000009&
      Height          =   2040
      Left            =   150
      TabIndex        =   1
      Top             =   465
      Width           =   5355
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000002&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   375
      Left            =   150
      TabIndex        =   0
      Top             =   2655
      Width           =   5350
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Chat window"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   75
      Width           =   2895
   End
End
Attribute VB_Name = "frmchat1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
OutSquare 0, 0, ScaleWidth - 1, ScaleHeight, Me
InSquare List1.Left, List1.Top, List1.Width - 1, List1.Height, Me
InSquare Text1.Left, Text1.Top, Text1.Width - 1, Text1.Height, Me
PaintPicture frmChess.Image2.Picture, 15, 5
PaintPicture frmChess.Image2.Picture, ScaleWidth - 35, 5
'Left = frmChess.Left + frmChess.Width
'Top = frmChess.Top + ((frmChess.Height - Height) / 2)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 And Trim(Text1.Text) <> "" Then
  KeyAscii = 0
  List1.AddItem NickName & " Say >> " & Text1.Text
  List1.ListIndex = List1.ListCount - 1
  If Connected Then
    If frmChess.SockClient.State = 7 Then
      frmChess.SockClient.SendData "#010text#" & Text1.Text
    End If
  End If
  Text1.Text = ""
End If
End Sub
