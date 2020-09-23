Attribute VB_Name = "Module1"
Public FrmChessWidth
Public FrmChessHeight
Public ChessBoardTop
Public ChessBoardLeft
Public PlayOffline As Boolean
Public Connected As Boolean
Public PartiEnCour As Boolean
Public JoueurHote As Boolean
Public HostName As String
Public TryConnect As Boolean
Public NickName As String
Public VisitorName As String
Public waitforplayer As Boolean
Public fromServer As Boolean
Const TRSPBORDER = 40
Const BORDERTHIN = 5
Sub Main()
    frmSplash.Show
    frmSplash.Refresh
    Unload frmSplash
    frmChess.Top = 10
    frmChess.Left = 10
    frmchat1.Top = 3250
    frmchat1.Left = frmChess.Left + 2000
    frmChess.Show
End Sub
Public Sub OutSquare(col1, lig1, scwidth, scheight, from)
Dim col, lig, bl, rd, gr, coul, wth, ind, ind2
wth = scwidth
col = col1: lig = lig1: ind = 1
bl = TRSPBORDER: gr = TRSPBORDER: rd = TRSPBORDER
For i2 = 0 To 1
 For i = 0 To BORDERTHIN
   from.Line (col, lig)-(wth - col, lig), coul, BF
   lig = lig + ind:   col = col + 1
   bl = bl + TRSPBORDER:   rd = rd + TRSPBORDER:  gr = gr + TRSPBORDER
   coul = RGB(rd, gr, bl)
 Next i
 lig = scheight
 ind = -1: col = col1: bl = TRSPBORDER:   rd = TRSPBORDER:   gr = TRSPBORDER
 coul = RGB(rd, gr, bl)
Next i2
col = col1: ind = 1: lig = lig1
bl = TRSPBORDER: gr = TRSPBORDER: rd = TRSPBORDER: wth = scheight
For i2 = 0 To 1
ind2 = 0
For i = 0 To BORDERTHIN
 from.Line (col, lig)-(col, wth - ind2), coul, BF
lig = lig + 1: col = col + ind
bl = bl + TRSPBORDER: rd = rd + TRSPBORDER: gr = gr + TRSPBORDER
coul = RGB(rd, gr, bl)
ind2 = ind2 + 1
Next i
 lig = lig1: col = scwidth: wth = scheight
 ind = -1: bl = TRSPBORDER: gr = TRSPBORDER: rd = TRSPBORDER
 coul = RGB(rd, gr, bl)
Next i2
from.Refresh
End Sub
Public Sub InSquare(col1, lig1, scwidth, scheight, from)
Dim col, lig, bl, rd, gr, coul, wth, ind, ind2
wth = scwidth + col1
bl = TRSPBORDER: rd = TRSPBORDER: gr = TRSPBORDER
col = col1: ind2 = 1: lig = lig1: ind = 1
For i2 = 0 To 1
ind2 = 0
 For i = 0 To BORDERTHIN
   from.Line (col, lig)-(wth + ind2, lig), coul, BF
   lig = lig - ind:   col = col - 1
   bl = bl + TRSPBORDER: rd = rd + TRSPBORDER:  gr = gr + TRSPBORDER
   coul = RGB(rd, gr, bl)
   ind2 = ind2 + 1
 Next i
 lig = scheight + lig1: ind = -1: col = col1
 bl = TRSPBORDER: rd = TRSPBORDER: gr = TRSPBORDER
 coul = RGB(rd, gr, bl)
Next i2
col = col1: ind = 1: lig = lig1
bl = TRSPBORDER: rd = TRSPBORDER: gr = TRSPBORDER
wth = scheight + lig1
For i2 = 0 To 1
ind2 = 0
For i = 0 To BORDERTHIN
 from.Line (col, lig)-(col, wth + ind2), coul, BF
lig = lig - 1: col = col - ind
bl = bl + TRSPBORDER: rd = rd + TRSPBORDER: gr = gr + TRSPBORDER
coul = RGB(rd, gr, bl)
ind2 = ind2 + 1
Next i
 lig = lig1: col = scwidth + col1
 ind = -1
 bl = TRSPBORDER: rd = TRSPBORDER: gr = TRSPBORDER
 coul = RGB(rd, gr, bl)
Next i2
from.Refresh
End Sub


