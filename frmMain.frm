VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "vIsoX"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6930
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   309
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   462
   Begin VB.PictureBox picMini 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   60
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   60
      Width           =   480
   End
   Begin VB.PictureBox picMini2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   60
      Picture         =   "frmMain.frx":0C42
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   34
      Top             =   60
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picScroll 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   930
      Left            =   600
      ScaleHeight     =   62
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   76
      TabIndex        =   33
      Top             =   3540
      Width           =   1140
   End
   Begin VB.PictureBox picMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   765
      Index           =   15
      Left            =   5940
      Picture         =   "frmMain.frx":1884
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   32
      Top             =   2580
      Width           =   780
   End
   Begin VB.PictureBox picMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   765
      Index           =   14
      Left            =   5100
      Picture         =   "frmMain.frx":3338
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   31
      Top             =   2580
      Width           =   780
   End
   Begin VB.PictureBox picMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   765
      Index           =   13
      Left            =   4260
      Picture         =   "frmMain.frx":4DEC
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   30
      Top             =   2580
      Width           =   780
   End
   Begin VB.PictureBox picMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   765
      Index           =   12
      Left            =   3420
      Picture         =   "frmMain.frx":68A0
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   29
      Top             =   2580
      Width           =   780
   End
   Begin VB.PictureBox picMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   765
      Index           =   11
      Left            =   2580
      Picture         =   "frmMain.frx":8354
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   28
      Top             =   2580
      Width           =   780
   End
   Begin VB.PictureBox picMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   765
      Index           =   10
      Left            =   1740
      Picture         =   "frmMain.frx":9E08
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   27
      Top             =   2580
      Width           =   780
   End
   Begin VB.PictureBox picMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   765
      Index           =   9
      Left            =   900
      Picture         =   "frmMain.frx":B8BC
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   26
      Top             =   2580
      Width           =   780
   End
   Begin VB.PictureBox picMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   765
      Index           =   8
      Left            =   60
      Picture         =   "frmMain.frx":D36E
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   25
      Top             =   2580
      Width           =   780
   End
   Begin VB.PictureBox picMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   765
      Index           =   7
      Left            =   5940
      Picture         =   "frmMain.frx":EE22
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   24
      Top             =   1740
      Width           =   780
   End
   Begin VB.PictureBox picMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   765
      Index           =   6
      Left            =   5100
      Picture         =   "frmMain.frx":108D6
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   23
      Top             =   1740
      Width           =   780
   End
   Begin VB.PictureBox picMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   765
      Index           =   5
      Left            =   4260
      Picture         =   "frmMain.frx":1238A
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   22
      Top             =   1740
      Width           =   780
   End
   Begin VB.PictureBox picMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   765
      Index           =   4
      Left            =   3420
      Picture         =   "frmMain.frx":13E3E
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   21
      Top             =   1740
      Width           =   780
   End
   Begin VB.PictureBox picMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   765
      Index           =   3
      Left            =   2580
      Picture         =   "frmMain.frx":158F2
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   20
      Top             =   1740
      Width           =   780
   End
   Begin VB.PictureBox picMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   765
      Index           =   2
      Left            =   1740
      Picture         =   "frmMain.frx":173A6
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   19
      Top             =   1740
      Width           =   780
   End
   Begin VB.PictureBox picMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   765
      Index           =   1
      Left            =   900
      Picture         =   "frmMain.frx":18E5A
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   18
      Top             =   1740
      Width           =   780
   End
   Begin VB.PictureBox picMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   765
      Index           =   0
      Left            =   60
      Picture         =   "frmMain.frx":1A90E
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   17
      Top             =   1740
      Width           =   780
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   705
      Index           =   0
      Left            =   60
      Picture         =   "frmMain.frx":1C3C2
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   16
      Top             =   60
      Width           =   720
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   765
      Index           =   1
      Left            =   900
      Picture         =   "frmMain.frx":1DE76
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   15
      Top             =   60
      Width           =   780
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   765
      Index           =   2
      Left            =   1740
      Picture         =   "frmMain.frx":1F92A
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   14
      Top             =   60
      Width           =   780
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   765
      Index           =   3
      Left            =   2580
      Picture         =   "frmMain.frx":213DE
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   13
      Top             =   60
      Width           =   780
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   765
      Index           =   4
      Left            =   3420
      Picture         =   "frmMain.frx":22E92
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   12
      Top             =   60
      Width           =   780
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   765
      Index           =   5
      Left            =   4260
      Picture         =   "frmMain.frx":24946
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   11
      Top             =   60
      Width           =   780
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   765
      Index           =   6
      Left            =   5100
      Picture         =   "frmMain.frx":263FA
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   10
      Top             =   60
      Width           =   780
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   765
      Index           =   7
      Left            =   5940
      Picture         =   "frmMain.frx":27EAE
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   9
      Top             =   60
      Width           =   780
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   765
      Index           =   8
      Left            =   60
      Picture         =   "frmMain.frx":29962
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   8
      Top             =   900
      Width           =   780
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   765
      Index           =   9
      Left            =   900
      Picture         =   "frmMain.frx":2B416
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   7
      Top             =   900
      Width           =   780
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   765
      Index           =   10
      Left            =   1740
      Picture         =   "frmMain.frx":2CEC8
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   6
      Top             =   900
      Width           =   780
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   765
      Index           =   11
      Left            =   2580
      Picture         =   "frmMain.frx":2E97C
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   5
      Top             =   900
      Width           =   780
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   765
      Index           =   12
      Left            =   3420
      Picture         =   "frmMain.frx":30430
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   4
      Top             =   900
      Width           =   780
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   765
      Index           =   13
      Left            =   4260
      Picture         =   "frmMain.frx":31EE4
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   3
      Top             =   900
      Width           =   780
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   765
      Index           =   14
      Left            =   5100
      Picture         =   "frmMain.frx":33998
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   2
      Top             =   900
      Width           =   780
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   765
      Index           =   15
      Left            =   5940
      Picture         =   "frmMain.frx":3544C
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   1
      Top             =   900
      Width           =   780
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type MapPoint
x As Long
y As Long
End Type
Dim Tile(31, 31) As Integer
Dim Map(31, 31) As Integer
Dim MapX As Long
Dim MapY As Long
Dim P As MapPoint

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyRight
If picScroll.Left < (Me.ScaleWidth - (picScroll.Width - 10)) Then Exit Sub
picScroll.Left = picScroll.Left - 10
Case vbKeyLeft
If picScroll.Left > -10 Then Exit Sub
picScroll.Left = picScroll.Left + 10
Case vbKeyUp
If picScroll.Top > -5 Then Exit Sub
picScroll.Top = picScroll.Top + 10
Case vbKeyDown
If picScroll.Top < (Me.ScaleHeight - (picScroll.Height - 30)) Then Exit Sub

picScroll.Top = picScroll.Top - 10
End Select
End Sub

Private Sub Form_Load()
Me.Move 0, 0, 12000, 8560
ReadMap
DrawTiles 48, 48, 12, 31 '31 is the max size for autoredraw with tiles that are 128 pixels wide
End Sub

Public Sub DrawTiles(lngTileWidth As Long, lngTileHeight As Long, OffSetY As Long, lngBoardSize As Long)
Dim MapX As Long
Dim MapY As Long
Dim P As MapPoint
With picScroll
.Width = (lngTileWidth * lngBoardSize)
.Height = ((lngTileHeight * lngBoardSize) / 2) + lngTileHeight
.Top = Me.ScaleHeight / 2 - .Height / 2
.Left = Me.ScaleWidth / 2 - .Width / 2
End With
Dim Sec As Long
For MapY = 0 To lngBoardSize - 1
For MapX = 0 To lngBoardSize - 1
P.x = (MapX - MapY) * (lngTileWidth / 2) + ((lngBoardSize * lngTileWidth) / 2) - (lngTileWidth / 2)
P.y = (MapX + MapY) * ((lngTileHeight / 2) - OffSetY)
Sec = StretchBlt(picScroll.hdc, P.x, P.y, lngTileWidth, lngTileHeight, picTile(Tile(MapX, MapY)).hdc, 0, 0, 48, 47, vbSrcAnd)
Sec = StretchBlt(picScroll.hdc, P.x, P.y, lngTileWidth, lngTileHeight, picMask(Tile(MapX, MapY)).hdc, 0, 0, 48, 47, vbSrcPaint)
Next
Next

End Sub

Public Sub ReadMap()
Dim PointString As String
Dim x As Integer
Dim y As Integer
Dim Color1 As Long
Dim Color2 As Long
Dim Color3 As Long
Dim Color4 As Long
Dim Z As Integer
    For y = 0 To 30
        For x = 0 To 30
            Color1 = GetPixel(picMini.hdc, x, y)
            Color2 = GetPixel(picMini.hdc, x + 1, y)
            Color3 = GetPixel(picMini.hdc, x + 1, y + 1)
            Color4 = GetPixel(picMini.hdc, x, y + 1)
            If Color1 = vbBlue Then Map(x, y) = 1 Else Map(x, y) = 0
            If Color2 = vbBlue Then Map(x + 1, y) = 1 Else Map(x + 1, y) = 0
            If Color3 = vbBlue Then Map(x + 1, y + 1) = 1 Else Map(x + 1, y + 1) = 0
            If Color4 = vbBlue Then Map(x, y + 1) = 1 Else Map(x, y + 1) = 0
            PointString = Map(x, y) & Map(x + 1, y) & Map(x + 1, y + 1) & Map(x, y + 1)
            Select Case PointString ' Get The Correct Tile And Draw It
            Case "0000" 'tile= 0
               Tile(x, y) = 0
            Case "0100" 'tile=1 iso=2
                Tile(x, y) = 2
            Case "0010" 'tile=2 iso=4
                Tile(x, y) = 4
            Case "0110" 'tile=3 iso=6
                Tile(x, y) = 6
            Case "0001" 'tile=4 iso=8
                Tile(x, y) = 8
            Case "0101" 'tile=5 iso=10
                Tile(x, y) = 10
            Case "0011" 'tile=6 iso=12
                Tile(x, y) = 12
            Case "0111" 'tile=7 iso=14
                Tile(x, y) = 14
            Case "1000" 'tile=8 iso=1
                Tile(x, y) = 1
            Case "1100" 'tile=9 iso=3
                Tile(x, y) = 3
            Case "1010" 'tile=10 iso=5
                Tile(x, y) = 5
            Case "1110" 'tile=11 iso=7
                Tile(x, y) = 7
            Case "1001" 'tile=12 iso=9
                Tile(x, y) = 9
            Case "1101" 'tile=13 iso=11
                Tile(x, y) = 11
            Case "1011" 'tile=14 iso=13
                Tile(x, y) = 13
            Case "1111" 'tile=15
                Tile(x, y) = 15
            End Select
        Next
    Next
End Sub

