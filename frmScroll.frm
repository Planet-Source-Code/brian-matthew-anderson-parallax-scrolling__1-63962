VERSION 5.00
Begin VB.Form frmParallax 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parallax Scrolling Demo"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9045
   ForeColor       =   &H00FFFFFF&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   300
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   603
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picInstructions 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1395
      Left            =   120
      Picture         =   "frmScroll.frx":0000
      ScaleHeight     =   93
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   180
      TabIndex        =   5
      Top             =   240
      Width           =   2700
   End
   Begin VB.PictureBox picMoon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5055
      Left            =   480
      Picture         =   "frmScroll.frx":21EE
      ScaleHeight     =   337
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   653
      TabIndex        =   4
      Top             =   4200
      Visible         =   0   'False
      Width           =   9795
   End
   Begin VB.PictureBox picWalk 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2910
      Left            =   600
      Picture         =   "frmScroll.frx":5C25
      ScaleHeight     =   194
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   288
      TabIndex        =   3
      Top             =   3720
      Visible         =   0   'False
      Width           =   4320
   End
   Begin VB.PictureBox picGrass 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   3000
      Left            =   5760
      Picture         =   "frmScroll.frx":2EB27
      ScaleHeight     =   200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3600
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.PictureBox picBackBuffer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4575
      Left            =   0
      ScaleHeight     =   305
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   633
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   9495
   End
   Begin VB.PictureBox picWorkbox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   -240
      ScaleHeight     =   209
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   689
      TabIndex        =   0
      Top             =   960
      Visible         =   0   'False
      Width           =   10335
   End
End
Attribute VB_Name = "frmParallax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Parallax scrolling is a method of animating segments of background
'images to give the illusion of real motion. In real life when you're driving a car,
'the fence on the side of the road is wizzing by very quickly while the houses
'in the distance appear to be moving slower. This same concept can be done
'in a 2D environment by scrolling layers of background images. The images
'that appear closest will scroll quicker than images further in the background.
'This concept was inspired by a game I enjoyed on the old Commodore Amiga system
'called "Shadow of the Beast". The parallax scrolling is applied to the grass in
'this example, but clouds or other backgrounds could easily be used.

'I plagiarized some bitmaps - namely the walking character. I have no idea what
'game its from - probably some RPG.

'This program allows the user to change the direction and speed of the scrolling
'by using the arrow keys. Down arrow stops moving. Press 'Q' to exit.

'Created by Brian Anderson 1-07-06.

Option Explicit

Private Declare Function GetTickCount& Lib "kernel32" ()

Private Declare Function BitBlt Lib "gdi32.dll" _
(ByVal hdcDest As Long, ByVal xDest As Long, ByVal yDest As Long, _
ByVal nWidth As Long, ByVal nHeight As Long, _
ByVal hdcSrc As Long, ByVal xSource As Long, ByVal ySource As Long, _
ByVal RasterOp As Long) As Long

Const xSpeedMax As Long = 50

Dim bRunning As Boolean
Dim xSpeed As Long
Dim GrassLayerSpeed(4) As Single
Dim xLoc(4) As Single
Dim FrameCount As Single



Private Sub Form_Load()
Dim CurTime As Long, OldTime As Long
Dim Temp1 As Long
Dim bBlitsDone As Boolean

Randomize (GetTickCount)
bRunning = True
bBlitsDone = False
FrameCount = 0

Me.Width = Screen.TwipsPerPixelX * 560
Me.ScaleHeight = Screen.TwipsPerPixelY * 200
Me.Show

'The workbox holds 2 "grass" images and 2 masks for the grass. It allows me to
'blit any section of the grass in one blit instead of doing part of one
'section and sticking another section on to the end of it.
For Temp1 = 0 To 1
  BitBlt picWorkbox.hDC, Temp1 * 100, 0, 100, _
    picGrass.Height, picGrass.hDC, 0, 0, vbSrcCopy
  BitBlt picWorkbox.hDC, (Temp1 + 2) * 100, 0, 100, _
    picGrass.Height, picGrass.hDC, 100, 0, vbSrcCopy
Next Temp1

'Main loop
While bRunning = True
  DoEvents
  CurTime = GetTickCount
  If bBlitsDone = False Then

'Sets the relative speeds of each section of the grass. There are 5 sections, but
'you could easily create 10 or 20. The more layers, the better it looks. Since this
'is a demo I want people to actually see the color differences.
    GrassLayerSpeed(0) = xSpeed * 0.151
    GrassLayerSpeed(1) = xSpeed * 0.222
    GrassLayerSpeed(2) = xSpeed * 0.378
    GrassLayerSpeed(3) = xSpeed * 0.422
    GrassLayerSpeed(4) = xSpeed * 0.578

'Sets the X origin. Hopefully it makes sense now why I have 2 grass images and
'2 masks.
    For Temp1 = 0 To 4
      xLoc(Temp1) = xLoc(Temp1) + GrassLayerSpeed(Temp1)
      If xLoc(Temp1) > 100 Then
        xLoc(Temp1) = xLoc(Temp1) - 100
      End If
      If xLoc(Temp1) < 0 Then
        xLoc(Temp1) = xLoc(Temp1) + 100
      End If
    Next Temp1

'Used for animating the walking character. FrameCount is a Single variable as opposed
'to a Long because at slow speeds we don't necessarily want to advance to another
'frame of the character's animation.
    FrameCount = (FrameCount + Abs(GrassLayerSpeed(1)))

'Clears the background and replaces it with the moon image.
    BitBlt picBackBuffer.hDC, 0, 0, Me.Width, Me.Height, picMoon.hDC, 0, 0, vbSrcCopy

'The work is done here. Each section of the grass is blitted to the backbuffer
'starting with the top layer (which is the furthest back) to the bottom layer
'(which is the closest to the foreground).
    For Temp1 = 0 To 5
      BitBlt picBackBuffer.hDC, Temp1 * 100, 230, 100, 39, picWorkbox.hDC, _
        xLoc(0), 0, vbSrcPaint
      BitBlt picBackBuffer.hDC, Temp1 * 100, 230, 100, 39, picWorkbox.hDC, _
        xLoc(0) + 200, 0, vbSrcAnd
      BitBlt picBackBuffer.hDC, Temp1 * 100, 240, 100, 39, picWorkbox.hDC, _
        xLoc(1), 40, vbSrcPaint
      BitBlt picBackBuffer.hDC, Temp1 * 100, 240, 100, 39, picWorkbox.hDC, _
        xLoc(1) + 200, 40, vbSrcAnd
      BitBlt picBackBuffer.hDC, Temp1 * 100, 250, 100, 39, picWorkbox.hDC, _
        xLoc(2), 80, vbSrcPaint
      BitBlt picBackBuffer.hDC, Temp1 * 100, 250, 100, 39, picWorkbox.hDC, _
        xLoc(2) + 200, 80, vbSrcAnd
      BitBlt picBackBuffer.hDC, Temp1 * 100, 260, 100, 39, picWorkbox.hDC, _
        xLoc(3), 120, vbSrcPaint
      BitBlt picBackBuffer.hDC, Temp1 * 100, 260, 100, 39, picWorkbox.hDC, _
        xLoc(3) + 200, 120, vbSrcAnd
      BitBlt picBackBuffer.hDC, Temp1 * 100, 270, 100, 39, picWorkbox.hDC, _
        xLoc(4), 160, vbSrcPaint
      BitBlt picBackBuffer.hDC, Temp1 * 100, 270, 100, 39, picWorkbox.hDC, _
        xLoc(4) + 200, 160, vbSrcAnd
    Next Temp1

'And now the "Walking Man".
    If xSpeed <> 0 Then
'Walking right...
      If xSpeed > 0 Then
        BitBlt picBackBuffer.hDC, (Me.ScaleWidth / 2) - 24, 195, 48, 64, _
          picWalk.hDC, (FrameCount / 10 Mod 3) * 48, 0, vbSrcPaint
        BitBlt picBackBuffer.hDC, (Me.ScaleWidth / 2) - 24, 195, 48, 64, _
          picWalk.hDC, ((FrameCount / 10 Mod 3) + 3) * 48, 0, vbSrcAnd
      End If
'Walking left...
      If xSpeed < 0 Then
        BitBlt picBackBuffer.hDC, (Me.ScaleWidth / 2) - 24, 195, 48, 64, _
          picWalk.hDC, (FrameCount / 10 Mod 3) * 48, 128, vbSrcPaint
        BitBlt picBackBuffer.hDC, (Me.ScaleWidth / 2) - 24, 195, 48, 64, _
          picWalk.hDC, ((FrameCount / 10 Mod 3) + 3) * 48, 128, vbSrcAnd
      End If
'Standing still...
    Else
      BitBlt picBackBuffer.hDC, (Me.ScaleWidth / 2) - 24, 195, 48, 64, _
        picWalk.hDC, (FrameCount / 10 Mod 3) * 48, 64, vbSrcPaint
      BitBlt picBackBuffer.hDC, (Me.ScaleWidth / 2) - 24, 195, 48, 64, _
        picWalk.hDC, ((FrameCount / 10 Mod 3) + 3) * 48, 64, vbSrcAnd
    End If
'The backbuffer is done and awaiting a blit to the main form.
  bBlitsDone = True
  End If

'At each tick of the "clock", we do one simple blit to the main screen
If OldTime <> CurTime Then
  BitBlt Me.hDC, 0, 0, Temp1 * 100, 300, picBackBuffer.hDC, 0, 0, vbSrcCopy
  bBlitsDone = False
  OldTime = CurTime
End If

Wend

End

'I hope you enjoyed my example! If you like it, vote. I haven't seen a parallax
'scrolling demo on PSC, so I hope this is helpful for anyone making a RPG or
'side-scrolling game!
End Sub









Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyEscape
  bRunning = False
Case vbKeyLeft
  If Abs(xSpeed - 1) < xSpeedMax Then
    xSpeed = xSpeed - 1
  End If
Case vbKeyRight
  If xSpeed < xSpeedMax - 1 Then
    xSpeed = xSpeed + 1
  End If
Case vbKeyDown
  xSpeed = 0
End Select
  picInstructions.Visible = False
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  End
End Sub
