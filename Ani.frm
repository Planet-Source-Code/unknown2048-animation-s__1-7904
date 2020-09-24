VERSION 5.00
Begin VB.Form Animation 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'Kein
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
End
Attribute VB_Name = "Animation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'2020.exe ist a Example
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Sub CHide()
a = ShowCursor(1)
Do While a >= 0
a = ShowCursor(0)
Loop
End Sub
Private Sub CShow()
a = ShowCursor(0)
Do While a < 0
a = ShowCursor(1)
Loop
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
CShow
End
End Sub

Private Sub Form_Load()
Me.Show
CHide
  Me.WindowState = 2
  ScaleMode = 3
  DrawWidth = 60
If AktAufloesung$ = "800 x 600" Then Else MsgBox "Optimal Pixel 800 x 600", vbInformation, ""
Do
Animation1
Animation2
Animation3
Animation4
Animation5
Loop
End Sub
Private Sub Animation1()
  ScaleMode = 3
  DrawWidth = 60
For I = 1 To 100
PsetX 800, I, 20, 10, 255, 255, 255
Next
End Sub
Private Sub Animation2()
  ScaleMode = 3
  DrawWidth = 60
For I = 1 To 100
PsetX 800, 20, I, 10, 255, 255, 255
Next
End Sub
Private Sub Animation3()
  ScaleMode = 1
  DrawWidth = Int(10 * Rnd) + 1
Randomize Timer
YX = Int(10000 * Rnd)
Do While Zahl < YX
Zahl = Zahl + 1
 red = Int(255 * Rnd)
Green = Int(120 * Rnd)
Blue = Int(75 * Rnd)
X = Int(Me.Width * Rnd)
Y = Int(Me.Height * Rnd)
R = Int(Me.ScaleWidth / 10 * Rnd)
Circle (X, Y), R, RGB(red, Green, Blue), , , 1
DoEvents
Loop
End Sub
Private Sub Animation4()
  ScaleMode = 1
  DrawWidth = Int(10 * Rnd) + 1
Randomize Timer
YX = Int(10000 * Rnd)
Zahl = 0
Do While Zahl < YX
Zahl = Zahl + 1
 red = Int(255 * Rnd)
Green = Int(120 * Rnd)
Blue = Int(75 * Rnd)
X = Int(Me.Width * Rnd)
Y = Int(Me.Height * Rnd)
R = Int(Me.ScaleWidth / 10 * Rnd)
Circle (X, Y), R, RGB(red, Green, Blue), , , 1
DoEvents
Loop
End Sub
Private Sub Animation5()
Me.Show
ScaleMode = 1
DrawWidth = 10
Randomize Timer
YX = Int(10000 * Rnd)
Zahl = 0
Do While Zahl < YX
Zahl = Zahl + 1
LineHeight = Int(2000 * Rnd)
LineWidth = Int(20000 * Rnd)
LineColor = Int(15 * Rnd)
Line (0, 0)-(LineWidth, LineHeight), RGB(Int(255 * Rnd), Int(255 * Rnd), Int(70 * Rnd))
Line (6000, 6000)-(LineWidth, LineHeight), RGB(Int(255 * Rnd), Int(120 * Rnd), Int(70 * Rnd))
Line (12000, 12000)-(LineWidth, LineHeight), RGB(Int(70 * Rnd), Int(95 * Rnd), Int(255 * Rnd))
Line (1200, 6500)-(LineWidth, LineHeight), RGB(Int(255 * Rnd), Int(90 * Rnd), Int(120 * Rnd))
Line (0, 12500)-(LineWidth, LineHeight), RGB(Int(55 * Rnd), Int(255 * Rnd), Int(255 * Rnd))
DoEvents
Loop
End Sub
Private Sub PsetX(WhileXis, YPosPlus, XPosPlus, Color As Integer, red, Green, Blue)
Do While XPos < WhileXis
YPos = YPos + YPosPlus 'Width
XPos = XPos + XPosPlus 'Height
Cred = Int(Rnd * red)
CGreen = Int(Rnd * Green)
CBlue = Int(Rnd * Blue)
PSet (XPos, YPos), RGB(Cred, CGreen, CBlue)
DoEvents
Loop
End Sub
