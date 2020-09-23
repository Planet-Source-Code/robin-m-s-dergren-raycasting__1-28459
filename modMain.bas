Attribute VB_Name = "modMain"
Option Explicit

Private Sub MoveMan()
    'If the man is to turn to the left
    If Man.TurnLeft Then
        'Decrease its angle
        Man.Angle = Man.Angle - 6
        'Update the map
        bMAPChanged = True
    'Turn to the right
    ElseIf Man.TurnRight Then
        'Increase the angle
        Man.Angle = Man.Angle + 6
        'Update the map
        bMAPChanged = True
    End If
    'If the man is to walk forward
    If Man.WalkFwrd Then
        'Check for outside-walls
        If ((Man.X + CT(Man.Angle) * 10) / Options.GRID_SIZE) < 0 Or ((Man.X + CT(Man.Angle) * 10) / Options.GRID_SIZE) > Lvl.WORLD_SIZE Or _
           ((Man.Y + ST(Man.Angle) * 10) / Options.GRID_SIZE) < 0 Or ((Man.Y + ST(Man.Angle) * 10) / Options.GRID_SIZE) > Lvl.WORLD_SIZE Then Exit Sub
        'Change its position depending on the angle, using cos/sin-math
        Man.X = Man.X + CT(Man.Angle) * 10
        Man.Y = Man.Y + ST(Man.Angle) * 10
        'Update map
        bMAPChanged = True
    'If the man is to walk backwards
    ElseIf Man.WalkBwrd Then
        'Check for outside-walls
        If ((Man.X - CT(Man.Angle) * 10) / Options.GRID_SIZE) < 0 Or ((Man.X - CT(Man.Angle) * 10) / Options.GRID_SIZE) > Lvl.WORLD_SIZE Or _
           ((Man.Y - ST(Man.Angle) * 10) / Options.GRID_SIZE) < 0 Or ((Man.Y - ST(Man.Angle) * 10) / Options.GRID_SIZE) > Lvl.WORLD_SIZE Then Exit Sub
        'Change its position depending on the angle, using cos/sin-math
        Man.X = Man.X - CT(Man.Angle) * 10
        Man.Y = Man.Y - ST(Man.Angle) * 10
        'Update the map
        bMAPChanged = True
    End If
End Sub

Public Sub MainLoop()

'The framecount
Dim FPS As Integer
'LastTick
Dim LT As Long
Dim i As Integer

'This is the core of the game - everything happens from here

Do

    'If more than 500 milliseconds(.5 sec) has passed since last time
    If GetTickCount() - LT >= 500 Then
        'Save the current tick
        LT = GetTickCount()
        'Print the framerate
        frmMain.Caption = "RayCasting at " & FPS & " frames/sec"
        'Reset the counter
        FPS = 0
    End If
    
    'Check for user-input
    Call DoKeys
    
    'Move the man, maybe
    Call MoveMan
    
    If bShowMap Then
        'Update the map, sometimes
        Call DoMAP
    End If

    'Cast the rays, the actual rendering
    Call CastRays
    
    'Increase this by 2, so we get a higher
    'update interval without have to multiply! (That slows everything down!)
    FPS = FPS + 2
    
    'So the changes will show, otherwise windows will crash, as usual!! :o)
    For i = 0 To Options.SPEED: DoEvents: Next i

Loop

End Sub

Private Sub DoMAP()

'If the map doesn't need to be updated, why do it?
If Not bMAPChanged Then Exit Sub

'Paint over the last frame with the level seen from top, that
'we drew in the LoadLvl-procedure
BitBlt frmMain.picMAP.hdc, 0, 0, Lvl.WORLD_SIZE * 10, Lvl.WORLD_SIZE * 10, frmMain.picMAP_BB.hdc, 0, 0, SRCCOPY

'Check for overflow
HMM:
If Man.Angle < 0 Then Man.Angle = Man.Angle + 360: GoTo HMM
If Man.Angle > 360 Then Man.Angle = Man.Angle - 360: GoTo HMM


'Define the man
Dim lpMan(2) As sPOINTAPI

'       lp1
'      /
'lp0 < v=FOV
'      \
'       lp2

'lp0
lpMan(0).X = Man.X / Options.GRID_SIZE * 10
lpMan(0).Y = Man.Y / Options.GRID_SIZE * 10
'lp1
lpMan(1).X = lpMan(0).X + CT(Man.Angle - Options.FOVdiv2) * 3
lpMan(1).Y = lpMan(0).Y + ST(Man.Angle - Options.FOVdiv2) * 3
'lp2
lpMan(2).X = lpMan(0).X + CT(Man.Angle + Options.FOVdiv2) * 4
lpMan(2).Y = lpMan(0).Y + ST(Man.Angle + Options.FOVdiv2) * 4

'Change the color so the man will be drawn in its own color
frmMain.picMAP.ForeColor = Options.MAN_CLR

'Move to point 1
MoveToEx frmMain.picMAP.hdc, lpMan(1).X, lpMan(1).Y, tmp
'Draw a line to point 0
LineTo frmMain.picMAP.hdc, lpMan(0).X, lpMan(0).Y
'And from there to point 2
LineTo frmMain.picMAP.hdc, lpMan(2).X, lpMan(2).Y

'picMAP.Line (lpMan(0).X, lpMan(0).Y)-(lpMan(1).X, lpMan(1).Y), vbGreen
'picMAP.Line (lpMan(0).X, lpMan(0).Y)-(lpMan(2).X, lpMan(2).Y), vbGreen

StretchBlt frmMain.picView.hdc, 0, 0, Options.G_WIDTH, Options.G_HEIGHT, frmMain.picMAP.hdc, 0, 0, Lvl.WORLD_SIZE * 10, Lvl.WORLD_SIZE * 10, SRCCOPY

'We just updated so we don't have to do it again
bMAPChanged = False

End Sub

Private Sub DoKeys()
'This procedure check for user-input by using
'the DirectInput-class that I've written.
'You don't have to know how DI works, the class
'is really easey to use!

'Check if the escape-button was pressed
If cDI.State.Key(DIK_ESCAPE) Then Call EndNow

'Check for right-arrow
If cDI.State.Key(DIK_RIGHT) Then
    Man.TurnRight = True
    Man.TurnLeft = False
'Check for left-arrow
ElseIf cDI.State.Key(DIK_LEFT) Then
    Man.TurnLeft = True
    Man.TurnRight = False
'None of the above, no turning
Else
    Man.TurnRight = False
    Man.TurnLeft = False
End If

'Check for up-arrow
If cDI.State.Key(DIK_UP) Then
    Man.WalkFwrd = True
    Man.WalkBwrd = False
'Check for down-arrow
ElseIf cDI.State.Key(DIK_DOWN) Then
    Man.WalkBwrd = True
    Man.WalkFwrd = False
'None of the above, no walking
Else
    Man.WalkFwrd = False
    Man.WalkBwrd = False
End If

If cDI.State.Key(DIK_TAB) Then
    bShowMap = True
Else
    bShowMap = False
End If

End Sub

Private Sub CastRays()
'This procedure renders the view.
'It casts a certain number of rays in different directions,
'and when a ray hits a wall, it draws a line that is higher
'the closer to the player it is, and smaller the further away!
'Isn't it clever?! The classical Doom and Wolfenstein 3D uses the
'very same technic!!

'First draw the ceiling..
StretchBlt frmMain.picBB.hdc, 0, 0, Options.G_WIDTH, Options.HORIZON, frmMain.picDummy.hdc, 0, 0, 1, 1, SRCCOPY
'..and the floor
StretchBlt frmMain.picBB.hdc, 0, Options.HORIZON, Options.G_WIDTH, Options.G_HEIGHT - Options.HORIZON, frmMain.picDummy.hdc, 1, 1, 1, 1, SRCCOPY

'BitBlt picBB.hdc, 0, 0, G_WIDTH, G_HEIGHT , picDummy.hdc, 0, 0, blackness

'ScreenX - where on the screen the current ray is to be drawn
Dim ScrX As Integer
'How much X is to be moved each time
Dim StepX As Single
'How much Y is to be moved each time
Dim StepY As Single
'This is the ray's coordinates
Dim X As Single
Dim Y As Single

'What type of cell did we hit?!
Dim Hit As Byte

'This is the angle of the current ray
Dim RA As Single 'RayAngle

'This is the length of the current ray
Dim Length As Integer
'This is the wall's height
Dim Height As Integer

'The color
Dim rgbClr As tRGBColor
'Used to shade
Dim Shade As Single
'The new color
Dim newRGB As Integer
'This too!
Dim Clr As Long

'Check for overflow
HMM:
If Man.Angle < 0 Then Man.Angle = Man.Angle + 360: GoTo HMM
If Man.Angle > 360 Then Man.Angle = Man.Angle - 360: GoTo HMM

'Begin casting at the left of the man's view
RA = Man.Angle - Options.FOVdiv2

'Loop through all the rays
For ScrX = 0 To Options.RAYS - 1

    'Get the start-position
    X = Man.X
    Y = Man.Y
    
    'Calculate how much the ray is to move each time useing cos/sin
    StepX = CT(RA) * Options.STEP
    StepY = ST(RA) * Options.STEP
    
    'Reset some variables
    'Length = 0
    Hit = 0
    
    'Move the ray until it hits a wall or it gets out of the eyes range
    Do
    
        'Move the ray a little bit
        X = X + StepX
        Y = Y + StepY
        
        'Increase the length
        'Length = Length + 1
        
        'Check if it is out of the eye's range
        'If Length > MAXDIST Then GoTo FAST
        
        'Check for overflow
        If X / Options.GRID_SIZE > Lvl.WORLD_SIZE - 1 Then GoTo FAST
        If Y / Options.GRID_SIZE > Lvl.WORLD_SIZE - 1 Then GoTo FAST
        
        'Get the current cell-value
        Hit = Lvl.Tile(X / Options.GRID_SIZE, Y / Options.GRID_SIZE)
    
    Loop While Hit = 0
    
    'Calculate the length to the hit
    Length = Sqr((Man.X - X) ^ 2 + (Man.Y - Y) ^ 2)
    
    'Check so it isn't out of the eye's range
    If Length > Options.MAXDIST Then GoTo FAST

    'Get the height from the look-up table
    Height = dHeight(Length)
    
    'Get how much of red, green and blue there is in the color of the wall we just hit
    rgbClr = GetRGBColor(Color(Hit))
    'Shade it so it gets darker the furhter away it is
    Shade = Length / Options.LIGHT
    
    '** Shade the different colors and do some error-checking **'
    newRGB = rgbClr.R / Shade
    If newRGB > 255 Then newRGB = 255
    rgbClr.R = newRGB
    
    newRGB = rgbClr.G / Shade
    If newRGB > 255 Then newRGB = 255
    rgbClr.G = newRGB
    
    newRGB = rgbClr.B / Shade
    If newRGB > 255 Then newRGB = 255
    rgbClr.B = newRGB
    '**.**'
    
    'This is our new color
    Clr = RGB(rgbClr.R, rgbClr.G, rgbClr.B)
    
    'Set the drawcolor to the one we just calculated
    frmMain.picBB.ForeColor = Clr
    'Move to the top of the current wall..
    MoveToEx frmMain.picBB.hdc, ScrX, Options.HORIZON - Height, tmp
    '..and draw a line to the bottom
    LineTo frmMain.picBB.hdc, ScrX, Options.HORIZON + Height
    
    'This is where we get if the ray was out of the eyes range
FAST:
    'Increase the ray's angle, so the next angle
    RA = RA + Options.RAY_INC

Next ScrX

If Not bShowMap Then
    'Blit everything onto the view-pic
    'BitBlt frmMain.picView.hdc, 0, 0, Options.G_WIDTH, Options.G_HEIGHT, frmMain.picBB.hdc, 0, 0, SRCCOPY
    StretchBlt frmMain.picView.hdc, 0, 0, Options.G_WIDTH, Options.G_HEIGHT, frmMain.picBB.hdc, 0, 0, Options.RAYS, 200, SRCCOPY
End If

End Sub

'*********************************'
'This proceure gets a long-color
'and from that, calculates how
'much red,green and blue there
'is in that color!
'*********************************'
Private Function GetRGBColor(ByVal lColor As Long) As tRGBColor
    GetRGBColor.R = lColor And 255 'Red
    GetRGBColor.G = (lColor And 65280) \ 256& 'Green
    GetRGBColor.B = (lColor And 16711680) \ 65535 'Blue
End Function
