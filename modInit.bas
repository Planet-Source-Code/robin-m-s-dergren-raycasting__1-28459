Attribute VB_Name = "modInit"
Option Explicit

Private Sub Main()

Call LoadOptions

Dim i As Integer

Call Calc_Tables

'The colors
For i = 1 To 8
    Color(i) = QBColor(i)
Next i

'Some default values
Man.TurnLeft = False
Man.TurnRight = False
Man.WalkBwrd = False
Man.WalkFwrd = False

'Set the colors of the ceiling and the floor
Call SetPixel(frmMain.picDummy.hdc, 0, 0, Options.CEILING_CLR)
Call SetPixel(frmMain.picDummy.hdc, 1, 1, Options.FLOOR_CLR)

'Initialize the DirectInput-stuff
Call cDI.Init(frmMain.hWnd)

'Change the size of the pic-boxes
frmMain.picView.Width = Options.G_WIDTH
frmMain.picView.Height = Options.G_HEIGHT
frmMain.picBB.Width = Options.RAYS
frmMain.picBB.Height = 200

'Change the size of the form
frmMain.Width = (Options.G_WIDTH + 21) * Screen.TwipsPerPixelX

'Show the form
frmMain.Show

bShowMap = False

'Load a level
Call LoadLvl("1.txt")

'Enter the main loop (the actual game)
Call MainLoop

End Sub

Private Sub Calc_Tables()

Dim i As Integer

ReDim dHeight(1 To Options.MAXDIST)

'Calculate the heights depending on the distance
For i = 1 To Options.MAXDIST
    dHeight(i) = Options.WALLHEIGHT / i
Next i

'Change the dimensions
ReDim ST(-Options.FOV To 360 + Options.FOV)
ReDim CT(-Options.FOV To 360 + Options.FOV)

'Calculate the Sin and Cos tables
For i = -Options.FOV To 360 + Options.FOV
    ST(i) = Sin(i * RAD)
    CT(i) = Cos(i * RAD)
Next i

End Sub

Public Sub LoadLvl(ByVal sLevel As String)

Dim Path As String

'The complete path to the level
Path = App.Path & "\LVL\" & sLevel

'If the level doesn't exist, get the hell out of here!
If Dir(Path) = "" Then Exit Sub

'Used to temporary hold the tile/cell-value
Dim tmp As String
'Wich of the cells we are at
Dim X As Integer, Y As Integer
'Just so we don't get a error, not that it happens so often
'but better safe than sorry!
Dim Free As Integer

'Get a number that is free
Free = FreeFile()
'Reset the coordinates, not really necessary
X = 0
Y = 0

'Open the level
Open Path For Input As #Free

    'Get the size of the level
    Input #Free, tmp
    Lvl.WORLD_SIZE = tmp
    
    'Change the size of the level
    ReDim Lvl.Tile(Lvl.WORLD_SIZE - 1, Lvl.WORLD_SIZE - 1) As Byte
    
    'Change the size of the map-picture
    frmMain.picMAP.Width = Lvl.WORLD_SIZE * 10
    frmMain.picMAP_BB.Width = Lvl.WORLD_SIZE * 10
    frmMain.picMAP.Height = Lvl.WORLD_SIZE * 10
    frmMain.picMAP_BB.Height = Lvl.WORLD_SIZE * 10
    
    'Change the height of the form
    frmMain.Height = (frmMain.picView.Height + 38) * Screen.TwipsPerPixelY
    
    'Loop through all the tiles/cells
    For Y = 0 To Lvl.WORLD_SIZE - 1
    For X = 0 To Lvl.WORLD_SIZE - 1
        
        'Get a value
        Input #Free, tmp
        
        '9=man's startposition
        If tmp = 9 Then
            Man.X = X * Options.GRID_SIZE
            Man.Y = Y * Options.GRID_SIZE
            'Change it to a floor-cell
            tmp = 0
        End If
        
        'Save the value
        Lvl.Tile(X, Y) = CByte(tmp)
        
        'If it isn't a floor, then draw a box on the map
        If tmp <> 0 Then
            frmMain.picMAP_BB.Line (X * 10, Y * 10)-(X * 10 + 10, Y * 10 + 10), Color(tmp), BF
        End If
    
    Next X
    Next Y

Close #Free

'This doesn't work at the time!!
BitBlt frmMain.picMAP.hdc, 0, 0, 100, 100, frmMain.picMAP_BB.hdc, 0, 0, SRCCOPY

End Sub

Public Sub DoDefaultOptions()
'Set the default settings
With Options

    .RAYS = 320 'The number of rays
    .FOV = 60 'The Field Of View in degrees
    .FOVdiv2 = 30 'Just to speed it up a little (FOV/2)
    .RAY_INC = 60 / 320 'How many degrees their are between each ray
    .STEP = 1.85 'A higher number will result in faster gameplay but lousy graphics
    .LIGHT = 120 'How much ligth
    .MAXDIST = 500 'How far the eye can see
    .WALLHEIGHT = 2560 'How high each wall is
    .GRID_SIZE = 64 'The size of each cell
    .G_WIDTH = 320 'How wide the view is
    .G_HEIGHT = 200 'How high the view is
    .HORIZON = 120 'Where the horizon is
    .MAN_CLR = &HFF00& 'The color of the man
    .CEILING_CLR = &H776666 'The color of the ceiling
    .FLOOR_CLR = &HAA9999 'The color of the floor
    .SPEED = 0 'The higher speed, the slower game

End With

End Sub

Public Sub SaveOptions()

Dim Free As Integer

'Get a number that isn't used
Free = FreeFile()

With Options

'Save all the options
Open App.Path & "\RAYS.INI" For Output As #Free
    Print #Free, "RAYS=" & .RAYS
    Print #Free, "FOV=" & .FOV
    Print #Free, "STEP=" & .STEP
    Print #Free, "LIGHT=" & .LIGHT
    Print #Free, "MAXDIST=" & .MAXDIST
    Print #Free, "WALLHEIGHT=" & .WALLHEIGHT
    Print #Free, "GRID_SIZE=" & .GRID_SIZE
    Print #Free, "G_WIDTH=" & .G_WIDTH
    Print #Free, "G_HEIGHT=" & .G_HEIGHT
    Print #Free, "HORIZON=" & .HORIZON
    Print #Free, "MAN_CLR=" & .MAN_CLR
    Print #Free, "CEILING_CLR=" & .CEILING_CLR
    Print #Free, "FLOOR_CLR=" & .FLOOR_CLR
    Print #Free, "SPEED=" & .SPEED
Close #Free

End With

End Sub

Public Sub LoadOptions()

'If the file doesn't exist, set default settings
If Dir(App.Path & "\RAYS.INI") = "" Then Call DoDefaultOptions: Exit Sub

Dim Free As Integer
Dim tmp As String

'Get a number that isn't used
Free = FreeFile()

With Options

'This just loads all the settings
Open App.Path & "\RAYS.INI" For Input As #Free
    Do
        Line Input #Free, tmp
        Select Case UCase(Left(tmp, InStr(1, tmp, "=") - 1))
            Case "RAYS": .RAYS = Right(tmp, Len(tmp) - InStr(1, tmp, "="))
            Case "FOV": .FOV = Right(tmp, Len(tmp) - InStr(1, tmp, "="))
            Case "STEP": .STEP = Right(tmp, Len(tmp) - InStr(1, tmp, "="))
            Case "LIGHT": .LIGHT = Right(tmp, Len(tmp) - InStr(1, tmp, "="))
            Case "MAXDIST": .MAXDIST = Right(tmp, Len(tmp) - InStr(1, tmp, "="))
            Case "WALLHEIGHT": .WALLHEIGHT = Right(tmp, Len(tmp) - InStr(1, tmp, "="))
            Case "GRID_SIZE": .GRID_SIZE = Right(tmp, Len(tmp) - InStr(1, tmp, "="))
            Case "G_WIDTH": .G_WIDTH = Right(tmp, Len(tmp) - InStr(1, tmp, "="))
            Case "G_HEIGHT": .G_HEIGHT = Right(tmp, Len(tmp) - InStr(1, tmp, "="))
            Case "HORIZON": .HORIZON = Right(tmp, Len(tmp) - InStr(1, tmp, "="))
            Case "MAN_CLR": .MAN_CLR = Right(tmp, Len(tmp) - InStr(1, tmp, "="))
            Case "CEILING_CLR": .CEILING_CLR = Right(tmp, Len(tmp) - InStr(1, tmp, "="))
            Case "FLOOR_CLR": .FLOOR_CLR = Right(tmp, Len(tmp) - InStr(1, tmp, "="))
            Case "SPEED": .SPEED = Right(tmp, Len(tmp) - InStr(1, tmp, "="))
        End Select
    Loop Until EOF(Free)
Close #Free

'Set some variables
.FOVdiv2 = .FOV / 2
.RAY_INC = .FOV / .RAYS

End With

End Sub

Public Sub EndNow()
    Call SaveOptions
    End
End Sub
