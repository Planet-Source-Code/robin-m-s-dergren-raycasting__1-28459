Attribute VB_Name = "modDeclare"
Option Explicit

'A new instance of my DirectInput-object
Public cDI As New cDI

'Used to get the color of a specified pixel
'Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
'Used to set the color of a specified pixel
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
'Used to calculate the framerate
Public Declare Function GetTickCount Lib "kernel32" () As Long
'First move to a position with this..
Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
'..the draw a line from there to this position
Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
'Blit pictures
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'Blit pictures, but able to change the size of the pic, stretch!
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

'** BitBlt/StretchBlt constants **'
'Copy the sourceimage to the destination
Public Const SRCCOPY = &HCC0020
'Just draw a black square
Public Const BLACKNESS = &H42
'**.**'

Public Type tOptions
    RAYS            As Integer '320
    FOV             As Integer '60
    FOVdiv2         As Integer 'FOV/2
    RAY_INC         As Single 'FOV/RAYS
    STEP            As Single '1.85
    LIGHT           As Integer '120
    MAXDIST         As Integer '400
    WALLHEIGHT      As Integer '2560
    GRID_SIZE       As Integer '64
    G_WIDTH         As Integer '320
    G_HEIGHT        As Integer '200
    HORIZON         As Integer '120
    MAN_CLR         As Long '&H00ff00&
    CEILING_CLR     As Long '&H05232e
    FLOOR_CLR       As Long '&H479732
    SPEED           As Integer '0
End Type

'This is pi!!
Public Const PI = 3.14159265358979
'Multiply the number of degrees by RAD to get it in radians
Public Const RAD = PI / 180

'Sinus-Table, makes the game go MUCH faster
Public ST() As Single
'Cosinus-Table, same as above
Public CT() As Single

'Holds the heights the walls are at different distance
Public dHeight() As Single

'Some colors
Public Color(1 To 8) As Long

'A RGBcolor-type
Public Type tRGBColor
    R As Byte
    G As Byte
    B As Byte
End Type

'This holds the level
Public Type tLvl
    'The tiles/cells
    Tile()          As Byte
    WORLD_SIZE      As Integer
End Type

'This is the man
Public Type tMan
    X           As Single 'Its x-position
    Y           As Single 'Its y-position
    Angle       As Single 'Its angle in degrees
    TurnRight   As Boolean 'If it's to turn
    TurnLeft    As Boolean
    WalkFwrd    As Boolean 'or move
    WalkBwrd    As Boolean
End Type

'Just a coordinate type
Public Type sPOINTAPI
    X As Single
    Y As Single
End Type
'And the original one!
Public Type POINTAPI
    X As Long
    Y As Long
End Type

'This is the level
Public Lvl As tLvl
'And the man
Public Man As tMan
Public Options As tOptions

'If the map is to be updated
Public bMAPChanged As Boolean
'If the map is to be shown
Public bShowMap As Boolean

'Just a temporary variabel, never really used but is is necessary
Public tmp As POINTAPI
