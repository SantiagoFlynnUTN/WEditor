Attribute VB_Name = "modDXEngine"
Option Explicit

'Private Const DegreeToRadian As Single = 0.0174532925

'***************************
'Estructures
'***************************
Public polygonCount(1) As Single
Public PolygonX As Single
Public POlyGonXMin As Single
Public PREVIEW_INDEX As Integer
Public Type tNewIndice
    X As Integer
    Y As Integer
    Grafico As Integer
End Type
Public Type tNewAnimation
    Numero As Integer
    Grafico As Integer
    NumFrames As Byte
    Filas As Byte
    Columnas As Byte
    Indice() As tNewIndice
    Width As Integer
    Height As Integer
    IndiceCounter As Single
    Velocidad As Single
    TileWidth As Single
    TileHeight As Single
    
    Romboidal As Byte
    Direction As Integer
    
    OffsetX As Integer
    OffsetY As Integer
    Initial As Integer
    TipoAnimacion As Byte
End Type


Public AnimacionPrueba As tNewAnimation
Public Num_NwAnim As Integer
Public NewAnimationData() As tNewAnimation
Enum eSPOT_LIGHT_COLOR
    eSPOT_WHITE = 1
    eSPOT_RED
    eSPOT_BLUE
    eSPOT_GREEN
    eSPOT_YELLOW
    eSPOT_BLACK
    eSPOT_MATE
    eSPOT_CUSTOM = 99
End Enum

Type tSPOT_LIGHTS
    OffsetX As Integer
    OffsetY As Integer
    Mx As Byte
    My As Byte
    
    Mustbe_Render As Boolean
    
    X As Integer
    Y As Integer
    CHarIndex As Integer
    BIND_TO As Byte ' 0PANTALLA, 1MAPA, 2CHARINDEX
    
    SPOT_COLOR_BASE As Byte
    SPOT_COLOR_EXTRA As Byte
    SPOT_TIPO As Integer 'SIZE AND SHAPE-> ANIMATION.
    
    Color As Long
    COLOR_EXTRA As Long
    INDEX_IN_COL As Byte
    
    INTENSITY As Byte
    
    EXTRA_GRAFICO As Integer
    Grafico As Integer
    
    Anim As tNewAnimation
End Type

Public SPOTLIGHTS_COLORES() As Long
Public NUM_SPOTLIGHTS_COLORES As Byte

Public NUM_SPOTLIGHTS_ANIMATION As Byte
Public SPOTLIGHTS_ANIMATION() As Integer


Public SCREEN_SPOTS As New Collection
Public SPOT_LIGHTS() As tSPOT_LIGHTS
Public Num_SPOTLIGHTS As Integer
'SISTEMA DE INTERIORES.

Public S As Direct3DSurface8
Public ambient_light() As Long
Public base_light As Long
Public day_r_old As Byte
Public day_g_old As Byte
Public day_b_old As Byte
Type luzxhora
    R As Long
    G As Long
    B As Long
End Type
Public luz_dia(0 To 24) As luzxhora '¬¬ la hora 24 dura 1 minuto entre las 24 y las 0



Public extra_light(10) As Long
Public HoraLuz As Integer
Public Type Particle 'LEAN_PART
    friction As Single
    X As Single
    Y As Single
    vector_x As Single
    vector_y As Single
    AngleC As Single
    index As Integer
    fC As Single
    alive_counter As Long
    x1 As Integer
    x2 As Integer
    y1 As Integer
    y2 As Integer
    vecx1 As Integer
    vecx2 As Integer
    vecy1 As Integer
    vecy2 As Integer
    life1 As Long
    life2 As Long
    fric As Integer
    spin_speedL As Single
    spin_speedH As Single
    gravity As Boolean
    grav_strength As Long
    bounce_strength As Long
    Spin As Boolean
    XMove As Boolean
    YMove As Boolean
    move_x1 As Integer
    move_x2 As Integer
    move_y1 As Integer
    move_y2 As Integer
    rgb_list(0 To 3) As Long
    grh_resize As Boolean
    grh_resizex As Integer
    grh_resizey As Integer
End Type
 
'Modified by: Ryan Cain (Onezero)
'Last modify date: 5/14/2003
Public Type particle_group 'LEAN_PART
    Active As Boolean
    id As Long
    map_x As Integer
    map_y As Integer
    char_index As Long
    frame_counter As Single
    frame_speed As Single
    stream_type As Byte
    particle_stream() As Particle
    particle_count As Long
    grh_index_list() As Long
    grh_index_count As Long
    alpha_blend As Boolean
    alive_counter As Long
    never_die As Boolean
    x1 As Integer
    x2 As Integer
    y1 As Integer
    y2 As Integer
    Angle As Integer
    vecx1 As Integer
    vecx2 As Integer
    vecy1 As Integer
    vecy2 As Integer
    life1 As Long
    life2 As Long
    fric As Long
    spin_speedL As Single
    spin_speedH As Single
    gravity As Boolean
    grav_strength As Long
    bounce_strength As Long
    Spin As Boolean
    XMove As Boolean
    YMove As Boolean
    move_x1 As Integer
    move_x2 As Integer
    move_y1 As Integer
    move_y2 As Integer
    rgb_list(0 To 3) As Long
    Speed As Single
    life_counter As Long
    grh_resize As Boolean
    grh_resizex As Integer
    grh_resizey As Integer
End Type


'This structure describes a transformed and lit vertex.
Public Al(1 To 39) As D3DCOLORVALUE
Public ALc(1 To 39) As Long
Public ALH As Byte

Public Lucez(1 To 20) As Long
Public Type TLVERTEX
    X As Single
    Y As Single
    z As Single
    rhw As Single
    Color As Long
    tu As Single
    tv As Single
End Type

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
ByRef destination As Any, ByRef Source As Any, ByVal numbytes As Long)
Public Const D3DGOLD As Long = -2645468
Public Const D3DWHITE As Long = -1
Public Const D3DRED As Long = -65536

 
'Describes the return from a texture init
Private Type D3DXIMAGE_INFO_A
    Width As Long
    Height As Long
    Depth As Long
    MipLevels As Long
    Format As CONST_D3DFORMAT
    ResourceType As CONST_D3DRESOURCETYPE
    ImageFileFormat As Long
End Type
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Type CharVA
    Vertex(0 To 3) As TLVERTEX
End Type
Private Type VFH
    BitmapWidth As Long         'Size of the bitmap itself
    BitmapHeight As Long
    CellWidth As Long           'Size of the cells (area for each character)
    CellHeight As Long
    BaseCharOffset As Byte      'The character we start from
    CharWidth(0 To 255) As Byte 'The actual factual width of each character
    CharVA(0 To 255) As CharVA
End Type
Private Type CustomFont
    HeaderInfo As VFH           'Holds the header information
    Texture As Direct3DTexture8 'Holds the texture of the text
    RowPitch As Integer         'Number of characters per row
    RowFactor As Single         'Percentage of the texture width each character takes
    ColFactor As Single         'Percentage of the texture height each character takes
    CharHeight As Byte          'Height to use for the text - easiest to start with CellHeight value, and keep lowering until you get a good value
    TextureSize As POINTAPI     'Size of the texture
End Type
 
'Private Const Font_Default_TextureNum As Long = -1   'The texture number used to represent this font - only used for AlternateRendering - keep negative to prevent interfering with game textures
Private cfonts(1 To 2) As CustomFont ' _Default2 As CustomFont
 

Private Type tGraphicChar
    Src_X As Integer
    Src_Y As Integer
End Type

Private Type tGraphicFont
    texture_index As Long
    Caracteres(0 To 255) As tGraphicChar 'Ascii Chars
    Char_Size As Byte 'In pixels
End Type

Private Type DXFont
    dFont As D3DXFont
    size As Integer
End Type


Public Enum FontAlignment
    fa_center = DT_CENTER
    fa_top = DT_TOP
    fa_left = DT_LEFT
    fa_topleft = DT_TOP Or DT_LEFT
    fa_bottomleft = DT_BOTTOM Or DT_LEFT
    fa_bottom = DT_BOTTOM
    fa_right = DT_RIGHT
    fa_bottomright = DT_BOTTOM Or DT_RIGHT
    fa_topright = DT_TOP Or DT_RIGHT
End Enum

'***************************
'Variables
'***************************
'Major DX Objects
Public dx As DirectX8
Public d3d As Direct3D8
Public ddevice As Direct3DDevice8
Public d3dx As D3DX8

Dim d3dpp As D3DPRESENT_PARAMETERS

'Texture Manager for Dinamic Textures
Public DXPool As clsTextureManager

'Main form handle
Dim form_hwnd As Long

'Display variables
Dim screen_hwnd As Long

'FPS Counters
Dim fps_last_time As Long 'When did we last check the frame rate?
Dim fps_frame_counter As Long 'How many frames have been drawn
Dim FPS As Long 'What the current frame rate is.....

Dim engine_render_started As Boolean

'Graphic Font List
Dim gfont_list() As tGraphicFont
Dim gfont_count As Long

'Font List
Private font_list() As DXFont
Private font_count As Integer


'***************************
'Constants
'***************************
'Engine
Private Const COLOR_KEY As Long = &HFF000000
Private Const FVF = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE
Private Const FVF2 = D3DFVF_XYZRHW Or D3DFVF_DIFFUSE Or D3DFVF_TEX2
Private Const PI As Single = 3.14159265358979

'Old fashion BitBlt functions
Private Const SRCCOPY = &HCC0020

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function TransparentBlt Lib "msimg32" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hdcsrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal crTransparent As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long


 
'particulas
Dim base_tile_size As Integer
 
 
Dim particle_group_list() As particle_group
Dim particle_group_count As Long
Dim particle_group_last As Long
'Particle Groups
Public TotalStreams As Integer
Public StreamData() As Stream
 
'RGB Type
Private Type RGB
    R As Long
    G As Long
    B As Long
End Type
 
Public Type Stream
    Name As String
    NumOfParticles As Long
    NumTrueParticles As Long
    NumGrhs As Long
    id As Long
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
    Angle As Long
    vecx1 As Long
    vecx2 As Long
    vecy1 As Long
    vecy2 As Long
    life1 As Long
    life2 As Long
    friction As Long
    Spin As Byte
    spin_speedL As Single
    spin_speedH As Single
    AlphaBlend As Byte
    gravity As Byte
    grav_strength As Long
    bounce_strength As Long
    XMove As Byte
    YMove As Byte
    move_x1 As Long
    move_x2 As Long
    move_y1 As Long
    move_y2 As Long
    grh_list() As Long
    colortint(0 To 3) As RGB
    Speed As Single
    life_counter As Long
    grh_resize As Boolean
    grh_resizex As Integer
    grh_resizey As Integer
End Type

 
'Old fashion BitBlt function
'Added by Juan Martín Sotuyo Dodero
Public Declare Function StretchBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
 Function GetVar(ByVal file As String, ByVal Main As String, ByVal Var As String) As String
      '*****************************************************************
      'Gets a Var from a text file
      '*****************************************************************

      Dim sSpaces As String ' This will hold the input that the program will retrieve
    
      sSpaces = Space$(500) ' This tells the computer how long the longest string can be. If you want, you can change the number 100 to any number you wish
    
      GetPrivateProfileString Main, Var, vbNullString, sSpaces, Len(sSpaces), file
    
      GetVar = RTrim$(sSpaces)
      GetVar = left$(GetVar, Len(GetVar) - 1)
End Function
Sub CargarParticulas()
 
Dim StreamFile As String
Dim LoopC As Long
Dim i As Long
Dim GrhListing As String
Dim TempSet As String
Dim ColorSet As Long
   
StreamFile = App.PATH & "\Resources\INIT\Particles.ini"
TotalStreams = Val(GetVar(StreamFile, "INIT", "Total"))
 
'resize StreamData array
ReDim StreamData(1 To TotalStreams) As Stream
 
    'fill StreamData array with info from Particles.ini
    For LoopC = 1 To TotalStreams
        StreamData(LoopC).Name = GetVar(StreamFile, Val(LoopC), "Name")
        StreamData(LoopC).NumOfParticles = GetVar(StreamFile, Val(LoopC), "NumOfParticles")
        StreamData(LoopC).x1 = GetVar(StreamFile, Val(LoopC), "X1")
        StreamData(LoopC).y1 = GetVar(StreamFile, Val(LoopC), "Y1")
        StreamData(LoopC).x2 = GetVar(StreamFile, Val(LoopC), "X2")
        StreamData(LoopC).y2 = GetVar(StreamFile, Val(LoopC), "Y2")
        StreamData(LoopC).Angle = GetVar(StreamFile, Val(LoopC), "Angle")
        StreamData(LoopC).vecx1 = GetVar(StreamFile, Val(LoopC), "VecX1")
        StreamData(LoopC).vecx2 = GetVar(StreamFile, Val(LoopC), "VecX2")
        StreamData(LoopC).vecy1 = GetVar(StreamFile, Val(LoopC), "VecY1")
        StreamData(LoopC).vecy2 = GetVar(StreamFile, Val(LoopC), "VecY2")
        StreamData(LoopC).life1 = GetVar(StreamFile, Val(LoopC), "Life1")
        StreamData(LoopC).life2 = GetVar(StreamFile, Val(LoopC), "Life2")
        StreamData(LoopC).friction = GetVar(StreamFile, Val(LoopC), "Friction")
        StreamData(LoopC).Spin = GetVar(StreamFile, Val(LoopC), "Spin")
        StreamData(LoopC).spin_speedL = GetVar(StreamFile, Val(LoopC), "Spin_SpeedL")
        StreamData(LoopC).spin_speedH = GetVar(StreamFile, Val(LoopC), "Spin_SpeedH")
        StreamData(LoopC).AlphaBlend = GetVar(StreamFile, Val(LoopC), "AlphaBlend")
        StreamData(LoopC).gravity = GetVar(StreamFile, Val(LoopC), "Gravity")
        StreamData(LoopC).grav_strength = GetVar(StreamFile, Val(LoopC), "Grav_Strength")
        StreamData(LoopC).bounce_strength = GetVar(StreamFile, Val(LoopC), "Bounce_Strength")
        StreamData(LoopC).XMove = GetVar(StreamFile, Val(LoopC), "XMove")
        StreamData(LoopC).YMove = GetVar(StreamFile, Val(LoopC), "YMove")
        StreamData(LoopC).move_x1 = GetVar(StreamFile, Val(LoopC), "move_x1")
        StreamData(LoopC).move_x2 = GetVar(StreamFile, Val(LoopC), "move_x2")
        StreamData(LoopC).move_y1 = GetVar(StreamFile, Val(LoopC), "move_y1")
        StreamData(LoopC).move_y2 = GetVar(StreamFile, Val(LoopC), "move_y2")
        StreamData(LoopC).life_counter = GetVar(StreamFile, Val(LoopC), "life_counter")
        StreamData(LoopC).Speed = Val(GetVar(StreamFile, Val(LoopC), "Speed"))
        StreamData(LoopC).grh_resize = Val(GetVar(StreamFile, Val(LoopC), "resize"))
        StreamData(LoopC).grh_resizex = Val(GetVar(StreamFile, Val(LoopC), "rx"))
        StreamData(LoopC).grh_resizey = Val(GetVar(StreamFile, Val(LoopC), "ry"))
        StreamData(LoopC).NumGrhs = GetVar(StreamFile, Val(LoopC), "NumGrhs")
       
        ReDim StreamData(LoopC).grh_list(1 To StreamData(LoopC).NumGrhs)
        GrhListing = GetVar(StreamFile, Val(LoopC), "Grh_List")
       
        For i = 1 To StreamData(LoopC).NumGrhs
            StreamData(LoopC).grh_list(i) = ReadField(Str(i), GrhListing, 44)
        Next i
        StreamData(LoopC).grh_list(i - 1) = StreamData(LoopC).grh_list(i - 1)
        For ColorSet = 1 To 4
            TempSet = GetVar(StreamFile, Val(LoopC), "ColorSet" & ColorSet)
            StreamData(LoopC).colortint(ColorSet - 1).R = ReadField(1, TempSet, 44)
            StreamData(LoopC).colortint(ColorSet - 1).G = ReadField(2, TempSet, 44)
            StreamData(LoopC).colortint(ColorSet - 1).B = ReadField(3, TempSet, 44)
        Next ColorSet
    Next LoopC
 
End Sub
 
Public Function General_Particle_Create(ByVal ParticulaInd As Long, ByVal X As Integer, ByVal Y As Integer, Optional ByVal particle_life As Long = 0) As Long

Dim rgb_list(0 To 3) As Long
rgb_list(0) = RGB(StreamData(ParticulaInd).colortint(0).R, StreamData(ParticulaInd).colortint(0).G, StreamData(ParticulaInd).colortint(0).B)
rgb_list(1) = RGB(StreamData(ParticulaInd).colortint(1).R, StreamData(ParticulaInd).colortint(1).G, StreamData(ParticulaInd).colortint(1).B)
rgb_list(2) = RGB(StreamData(ParticulaInd).colortint(2).R, StreamData(ParticulaInd).colortint(2).G, StreamData(ParticulaInd).colortint(2).B)
rgb_list(3) = RGB(StreamData(ParticulaInd).colortint(3).R, StreamData(ParticulaInd).colortint(3).G, StreamData(ParticulaInd).colortint(3).B)
 
General_Particle_Create = Particle_Group_Create(X, Y, StreamData(ParticulaInd).grh_list, rgb_list(), StreamData(ParticulaInd).NumOfParticles, ParticulaInd, _
    StreamData(ParticulaInd).AlphaBlend, IIf(particle_life = 0, StreamData(ParticulaInd).life_counter, particle_life), StreamData(ParticulaInd).Speed, , StreamData(ParticulaInd).x1, StreamData(ParticulaInd).y1, StreamData(ParticulaInd).Angle, _
    StreamData(ParticulaInd).vecx1, StreamData(ParticulaInd).vecx2, StreamData(ParticulaInd).vecy1, StreamData(ParticulaInd).vecy2, _
    StreamData(ParticulaInd).life1, StreamData(ParticulaInd).life2, StreamData(ParticulaInd).friction, StreamData(ParticulaInd).spin_speedL, _
    StreamData(ParticulaInd).gravity, StreamData(ParticulaInd).grav_strength, StreamData(ParticulaInd).bounce_strength, StreamData(ParticulaInd).x2, _
    StreamData(ParticulaInd).y2, StreamData(ParticulaInd).XMove, StreamData(ParticulaInd).move_x1, StreamData(ParticulaInd).move_x2, StreamData(ParticulaInd).move_y1, _
    StreamData(ParticulaInd).move_y2, StreamData(ParticulaInd).YMove, StreamData(ParticulaInd).spin_speedH, StreamData(ParticulaInd).Spin, StreamData(ParticulaInd).grh_resize, StreamData(ParticulaInd).grh_resizex, StreamData(ParticulaInd).grh_resizey)
 
End Function
Function ReadField(ByVal Pos As Integer, _
                   ByRef Text As String, _
                   ByVal SepASCII As Byte) As String
      '*****************************************************************
      'Gets a field from a delimited string
      'Author: Juan Martín Sotuyo Dodero (Maraxus)
      'Last Modify Date: 11/15/2004
      '*****************************************************************

      Dim i          As Long
      Dim lastPos    As Long
      Dim CurrentPos As Long
      Dim delimiter  As String * 1
    
      delimiter = Chr$(SepASCII)
    
      For i = 1 To Pos
            lastPos = CurrentPos
            CurrentPos = InStr(lastPos + 1, Text, delimiter, vbBinaryCompare)

      Next i
    
      If CurrentPos = 0 Then
            ReadField = mid$(Text, lastPos + 1, Len(Text) - lastPos)
      Else
            ReadField = mid$(Text, lastPos + 1, CurrentPos - lastPos - 1)
      End If

End Function

Public Function General_Char_Particle_Create(ByVal ParticulaInd As Long, ByVal char_index As Integer, Optional ByVal particle_life As Long = 0) As Long
 
Dim rgb_list(0 To 3) As Long
rgb_list(0) = RGB(StreamData(ParticulaInd).colortint(0).R, StreamData(ParticulaInd).colortint(0).G, StreamData(ParticulaInd).colortint(0).B)
rgb_list(1) = RGB(StreamData(ParticulaInd).colortint(1).R, StreamData(ParticulaInd).colortint(1).G, StreamData(ParticulaInd).colortint(1).B)
rgb_list(2) = RGB(StreamData(ParticulaInd).colortint(2).R, StreamData(ParticulaInd).colortint(2).G, StreamData(ParticulaInd).colortint(2).B)
rgb_list(3) = RGB(StreamData(ParticulaInd).colortint(3).R, StreamData(ParticulaInd).colortint(3).G, StreamData(ParticulaInd).colortint(3).B)
 
General_Char_Particle_Create = Char_Particle_Group_Create(char_index, StreamData(ParticulaInd).grh_list, rgb_list(), StreamData(ParticulaInd).NumOfParticles, ParticulaInd, _
    StreamData(ParticulaInd).AlphaBlend, IIf(particle_life = 0, StreamData(ParticulaInd).life_counter, particle_life), StreamData(ParticulaInd).Speed, , StreamData(ParticulaInd).x1, StreamData(ParticulaInd).y1, StreamData(ParticulaInd).Angle, _
    StreamData(ParticulaInd).vecx1, StreamData(ParticulaInd).vecx2, StreamData(ParticulaInd).vecy1, StreamData(ParticulaInd).vecy2, _
    StreamData(ParticulaInd).life1, StreamData(ParticulaInd).life2, StreamData(ParticulaInd).friction, StreamData(ParticulaInd).spin_speedL, _
    StreamData(ParticulaInd).gravity, StreamData(ParticulaInd).grav_strength, StreamData(ParticulaInd).bounce_strength, StreamData(ParticulaInd).x2, _
    StreamData(ParticulaInd).y2, StreamData(ParticulaInd).XMove, StreamData(ParticulaInd).move_x1, StreamData(ParticulaInd).move_x2, StreamData(ParticulaInd).move_y1, _
    StreamData(ParticulaInd).move_y2, StreamData(ParticulaInd).YMove, StreamData(ParticulaInd).spin_speedH, StreamData(ParticulaInd).Spin, StreamData(ParticulaInd).grh_resize, StreamData(ParticulaInd).grh_resizex, StreamData(ParticulaInd).grh_resizey)
 
End Function
 
'----------------------------------------PARTICULAS---------------------------------
Public Sub Convert_Heading_to_Direction(ByVal Heading As Long, ByRef direction_x As Integer, ByRef direction_y As Integer)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim addy As Long
    Dim addx As Long
   
    'Figure out which way to move
    Select Case Heading
        Case 1
            addy = -1
   
        Case 2
            addx = 1
   
        Case 3
            addy = 1
           
        Case 4
            addx = -1
   
    End Select
   
    direction_x = direction_x + addx
    direction_y = direction_y + addy
End Sub
 
 
Public Function Particle_Group_Next_Open() As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim LoopC As Long
   
    LoopC = 1
    Do Until particle_group_list(LoopC).Active = False
        If LoopC = particle_group_last Then
            Particle_Group_Next_Open = particle_group_last + 1
            Exit Function
        End If
        LoopC = LoopC + 1
    Loop
   
    Particle_Group_Next_Open = LoopC
Exit Function
ErrorHandler:
    Particle_Group_Next_Open = 1
End Function
Public Function Particle_Group_Create(ByVal map_x As Integer, ByVal map_y As Integer, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                        Optional ByVal particle_count As Long = 20, Optional ByVal stream_type As Long = 1, _
                                        Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                        Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
                                        Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal Angle As Integer, _
                                        Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                        Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                        Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                        Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                        Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                        Optional bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                                        Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                        Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                        Optional ByVal spin_speedH As Single, Optional ByVal Spin As Boolean, Optional grh_resize As Boolean, _
                                        Optional grh_resizex As Integer, Optional grh_resizey As Integer) As Long
   
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 12/15/2002
'Returns the particle_group_index if successful, else 0
'**************************************************************
    If (map_x <> -1) And (map_y <> -1) Then
    If Map_Particle_Group_Get(map_x, map_y) = 0 Then
        Particle_Group_Create = Particle_Group_Next_Open
        Particle_Group_Make Particle_Group_Create, map_x, map_y, particle_count, stream_type, grh_index_list(), rgb_list(), alpha_blend, alive_counter, frame_speed, id, x1, y1, Angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, x2, y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, Spin, grh_resize, grh_resizex, grh_resizey
    Else
        Particle_Group_Create = Particle_Group_Next_Open
        Particle_Group_Make Particle_Group_Create, map_x, map_y, particle_count, stream_type, grh_index_list(), rgb_list(), alpha_blend, alive_counter, frame_speed, id, x1, y1, Angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, x2, y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, Spin, grh_resize, grh_resizex, grh_resizey
    End If
    End If
End Function
 
Public Function Particle_Group_Remove(ByVal particle_group_index As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    'Make sure it's a legal index
    If Particle_Group_Check(particle_group_index) Then
        Particle_Group_Destroy particle_group_index
        Particle_Group_Remove = True
    End If
End Function
 
Public Function Particle_Group_Remove_All() As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    Dim index As Long
   
    For index = 1 To particle_group_last
        'Make sure it's a legal index
        If Particle_Group_Check(index) Then
            Particle_Group_Destroy index
        End If
    Next index
   
    Particle_Group_Remove_All = True
End Function
 
Public Function Particle_Group_Find(ByVal id As Long) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'Find the index related to the handle
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim LoopC As Long
   
    LoopC = 1
    Do Until particle_group_list(LoopC).id = id
        If LoopC = particle_group_last Then
            Particle_Group_Find = 0
            Exit Function
        End If
        LoopC = LoopC + 1
    Loop
   
    Particle_Group_Find = LoopC
Exit Function
ErrorHandler:
    Particle_Group_Find = 0
End Function
 
Public Sub Particle_Group_Make(ByVal particle_group_index As Long, ByVal map_x As Integer, ByVal map_y As Integer, _
                                ByVal particle_count As Long, ByVal stream_type As Long, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
                                Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal Angle As Integer, _
                                Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                Optional bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                                Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                Optional ByVal spin_speedH As Single, Optional ByVal Spin As Boolean, Optional grh_resize As Boolean, _
                                Optional grh_resizex As Integer, Optional grh_resizey As Integer)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Makes a new particle effect
'*****************************************************************
    'Update array size
    If particle_group_index > particle_group_last Then
        particle_group_last = particle_group_index
        ReDim Preserve particle_group_list(1 To particle_group_last)
    End If
    particle_group_count = particle_group_count + 1
   
    'Make active
    particle_group_list(particle_group_index).Active = True
   
    'Map pos
    If (map_x <> -1) And (map_y <> -1) Then
        particle_group_list(particle_group_index).map_x = map_x
        particle_group_list(particle_group_index).map_y = map_y
    End If
   
    'Grh list
    ReDim particle_group_list(particle_group_index).grh_index_list(1 To UBound(grh_index_list))
    particle_group_list(particle_group_index).grh_index_list() = grh_index_list()
    particle_group_list(particle_group_index).grh_index_count = UBound(grh_index_list)
   
    'Sets alive vars
    If alive_counter = -1 Then
        particle_group_list(particle_group_index).alive_counter = -1
        particle_group_list(particle_group_index).never_die = True
    Else
        particle_group_list(particle_group_index).alive_counter = alive_counter
        particle_group_list(particle_group_index).never_die = False
    End If
   
    'alpha blending
    particle_group_list(particle_group_index).alpha_blend = alpha_blend
   
    'stream type
    particle_group_list(particle_group_index).stream_type = stream_type
   
    'speed
    particle_group_list(particle_group_index).frame_speed = frame_speed
   
    particle_group_list(particle_group_index).x1 = x1
    particle_group_list(particle_group_index).y1 = y1
    particle_group_list(particle_group_index).x2 = x2
    particle_group_list(particle_group_index).y2 = y2
    particle_group_list(particle_group_index).Angle = Angle
    particle_group_list(particle_group_index).vecx1 = vecx1
    particle_group_list(particle_group_index).vecx2 = vecx2
    particle_group_list(particle_group_index).vecy1 = vecy1
    particle_group_list(particle_group_index).vecy2 = vecy2
    particle_group_list(particle_group_index).life1 = life1
    particle_group_list(particle_group_index).life2 = life2
    particle_group_list(particle_group_index).fric = fric
    particle_group_list(particle_group_index).Spin = Spin
    particle_group_list(particle_group_index).spin_speedL = spin_speedL
    particle_group_list(particle_group_index).spin_speedH = spin_speedH
    particle_group_list(particle_group_index).gravity = gravity
    particle_group_list(particle_group_index).grav_strength = grav_strength
    particle_group_list(particle_group_index).bounce_strength = bounce_strength
    particle_group_list(particle_group_index).XMove = XMove
    particle_group_list(particle_group_index).YMove = YMove
    particle_group_list(particle_group_index).move_x1 = move_x1
    particle_group_list(particle_group_index).move_x2 = move_x2
    particle_group_list(particle_group_index).move_y1 = move_y1
    particle_group_list(particle_group_index).move_y2 = move_y2
   
    particle_group_list(particle_group_index).rgb_list(0) = rgb_list(0)
    particle_group_list(particle_group_index).rgb_list(1) = rgb_list(1)
    particle_group_list(particle_group_index).rgb_list(2) = rgb_list(2)
    particle_group_list(particle_group_index).rgb_list(3) = rgb_list(3)
   
    particle_group_list(particle_group_index).grh_resize = grh_resize
    particle_group_list(particle_group_index).grh_resizex = grh_resizex
    particle_group_list(particle_group_index).grh_resizey = grh_resizey
   
    'create particle stream
    particle_group_list(particle_group_index).particle_count = particle_count
    ReDim particle_group_list(particle_group_index).particle_stream(1 To particle_count)
   
    'plot particle group on map
    MapData(map_x, map_y).particle_group = particle_group_index
End Sub
 
Public Sub Particle_Render(ByRef temp_particle As Particle, ByVal screen_x As Integer, ByVal screen_Y As Integer, _
                            ByVal grh_index As Long, ByRef rgb_list() As Long, _
                            Optional ByVal alpha_blend As Boolean, Optional ByVal no_move As Boolean, _
                            Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal Angle As Integer, _
                            Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                            Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                            Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                            Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                            Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                            Optional ByVal bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                            Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                            Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                            Optional ByVal spin_speedH As Single, Optional ByVal Spin As Boolean, Optional grh_resize As Boolean)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 4/24/2003
'
'**************************************************************
    If no_move = False Then
                If temp_particle.alive_counter = 0 Then

                    temp_particle.index = grh_index
                    temp_particle.X = RandomNumber(x1, x2)
                    temp_particle.Y = RandomNumber(y1, y2)
                    temp_particle.vector_x = RandomNumber(vecx1, vecx2)
                    temp_particle.vector_y = RandomNumber(vecy1, vecy2)
                    temp_particle.AngleC = Angle
                    temp_particle.alive_counter = RandomNumber(life1, life2)
                    temp_particle.friction = fric
                Else
                    'Continue old particle
                    'Do gravity
                    If gravity = True Then
                        temp_particle.vector_y = temp_particle.vector_y + grav_strength
                        If temp_particle.Y > 0 Then
                            'bounce
                            temp_particle.vector_y = bounce_strength
                        End If
                    End If
                    'Do rotation
                   If Spin = True Then temp_particle.AngleC = temp_particle.AngleC + (RandomNumber(spin_speedL, spin_speedH) / 100)
                    If temp_particle.AngleC >= 360 Then
                        temp_particle.AngleC = 0
                    End If
                               
                    If XMove = True Then temp_particle.vector_x = RandomNumber(move_x1, move_x2)
                    If YMove = True Then temp_particle.vector_y = RandomNumber(move_y1, move_y2)
                End If
 
        'Add in vector
        temp_particle.X = temp_particle.X + (temp_particle.vector_x \ temp_particle.friction)
        temp_particle.Y = temp_particle.Y + (temp_particle.vector_y \ temp_particle.friction)
   
        'decrement counter
         temp_particle.alive_counter = temp_particle.alive_counter - 1
    End If

    'Draw it
    If grh_resize = True Then
        If temp_particle.index Then
              Draw_NewIndex2 temp_particle.index, temp_particle.fC, temp_particle.X + screen_x, temp_particle.Y + screen_Y, 1, 0, rgb_list(), alpha_blend, True, temp_particle.AngleC
              
        End If
    Else
    'Draw it
    If temp_particle.index Then
        Draw_NewIndex2 temp_particle.index, temp_particle.fC, temp_particle.X + screen_x, temp_particle.Y + screen_Y, 1, 0, rgb_list(), alpha_blend, True, temp_particle.AngleC
                                    
    End If
    End If
    End Sub
    Private Function OverWriteAlpha(ByVal Color As Long, ByVal Alpha As Byte) As Long
Dim Dest(3) As Byte
CopyMemory Dest(0), Color, 4

OverWriteAlpha = D3DColorARGB(Alpha, Dest(2), Dest(1), Dest(0))

End Function
    Public Function PasoTiempo(Optional ByRef Counter As Long = -1) As Long
Static Contador As Long
Dim tl As Long
tl = GetTickCount
If Counter <> -1 Then
    If Counter = 0 Then
        PasoTiempo = 0
    Else
        PasoTiempo = tl - Counter
    End If
    Counter = tl
Else
    PasoTiempo = Contador - tl
    Contador = tl
    
End If


End Function
 Private Sub Draw_NewIndex2(ByVal nIndex As Integer, ByRef fC As Single, ByVal X As Integer, ByVal Y As Integer, ByVal center As Byte, ByVal Animate As Byte, ByRef Color() As Long, Optional ByVal alpha_blend As Byte, Optional ByVal NeglectNegro As Boolean, Optional ByVal Angle As Single)

    Dim ci As Integer
    Dim jL As Integer
    Dim jT As Integer
    Dim jw As Integer
    Dim jh As Integer
    Dim jth As Integer
    Dim jtw As Integer
    Dim jg As Integer

    If NewIndexData(nIndex).Dinamica > 0 Then
        With NewAnimationData(NewIndexData(nIndex).Dinamica)
            If Animate Then
            fC = fC + ((MEE * 0.002) * .NumFrames * .Velocidad)
            ci = fC
            If ci > .NumFrames Then
                ci = ci Mod .NumFrames
            ElseIf ci <= 0 Then
                ci = 1
            End If
            Else
                If fC < 0 Then
                    fC = 1
                ElseIf fC > .NumFrames Then
                    fC = .NumFrames
                End If
            End If
            jT = .Indice(ci).Y
            jL = .Indice(ci).X
            jw = .Width
            jh = .Height
            jtw = .TileWidth
            jth = .TileWidth
            jg = NewIndexData(nIndex).OverWriteGrafico + (.Indice(ci).Grafico - .Indice(1).Grafico)
            
        End With
    Else
        With EstaticData(NewIndexData(nIndex).Estatic)
            jT = .t
            jL = .L
            jw = .W
            jh = .H
            jth = .th
            jtw = .tw
            jg = NewIndexData(nIndex).OverWriteGrafico
        
        
        End With
    End If
    
    If center Then
        If jtw <> 1 Then
            X = X - (jtw * (16)) + 16
        End If

        If jth <> 1 Then
            Y = Y - (jth * 32) + 32
        End If
    End If

    Dim RGB(3) As Long
    If alpha_blend > 0 And alpha_blend < 255 Then
        RGB(0) = OverWriteAlpha(Color(0), alpha_blend)
        RGB(1) = OverWriteAlpha(Color(1), alpha_blend)
        RGB(2) = OverWriteAlpha(Color(2), alpha_blend)
        RGB(3) = OverWriteAlpha(Color(3), alpha_blend)
    Else
        RGB(0) = Color(0)
        RGB(1) = Color(1)
        RGB(2) = Color(2)
        RGB(3) = Color(3)
    
    End If
        
    
    Static src_rect As RECT
    Static dest_rect As RECT
    Static temp_verts(3) As TLVERTEX
    Static d3dTextures As D3D8Textures

        
    Set d3dTextures.Texture = DXPool.GetTexture(jg)
    DXPool.Texture_Dimension_Get jg, d3dTextures.texwidth, d3dTextures.texheight

        
    With src_rect
        .Bottom = jT + jh - 1
        .left = jL
        .Right = jL + jw - 1
        .top = jT
    End With

    With dest_rect
        .Bottom = (Y + jh)
        .left = X
        .Right = (X + jw)
        .top = Y
    End With


    Geometry_Create_Box temp_verts(), dest_rect, src_rect, RGB(), d3dTextures.texwidth - 1, d3dTextures.texheight - 1, Angle

    ddevice.SetTexture 0, d3dTextures.Texture

    If alpha_blend > 0 Or NeglectNegro Then
        'Set Rendering for alphablending

        If Not NeglectNegro Then
            ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
            ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA


        Else
            ddevice.SetRenderState D3DRS_DESTBLEND, 2
                        ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        'D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        End If

    End If
    
    'Draw the triangles that make up our square Textures

            ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0))

    
    If alpha_blend > 0 Or NeglectNegro Then
        'Set Rendering for colokeying
        ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        Call ddevice.SetTextureStageState(0, D3DTSS_ALPHAARG1, D3DTOP_SELECTARG1)
        Call ddevice.SetTextureStageState(0, D3DTSS_ALPHAARG2, D3DTOP_DISABLE)

        
    End If




End Sub
Public Sub Particle_Group_Render(ByVal particle_group_index As Long, ByVal screen_x As Integer, ByVal screen_Y As Integer)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 12/15/2002
'Renders a particle stream at a paticular screen point
'*****************************************************************
    Dim LoopC As Long
    Dim temp_rgb(0 To 3) As Long
    Dim no_move As Boolean
   
    'Set colors
'    If UserMinHP = 0 Then
'        temp_rgb(0) = D3DColorARGB(particle_group_list(particle_group_index).alpha_blend, 255, 255, 255)
'        temp_rgb(1) = D3DColorARGB(particle_group_list(particle_group_index).alpha_blend, 255, 255, 255)
'        temp_rgb(2) = D3DColorARGB(particle_group_list(particle_group_index).alpha_blend, 255, 255, 255)
'        temp_rgb(3) = D3DColorARGB(particle_group_list(particle_group_index).alpha_blend, 255, 255, 255)
'    Else
        temp_rgb(0) = particle_group_list(particle_group_index).rgb_list(0)
        temp_rgb(1) = particle_group_list(particle_group_index).rgb_list(1)
        temp_rgb(2) = particle_group_list(particle_group_index).rgb_list(2)
        temp_rgb(3) = particle_group_list(particle_group_index).rgb_list(3)
'    End If
'
    If particle_group_list(particle_group_index).alive_counter Then
   
        'See if it is time to move a particle
        particle_group_list(particle_group_index).frame_counter = particle_group_list(particle_group_index).frame_counter + timer_ticks_per_frame
        If particle_group_list(particle_group_index).frame_counter > particle_group_list(particle_group_index).frame_speed Then
            particle_group_list(particle_group_index).frame_counter = 0
            no_move = False
        Else
            no_move = True
        End If
   
   
        'If it's still alive render all the particles inside
        For LoopC = 1 To particle_group_list(particle_group_index).particle_count
       
            'Render particle
            Particle_Render particle_group_list(particle_group_index).particle_stream(LoopC), _
                            screen_x, screen_Y, _
                            particle_group_list(particle_group_index).grh_index_list(Round(RandomNumber(1, particle_group_list(particle_group_index).grh_index_count), 0)), _
                            temp_rgb(), _
                            particle_group_list(particle_group_index).alpha_blend, no_move, _
                            particle_group_list(particle_group_index).x1, particle_group_list(particle_group_index).y1, particle_group_list(particle_group_index).Angle, _
                            particle_group_list(particle_group_index).vecx1, particle_group_list(particle_group_index).vecx2, _
                            particle_group_list(particle_group_index).vecy1, particle_group_list(particle_group_index).vecy2, _
                            particle_group_list(particle_group_index).life1, particle_group_list(particle_group_index).life2, _
                            particle_group_list(particle_group_index).fric, particle_group_list(particle_group_index).spin_speedL, _
                            particle_group_list(particle_group_index).gravity, particle_group_list(particle_group_index).grav_strength, _
                            particle_group_list(particle_group_index).bounce_strength, particle_group_list(particle_group_index).x2, _
                            particle_group_list(particle_group_index).y2, particle_group_list(particle_group_index).XMove, _
                            particle_group_list(particle_group_index).move_x1, particle_group_list(particle_group_index).move_x2, _
                            particle_group_list(particle_group_index).move_y1, particle_group_list(particle_group_index).move_y2, _
                            particle_group_list(particle_group_index).YMove, particle_group_list(particle_group_index).spin_speedH, _
                            particle_group_list(particle_group_index).Spin, particle_group_list(particle_group_index).grh_resize
                           
        Next LoopC
       
        If no_move = False Then
            'Update the group alive counter
            If particle_group_list(particle_group_index).never_die = False Then
                particle_group_list(particle_group_index).alive_counter = particle_group_list(particle_group_index).alive_counter - 1
            End If
        End If
   
    Else
        'If it's dead destroy it
        particle_group_list(particle_group_index).particle_count = particle_group_list(particle_group_index).particle_count - 1
        If particle_group_list(particle_group_index).particle_count <= 0 Then Particle_Group_Destroy particle_group_index
    End If
End Sub
 Public Function RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
      'Initialize randomizer
      Randomize timer
    
      'Generate random number
      RandomNumber = (UpperBound - LowerBound) * Rnd + LowerBound
End Function
Public Function Particle_Type_Get(ByVal particle_index As Long) As Long
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero ([email=juansotuyo@hotmail.com]juansotuyo@hotmail.com[/email])
'Last Modify Date: 8/27/2003
'Returns the stream type of a particle stream
'*****************************************************************
    If Particle_Group_Check(particle_index) Then
        Particle_Type_Get = particle_group_list(particle_index).stream_type
    End If
End Function
 
Public Function Particle_Group_Check(ByVal particle_group_index As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'**************************************************************
    'check index
    If particle_group_index > 0 And particle_group_index <= particle_group_last Then
        If particle_group_list(particle_group_index).Active Then
            Particle_Group_Check = True
        End If
    End If
End Function
 
Public Function Particle_Group_Map_Pos_Set(ByVal particle_group_index As Long, ByVal map_x As Long, ByVal map_y As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/27/2003
'Returns true if successful, else false
'**************************************************************
    'Make sure it's a legal index
    If Particle_Group_Check(particle_group_index) Then
        'Make sure it's a legal move
        If InMapBounds(map_x, map_y) Then
            'Move it
            particle_group_list(particle_group_index).map_x = map_x
            particle_group_list(particle_group_index).map_y = map_y
   
            Particle_Group_Map_Pos_Set = True
        End If
    End If
End Function
 
Public Function Particle_Group_Move(ByVal particle_group_index As Long, ByVal Heading As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/27/2003
'Returns true if successful, else false
'**************************************************************
    Dim map_x As Long
    Dim map_y As Long
    Dim nX As Integer
    Dim nY As Integer
   
    'Check for valid heading
    If Heading < 1 Or Heading > 8 Then
        Particle_Group_Move = False
        Exit Function
    End If
   
    'Make sure it's a legal index
    If Particle_Group_Check(particle_group_index) Then
   
        map_x = particle_group_list(particle_group_index).map_x
        map_y = particle_group_list(particle_group_index).map_y
       
        nX = map_x
        nY = map_y
       
        Convert_Heading_to_Direction Heading, nX, nY
       
        'Make sure it's a legal move
        If InMapBounds(nX, nY) Then
            'Move it
            particle_group_list(particle_group_index).map_x = nX
            particle_group_list(particle_group_index).map_y = nY
           
            Particle_Group_Move = True
        End If
    End If
End Function
 
Public Sub Particle_Group_Destroy(ByVal particle_group_index As Long)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim temp As particle_group
    Dim i As Integer
   
    If particle_group_list(particle_group_index).map_x > 0 And particle_group_list(particle_group_index).map_y > 0 Then
        MapData(particle_group_list(particle_group_index).map_x, particle_group_list(particle_group_index).map_y).particle_group = 0
    End If
   
    particle_group_list(particle_group_index) = temp
           
    'Update array size
    If particle_group_index = particle_group_last Then
        Do Until particle_group_list(particle_group_last).Active
            particle_group_last = particle_group_last - 1
            If particle_group_last = 0 Then
                particle_group_count = 0
                Exit Sub
            End If
        Loop
        ReDim Preserve particle_group_list(1 To particle_group_last)
    End If
    particle_group_count = particle_group_count - 1
End Sub
 
Public Sub Char_Particle_Group_Make(ByVal particle_group_index As Long, ByVal char_index As Integer, ByVal particle_char_index As Integer, _
                                ByVal particle_count As Long, ByVal stream_type As Long, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
                                Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal Angle As Integer, _
                                Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                Optional bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                                Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                Optional ByVal spin_speedH As Single, Optional ByVal Spin As Boolean, Optional grh_resize As Boolean, _
                                Optional grh_resizex As Integer, Optional grh_resizey As Integer)
                               
'*****************************************************************
'Author: Aaron Perkins
'Modified by: Ryan Cain (Onezero)
'Last Modify Date: 5/15/2003
'Makes a new particle effect
'Modified by Juan Martín Sotuyo Dodero
'*****************************************************************
    'Update array size
    If particle_group_index > particle_group_last Then
        particle_group_last = particle_group_index
        ReDim Preserve particle_group_list(1 To particle_group_last)
    End If
    particle_group_count = particle_group_count + 1
   
    'Make active
    particle_group_list(particle_group_index).Active = True
   
    'Char index
    particle_group_list(particle_group_index).char_index = char_index
   
    'Grh list
    ReDim particle_group_list(particle_group_index).grh_index_list(1 To UBound(grh_index_list))
    particle_group_list(particle_group_index).grh_index_list() = grh_index_list()
    particle_group_list(particle_group_index).grh_index_count = UBound(grh_index_list)
   
    'Sets alive vars
    If alive_counter = -1 Then
        particle_group_list(particle_group_index).alive_counter = -1
        particle_group_list(particle_group_index).never_die = True
    Else
        particle_group_list(particle_group_index).alive_counter = alive_counter
        particle_group_list(particle_group_index).never_die = False
    End If
   
    'alpha blending
    particle_group_list(particle_group_index).alpha_blend = alpha_blend
   
    'stream type
    particle_group_list(particle_group_index).stream_type = stream_type
   
    'speed
    particle_group_list(particle_group_index).frame_speed = frame_speed
   
    particle_group_list(particle_group_index).x1 = x1
    particle_group_list(particle_group_index).y1 = y1
    particle_group_list(particle_group_index).x2 = x2
    particle_group_list(particle_group_index).y2 = y2
    particle_group_list(particle_group_index).Angle = Angle
    particle_group_list(particle_group_index).vecx1 = vecx1
    particle_group_list(particle_group_index).vecx2 = vecx2
    particle_group_list(particle_group_index).vecy1 = vecy1
    particle_group_list(particle_group_index).vecy2 = vecy2
    particle_group_list(particle_group_index).life1 = life1
    particle_group_list(particle_group_index).life2 = life2
    particle_group_list(particle_group_index).fric = fric
    particle_group_list(particle_group_index).Spin = Spin
    particle_group_list(particle_group_index).spin_speedL = spin_speedL
    particle_group_list(particle_group_index).spin_speedH = spin_speedH
    particle_group_list(particle_group_index).gravity = gravity
    particle_group_list(particle_group_index).grav_strength = grav_strength
    particle_group_list(particle_group_index).bounce_strength = bounce_strength
    particle_group_list(particle_group_index).XMove = XMove
    particle_group_list(particle_group_index).YMove = YMove
    particle_group_list(particle_group_index).move_x1 = move_x1
    particle_group_list(particle_group_index).move_x2 = move_x2
    particle_group_list(particle_group_index).move_y1 = move_y1
    particle_group_list(particle_group_index).move_y2 = move_y2
   
    'color
    particle_group_list(particle_group_index).rgb_list(0) = rgb_list(0)
    particle_group_list(particle_group_index).rgb_list(1) = rgb_list(1)
    particle_group_list(particle_group_index).rgb_list(2) = rgb_list(2)
    particle_group_list(particle_group_index).rgb_list(3) = rgb_list(3)
   
    particle_group_list(particle_group_index).grh_resize = grh_resize
    particle_group_list(particle_group_index).grh_resizex = grh_resizex
    particle_group_list(particle_group_index).grh_resizey = grh_resizey
 
    'handle
    particle_group_list(particle_group_index).id = id
   
    'create particle stream
    particle_group_list(particle_group_index).particle_count = particle_count
    ReDim particle_group_list(particle_group_index).particle_stream(1 To particle_count)
   
    'plot particle group on char
    CharList(char_index).particle_group(particle_char_index) = particle_group_index
End Sub
 
Public Function Char_Particle_Group_Create(ByVal char_index As Integer, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                        Optional ByVal particle_count As Long = 20, Optional ByVal stream_type As Long = 1, _
                                        Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                        Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
                                        Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal Angle As Integer, _
                                        Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                        Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                        Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                        Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                        Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                        Optional bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                                        Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                        Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                        Optional ByVal spin_speedH As Single, Optional ByVal Spin As Boolean, Optional grh_resize As Boolean, _
                                        Optional grh_resizex As Integer, Optional grh_resizey As Integer) As Long
'**************************************************************
'Author: Augusto José Rando
'**************************************************************
    Dim char_part_free_index As Integer
   
    'If Char_Particle_Group_Find(char_index, stream_type) Then Exit Function ' hay que ver si dejar o sacar esto...
    If Not Char_Check(char_index) Then Exit Function
    char_part_free_index = Char_Particle_Group_Next_Open(char_index)
   
    If char_part_free_index > 0 Then
        Char_Particle_Group_Create = Particle_Group_Next_Open
        Char_Particle_Group_Make Char_Particle_Group_Create, char_index, char_part_free_index, particle_count, stream_type, grh_index_list(), rgb_list(), alpha_blend, alive_counter, frame_speed, id, x1, y1, Angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, x2, y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, Spin, grh_resize, grh_resizex, grh_resizey
       
    End If
 
End Function
 
Public Function Char_Particle_Group_Find(ByVal char_index As Integer, ByVal stream_type As Long) As Integer
'*****************************************************************
'Author: Augusto José Rando
'Modified: returns slot or -1
'*****************************************************************
 
Dim i As Integer
 
For i = 1 To CharList(char_index).particle_count
    If particle_group_list(CharList(char_index).particle_group(i)).stream_type = stream_type Then
        Char_Particle_Group_Find = CharList(char_index).particle_group(i)
        Exit Function
    End If
Next i
 
Char_Particle_Group_Find = -1
 
End Function
 
Public Function Char_Particle_Group_Next_Open(ByVal char_index As Integer) As Integer
'*****************************************************************
'Author: Augusto José Rando
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim LoopC As Long
   
    LoopC = 1
    Do Until CharList(char_index).particle_group(LoopC) = 0
        If LoopC = CharList(char_index).particle_count Then
            Char_Particle_Group_Next_Open = CharList(char_index).particle_count + 1
            CharList(char_index).particle_count = Char_Particle_Group_Next_Open
            ReDim Preserve CharList(char_index).particle_group(1 To Char_Particle_Group_Next_Open) As Long
            Exit Function
        End If
        LoopC = LoopC + 1
    Loop
   
    Char_Particle_Group_Next_Open = LoopC
 
Exit Function
 
ErrorHandler:
    CharList(char_index).particle_count = 1
    ReDim CharList(char_index).particle_group(1 To 1) As Long
    Char_Particle_Group_Next_Open = 1
 
End Function
 
Public Function Char_Particle_Group_Remove(ByVal char_index As Integer, ByVal stream_type As Long)
'**************************************************************
'Author: Augusto José Rando
'**************************************************************
    Dim char_part_index As Integer
   
    If Char_Check(char_index) Then
        char_part_index = Char_Particle_Group_Find(char_index, stream_type)
        If char_part_index = -1 Then Exit Function
        Call Particle_Group_Remove(char_part_index)
    End If
 
End Function
 
Public Function Char_Particle_Group_Remove_All(ByVal char_index As Integer)
'**************************************************************
'Author: Augusto José Rando
'**************************************************************
    Dim i As Integer
   
    If Char_Check(char_index) Then
        For i = 1 To CharList(char_index).particle_count
            If CharList(char_index).particle_group(i) <> 0 Then Call Particle_Group_Remove(CharList(char_index).particle_group(i))
        Next i
    End If
   
End Function
 
Public Function Map_Particle_Group_Get(ByVal map_x As Integer, ByVal map_y As Integer) As Long
 
    If InMapBounds(map_x, map_y) Then
        Map_Particle_Group_Get = MapData(map_x, map_y).particle_group
    Else
        Map_Particle_Group_Get = 0
    End If
End Function
 
Public Sub Grh_Render_Advance(ByRef Grh As Grh, ByVal screen_x As Integer, ByVal screen_Y As Integer, ByRef rgb_list() As Long, Optional ByVal alpha_blend As Boolean = False)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero ([email=juansotuyo@hotmail.com]juansotuyo@hotmail.com[/email])
'Last Modify Date: 11/19/2003
'Similar to Grh_Render, but let´s you resize the Grh
'**************************************************************
    Dim grh_index As Long
   
    'Animation
    If Grh.Started Then
        Grh.frame_counter = Grh.frame_counter + (timer_ticks_per_frame * Grh.frame_speed)
        If Grh.frame_counter > grh_list(Grh.grh_index).frame_count Then
            'If Grh.noloop Then
            '    Grh.FrameCounter = GrhData(Grh.GrhIndex).NumFrames
            'Else
                Grh.frame_counter = 1
            'End If
        End If
    End If
   
    'Figure out what frame to draw (always 1 if not animated)
    If Grh.frame_counter = 0 Then Grh.frame_counter = 1
    grh_index = grh_list(Grh.grh_index).frame_list(Grh.frame_counter)
   
    'Center Grh over X, Y pos
    If grh_list(Grh.grh_index).src_width <> 1 Then
        screen_x = screen_x - Int(grh_list(Grh.grh_index).src_width * (base_tile_size \ 2)) + base_tile_size \ 2
    End If
   
    If grh_list(Grh.grh_index).src_height <> 1 Then
        screen_Y = screen_Y - Int(grh_list(Grh.grh_index).src_height * base_tile_size) + base_tile_size
    End If
   
    'Draw it to device
    modDXEngine.DXEngine_TextureRender grh_list(grh_index).texture_index, _
        screen_x, screen_Y, _
        grh_list(grh_index).src_width, grh_list(grh_index).src_height, _
        rgb_list, _
        grh_list(grh_index).Src_X, grh_list(grh_index).Src_Y, _
        grh_list(grh_index).src_width, grh_list(grh_index).src_height, alpha_blend, Grh.Angle
End Sub
 
Public Function Char_Check(ByVal char_index As Integer) As Boolean
    'check char_index
    If char_index > 0 And char_index <= LastChar Then
        Char_Check = (CharList(char_index).Heading > 0)
    End If
   
End Function
 
Public Sub Grh_Render(ByRef Grh As Grh, ByVal screen_x As Integer, ByVal screen_Y As Integer, ByRef rgb_list() As Long, Optional ByVal h_centered As Boolean = True, Optional ByVal v_centered As Boolean = True, Optional ByVal alpha_blend As Boolean = False)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/28/2003
'Modified by Juan Martín Sotuyo Dodero
'Added centering
'**************************************************************
On Error Resume Next
 
    Dim grh_index As Long
   
    If Grh.grh_index = 0 Then Exit Sub
       
    'Animation
    If Grh.Started = 1 Then
        Grh.frame_counter = Grh.frame_counter + (timer_elapsed_time * grh_list(Grh.grh_index).frame_count / Grh.frame_speed)
        If Grh.frame_counter > grh_list(Grh.grh_index).frame_count Then
            Grh.frame_counter = (Grh.frame_counter Mod grh_list(Grh.grh_index).frame_count) + 1
            If Grh.LoopTimes <> -1 Then
                If Grh.LoopTimes > 0 Then
                    Grh.LoopTimes = Grh.LoopTimes - 1
                Else
                    Grh.Started = 0
                End If
            End If
        End If
    End If
 
    'Figure out what frame to draw (always 1 if not animated)
    If Grh.frame_counter = 0 Then Grh.frame_counter = 1
    'If Not Grh_Check(Grh.grhindex) Then Exit Sub
    grh_index = grh_list(Grh.grh_index).frame_list(Grh.frame_counter)
    If grh_index <= 0 Then Exit Sub
    If grh_list(grh_index).texture_index = 0 Then Exit Sub
       
    'Modified by Augusto José Rando
    'Simplier function - according to basic ORE engine
    If h_centered Then
        If grh_list(Grh.grh_index).src_width <> 1 Then
            screen_x = screen_x - Int(grh_list(Grh.grh_index).src_width * (32 \ 2)) + 32 \ 2
        End If
    End If
   
    If v_centered Then
        If grh_list(Grh.grh_index).src_height <> 1 Then
            screen_Y = screen_Y - Int(grh_list(Grh.grh_index).src_height * 32) + 32
        End If
    End If
   
    'Draw it to device
    modDXEngine.DXEngine_TextureRender grh_list(grh_index).texture_index, _
        screen_x, screen_Y, _
        grh_list(grh_index).src_width, grh_list(grh_index).src_height, _
        rgb_list(), _
        grh_list(grh_index).Src_X, grh_list(grh_index).Src_Y, grh_list(grh_index).src_width, grh_list(grh_index).src_height _
        , alpha_blend
 
End Sub

'Initialization
Public Function DXEngine_Initialize(ByVal f_hwnd As Long, ByVal s_hwnd As Long, ByVal windowed As Boolean)
'On Error GoTo errhandler
    Dim d3dcaps As D3DCAPS8
    Dim d3ddm As D3DDISPLAYMODE
    
    DXEngine_Initialize = True
    
    'Main display
    screen_hwnd = s_hwnd
    form_hwnd = f_hwnd
    
    '*******************************
    'Initialize root DirectX8 objects
    '*******************************
    Set dx = New DirectX8
    'Create the Direct3D object
    Set d3d = dx.Direct3DCreate
    'Create helper class
    Set d3dx = New D3DX8
    
    '*******************************
    'Initialize video device
    '*******************************
    Dim DevType As CONST_D3DDEVTYPE
    DevType = D3DDEVTYPE_HAL
    'Get the capabilities of the Direct3D device that we specify. In this case,
    'we'll be using the adapter default (the primiary card on the system).
    Call d3d.GetDeviceCaps(D3DADAPTER_DEFAULT, DevType, d3dcaps)
    'Grab some information about the current display mode.
    Call d3d.GetAdapterDisplayMode(D3DADAPTER_DEFAULT, d3ddm)
    
    'Now we'll go ahead and fill the D3DPRESENT_PARAMETERS type.
    With d3dpp
        .windowed = 1
        .SwapEffect = D3DSWAPEFFECT_COPY
        .BackBufferFormat = d3ddm.Format 'current display depth
        .BackBufferHeight = 600
        .BackBufferWidth = 800
        .EnableAutoDepthStencil = 1
        .AutoDepthStencilFormat = D3DFMT_D16
        .hDeviceWindow = frmMain.hWnd
    End With
    'create device
    Set ddevice = d3d.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, screen_hwnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, d3dpp)
    
    SPOTLIGHTS_LOADDAT

    DeviceRenderStates
    
    '****************************************************
    'Inicializamos el manager de texturas
    '****************************************************
    Call DXPool.Texture_Initialize(500)
    
    '****************************************************
    'Clears the buffer to start rendering
    '****************************************************
    Device_Clear
    '****************************************************
    'Load Misc
    '****************************************************
    LoadGraphicFonts
    LoadFonts
    
    Engine_Init_FontSettings
    Engine_Init_FontTextures
    
    

      Font_Make "Verdana", 8, False, False

    'CargarParticulas
    Exit Function
ErrHandler:
    DXEngine_Initialize = False
End Function

Public Function DXEngine_BeginRender() As Boolean
On Error GoTo ErrorHandler:
    DXEngine_BeginRender = True
    
    'Check if we have the device
    If ddevice.TestCooperativeLevel <> D3D_OK Then
        Do
            DoEvents
        Loop While ddevice.TestCooperativeLevel = D3DERR_DEVICELOST
        
        DXPool.Texture_Remove_All
        Fonts_Destroy
        Device_Reset
        
        DeviceRenderStates
        LoadFonts
        LoadGraphicFonts
    End If
    
    '****************************************************
    'Render
    '****************************************************
    '*******************************
    'Erase the backbuffer so that it can be drawn on again
    Device_Clear
    '*******************************
    '*******************************
    'Start the scene
    ddevice.BeginScene
    '*******************************
    
    engine_render_started = True
Exit Function
ErrorHandler:
    DXEngine_BeginRender = False
    MsgBox "Error in Engine_Render_Start: " & Err.Number & ": " & Err.Description
End Function

Public Function DXEngine_EndRender() As Boolean
On Error GoTo ErrorHandler:
    DXEngine_EndRender = True

    If engine_render_started = False Then
        Exit Function
    End If
    
    '*******************************
    'End scene
    ddevice.EndScene
    '*******************************
    
    '*******************************
    'Flip the backbuffer to the screen
    Device_Flip
    '*******************************
    
    '*******************************
    'Calculate current frames per second
    If GetTickCount >= (fps_last_time + 1000) Then
        FPS = fps_frame_counter
        fps_frame_counter = 0
        fps_last_time = GetTickCount
    Else
        fps_frame_counter = fps_frame_counter + 1
    End If
    '*******************************
    

    
    
    engine_render_started = False
Exit Function
ErrorHandler:
    DXEngine_EndRender = False
    MsgBox "Error in Engine_Render_End: " & Err.Number & ": " & Err.Description
  End Function

Private Sub Device_Clear()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    'Clear the back buffer
    ddevice.Clear 0, ByVal 0&, D3DCLEAR_TARGET, 0, 1#, 0
End Sub

Private Function Device_Reset() As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Resets the device
'**************************************************************
On Error GoTo ErrHandler:
'On Error Resume Next

    'Be sure the scene is finished
    ddevice.EndScene
    'Reset device
    ddevice.Reset d3dpp
    
    DeviceRenderStates
       
Exit Function
ErrHandler:
    Device_Reset = Err.Number
End Function
Public Sub DXEngine_TextureRenderAdvance(ByVal texture_index As Long, ByVal dest_x As Long, ByVal dest_y As Long, ByVal Src_X As Long, ByVal Src_Y As Long, _
                                             ByVal dest_width As Long, ByVal dest_height As Long, ByVal src_width As Long, ByVal src_height As Long, ByRef rgb_list() As Long, _
                                            Optional ByVal alpha_blend As Boolean, Optional ByVal Angle As Single)
'**************************************************************
'This sub allow texture resizing
'
'**************************************************************

    
    Dim src_rect As RECT
    Dim dest_rect As RECT
    Dim temp_verts(3) As TLVERTEX
    Dim Texture As Direct3DTexture8
    Dim texture_width As Integer
    Dim texture_height As Integer
    Dim R(3) As Long
    'rgb_list(0) = RGB(255, 255, 255)
    'rgb_list(1) = RGB(255, 255, 255)
    'rgb_list(2) = RGB(255, 255, 255)
    'rgb_list(3) = RGB(255, 255, 255)
    
    'Set up the destination rectangle
    With dest_rect
        .Bottom = dest_y + dest_height
        .left = dest_x
        .Right = dest_x + dest_width
        .top = dest_y
    End With
    
    With src_rect
        .Bottom = Src_Y + src_height - 1
        .Right = Src_X + src_width - 1
        .top = Src_Y
        .left = Src_X
    End With
    
    If rgb_list(0) = 0 Then R(0) = base_light
    
    If rgb_list(1) = 0 Then R(1) = base_light
    If rgb_list(2) = 0 Then R(2) = base_light
    If rgb_list(3) = 0 Then R(3) = base_light
    
    Set Texture = DXPool.GetTexture(texture_index)
    Call DXPool.Texture_Dimension_Get(texture_index, texture_width, texture_height)
    
    'Set up the TempVerts(3) vertices
    Geometry_Create_Box temp_verts(), dest_rect, src_rect, R(), texture_width, texture_height, Angle
    
    'Set Texture
    ddevice.SetTexture 0, Texture
    
    If alpha_blend Then
       'Set Rendering for alphablending
        ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
        ddevice.SetRenderState D3DRS_COLORVERTEX, 1
    End If
    
    'Draw the triangles that make up our square texture
    ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0))
    
    If alpha_blend Then
        'Set Rendering for colokeying
        ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    End If
End Sub
Public Sub setup_ambient()

'Noche 87, 61, 43
luz_dia(0).R = 45
luz_dia(0).G = 55
luz_dia(0).B = 70

luz_dia(1).R = 45
luz_dia(1).G = 55
luz_dia(1).B = 70

luz_dia(2).R = 60
luz_dia(2).G = 70
luz_dia(2).B = 70

luz_dia(3).R = 90
luz_dia(3).G = 90
luz_dia(3).B = 90

'4 am 124,117,91
luz_dia(4).R = 110
luz_dia(4).G = 110
luz_dia(4).B = 90

'5,6 am 143,137,135
luz_dia(5).R = 130
luz_dia(5).G = 130
luz_dia(5).B = 100

luz_dia(6).R = 145
luz_dia(6).G = 145
luz_dia(6).B = 120

'7 am 212,205,207
luz_dia(7).R = 155
luz_dia(7).G = 155
luz_dia(7).B = 155

luz_dia(8).R = 165
luz_dia(8).G = 155
luz_dia(8).B = 160

luz_dia(9).R = 180
luz_dia(9).G = 175
luz_dia(9).B = 180

luz_dia(10).R = 195
luz_dia(10).G = 190
luz_dia(10).B = 195

luz_dia(11).R = 215
luz_dia(11).G = 215
luz_dia(11).B = 215


luz_dia(12).R = 230
luz_dia(12).G = 230
luz_dia(12).B = 230

luz_dia(13).R = 230
luz_dia(13).G = 230
luz_dia(13).B = 230

'Medio Dia 255, 200, 255
luz_dia(14).R = 230
luz_dia(14).G = 220
luz_dia(14).B = 230

luz_dia(15).R = 220
luz_dia(15).G = 220
luz_dia(15).B = 220

luz_dia(16).R = 210
luz_dia(16).G = 210
luz_dia(16).B = 210

'17/18 0, 100, 255
luz_dia(17).R = 180
luz_dia(17).G = 170
luz_dia(17).B = 140
'18/19 0, 100, 255

luz_dia(18).R = 160
luz_dia(18).G = 150
luz_dia(18).B = 90

'19/20 156, 142, 83
luz_dia(19).R = 130
luz_dia(19).G = 100
luz_dia(19).B = 80

luz_dia(20).R = 95
luz_dia(20).G = 95
luz_dia(20).B = 80

luz_dia(21).R = 80
luz_dia(21).G = 80
luz_dia(21).B = 80

luz_dia(22).R = 60
luz_dia(22).G = 65
luz_dia(22).B = 80

luz_dia(23).R = 50
luz_dia(23).G = 60
luz_dia(23).B = 75

luz_dia(24).R = 45
luz_dia(24).G = 55
luz_dia(24).B = 70
            
            
Dim t As Integer
Dim B As Byte
Dim X As Byte
ReDim ambient_light(1 To 240) As Long
Dim xr As Byte
Dim xg As Byte
Dim xb As Byte


HoraLuz = 14
extra_light(eE_Light.Oscuridad) = D3DColorXRGB(10, 10, 10)
extra_light(eE_Light.Cegador) = D3DColorXRGB(255, 255, 255)
extra_light(eE_Light.Azul1) = D3DColorXRGB(0, 0, 255)
extra_light(eE_Light.Azul2) = D3DColorXRGB(100, 100, 155)
extra_light(eE_Light.Azul3) = D3DColorXRGB(50, 50, 200)
extra_light(eE_Light.Rojo1) = D3DColorXRGB(200, 0, 0)
extra_light(eE_Light.Rojo2) = D3DColorXRGB(155, 100, 100)
extra_light(eE_Light.Rojo3) = D3DColorXRGB(200, 50, 50)
extra_light(eE_Light.Verde1) = D3DColorXRGB(0, 200, 0)
extra_light(eE_Light.Verde2) = D3DColorXRGB(100, 155, 100)
extra_light(eE_Light.Verde3) = D3DColorXRGB(50, 200, 50)

For t = 1 To 225

    X = X + 1

    If X > 5 Then 'aclaro
        If luz_dia(B).R + 60 > 255 Or luz_dia(B).G + 60 > 255 Or luz_dia(B).B + 60 > 255 Then
        'Dividimo el resto.
            xr = Int((255 - luz_dia(B).R) / 4)
            xb = Int((255 - luz_dia(B).B) / 4)
            xg = Int((255 - luz_dia(B).G) / 4)
            ambient_light(t) = D3DColorXRGB(luz_dia(B).R + ((X - 5) * xr), luz_dia(B).G + ((X - 5) * xg), luz_dia(B).B + ((X - 5) * xb))
        
        Else
            ambient_light(t) = D3DColorXRGB(luz_dia(B).R + ((X - 5) * 15), luz_dia(B).G + ((X - 5) * 15), luz_dia(B).B + ((X - 5) * 15))
        End If
    
    
    ElseIf X > 1 Then
        If luz_dia(B).R - 40 < 15 Or luz_dia(B).G - 40 < 15 - luz_dia(B).B + 40 < 15 Then
        'Dividimo el resto.
            xr = Int((luz_dia(B).R - 15) / 4)
            xb = Int((luz_dia(B).B - 15) / 4)
            xg = Int((luz_dia(B).G - 15) / 4)
    ambient_light(t) = D3DColorXRGB(luz_dia(B).R - ((X - 1) * xr), luz_dia(B).G - ((X - 1) * xg), luz_dia(B).B - ((X - 1) * xb))
        
        Else
    ambient_light(t) = D3DColorXRGB(luz_dia(B).R - ((X - 1) * 10), luz_dia(B).G - ((X - 1) * 10), luz_dia(B).B - ((X - 1) * 10))
        End If
        
    
    Else
        ambient_light(t) = D3DColorXRGB(luz_dia(B).R, luz_dia(B).G, luz_dia(B).B)
    End If
    If X = 9 Then
        B = B + 1 'cambia de hora
        X = 0
    End If
Next t


base_light = ambient_light(14 * 9 + 1)
            
End Sub
Public Sub DXEngine_TextureRender(ByVal texture_index As Long, ByVal dest_x As Long, ByVal dest_y As Long, ByVal src_width As Long, _
                                            ByVal src_height As Long, ByRef rgb_list() As Long, ByVal Src_X As Long, _
                                            ByVal Src_Y As Long, ByVal dest_width As Long, ByVal dest_height As Long, _
                                            Optional ByVal alpha_blend As Boolean, Optional ByVal Angle As Single)
'**************************************************************
'This sub doesnt allow texture resizing
'
'**************************************************************
    Dim src_rect As RECT
    Dim dest_rect As RECT
    Dim temp_verts(3) As TLVERTEX
    Dim texture_height As Integer
    Dim texture_width As Integer
    Dim Texture As Direct3DTexture8
    Dim R(3) As Long
    'Set up the source rectangle
    With src_rect
        .Bottom = Src_Y + src_height - 1
        .left = Src_X
        .Right = Src_X + src_width - 1
        .top = Src_Y
    End With
        
    'Set up the destination rectangle
    With dest_rect
        .Bottom = dest_y + dest_height
        .left = dest_x
        .Right = dest_x + dest_width
        .top = dest_y
    End With
    If rgb_list(0) = 0 Then R(0) = base_light Else R(0) = rgb_list(0)
    If rgb_list(1) = 0 Then R(1) = base_light Else R(1) = rgb_list(1)
    If rgb_list(2) = 0 Then R(2) = base_light Else R(2) = rgb_list(2)
    If rgb_list(3) = 0 Then R(3) = base_light Else R(3) = rgb_list(3)
    
    'ESTO NO ME GUSTA
    Set Texture = DXPool.GetTexture(texture_index)
    Call DXPool.Texture_Dimension_Get(texture_index, texture_width, texture_height)
    'Set up the TempVerts(3) vertices
    Geometry_Create_Box temp_verts(), dest_rect, src_rect, R(), texture_height, texture_width, Angle
    'Set Texture
    ddevice.SetTexture 0, Texture
    
    'Enable alpha-blending
    alpha_blend = False
    ddevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
'Call ddevice.SetTextureStageState(0, D3DTSS_ALPHAOP, D3DTOP_MODULATE)
'Call ddevice.SetTextureStageState(0, D3DTSS_ALPHAARG1, D3DTA_DIFFUSE Or D3DTA_TEXTURE)
'Call ddevice.SetTextureStageState(0, D3DTSS_ALPHAARG2, D3DTA_CURRENT)
'Call ddevice.SetTextureStageState(0, D3DTSS_COLORARG1, D3DTA_TEXTURE)
'Call ddevice.SetTextureStageState(0, D3DTSS_COLORARG2, D3DTA_TFACTOR)
'    Call ddevice.SetRenderState(D3DRS_TEXTUREFACTOR, D3DColorARGB(255, 0, 0, 0))
'Call ddevice.SetTextureStageState(0, D3DTSS_COLOROP, 18)
'CON EL 18  D3DTOP_MODULATEALPHA_ADDCOLOR hace las dos cosas, agrega alpha y da color. la intesidad del color se da por el valor
'no logro aplicar el GRADO DE ALPHA a las cosas.

    If alpha_blend Then
       'Set Rendering for alphablending
ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
        'ddevice.SetRenderState D3DRS_COLORVERTEX, 2
    End If
    
    'Draw the triangles that make up our square texture
    ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0))
    
    If alpha_blend Then
        'Set Rendering for colokeying
        ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    End If
    
    'Turn off alphablending after we're done
    'ddevice.SetRenderState D3DRS_ALPHABLENDENABLE, 0
End Sub
Public Sub Recalcular_LUZ(ByVal Luz As Byte)
If Luz > 24 Then Exit Sub

Dim X As Long
Dim Y As Long
For X = 1 To 100
For Y = 1 To 100


                                    If MapData(X, Y).Luz > 0 And MapData(X, Y).Luz <= 100 Then
                        If MapData(X, Y).Luz > 0 And MapData(X, Y).Luz < 9 Then 'Luces normales
                            MapData(X, Y).light_value(0) = ambient_light((Luz * 9) + MapData(X, Y).Luz + 1)
                            MapData(X, Y).light_value(1) = ambient_light((Luz * 9) + MapData(X, Y).Luz + 1)
                            MapData(X, Y).light_value(2) = ambient_light((Luz * 9) + MapData(X, Y).Luz + 1)
                            MapData(X, Y).light_value(3) = ambient_light((Luz * 9) + MapData(X, Y).Luz + 1)
                        ElseIf MapData(X, Y).Luz = 9 Then
                            MapData(X, Y).light_value(0) = extra_light(eE_Light.Oscuridad)
                            MapData(X, Y).light_value(1) = extra_light(eE_Light.Oscuridad)
                            MapData(X, Y).light_value(2) = extra_light(eE_Light.Oscuridad)
                           MapData(X, Y).light_value(3) = extra_light(eE_Light.Oscuridad)
                        ElseIf MapData(X, Y).Luz = 11 Then
                            MapData(X, Y).light_value(0) = extra_light(eE_Light.Azul1)
                            MapData(X, Y).light_value(1) = extra_light(eE_Light.Azul1)
                            MapData(X, Y).light_value(2) = extra_light(eE_Light.Azul1)
                           MapData(X, Y).light_value(3) = extra_light(eE_Light.Azul1)
                        ElseIf MapData(X, Y).Luz = 12 Then
                            MapData(X, Y).light_value(0) = extra_light(eE_Light.Azul2)
                            MapData(X, Y).light_value(1) = extra_light(eE_Light.Azul2)
                            MapData(X, Y).light_value(2) = extra_light(eE_Light.Azul2)
                           MapData(X, Y).light_value(3) = extra_light(eE_Light.Azul2)
                        ElseIf MapData(X, Y).Luz = 13 Then
                            MapData(X, Y).light_value(0) = extra_light(eE_Light.Azul3)
                            MapData(X, Y).light_value(1) = extra_light(eE_Light.Azul3)
                            MapData(X, Y).light_value(2) = extra_light(eE_Light.Azul3)
                           MapData(X, Y).light_value(3) = extra_light(eE_Light.Azul3)
                        End If
                    ElseIf MapData(X, Y).Luz > 0 Then 'Bordes, cargamos los cosos.
                    
                        If MapData(X, Y).LV(0) > 0 And MapData(X, Y).LV(0) < 9 Then 'Luces normales
                            MapData(X, Y).light_value(0) = ambient_light(((Luz * 9) + MapData(X, Y).LV(0) + 1))
                        ElseIf MapData(X, Y).LV(0) = 9 Then
                            MapData(X, Y).light_value(0) = extra_light(eE_Light.Oscuridad)
                        ElseIf MapData(X, Y).LV(0) = 11 Then
                            MapData(X, Y).light_value(0) = extra_light(eE_Light.Azul1)
                        ElseIf MapData(X, Y).LV(0) = 12 Then
                            MapData(X, Y).light_value(0) = extra_light(eE_Light.Azul2)
                        ElseIf MapData(X, Y).LV(0) = 13 Then
                            MapData(X, Y).light_value(0) = extra_light(eE_Light.Azul3)
                                                Else
                            MapData(X, Y).light_value(0) = 0
                        End If
                        If MapData(X, Y).LV(1) > 0 And MapData(X, Y).LV(1) < 9 Then 'Luces normales
                            MapData(X, Y).light_value(1) = ambient_light((Luz * 9) + MapData(X, Y).LV(1) + 1)
                        ElseIf MapData(X, Y).LV(1) = 9 Then
                            MapData(X, Y).light_value(1) = extra_light(eE_Light.Oscuridad)
                        ElseIf MapData(X, Y).LV(1) = 11 Then
                            MapData(X, Y).light_value(1) = extra_light(eE_Light.Azul1)
                        ElseIf MapData(X, Y).LV(1) = 12 Then
                            MapData(X, Y).light_value(1) = extra_light(eE_Light.Azul2)
                        ElseIf MapData(X, Y).LV(1) = 13 Then
                            MapData(X, Y).light_value(1) = extra_light(eE_Light.Azul3)
                                                Else
                            MapData(X, Y).light_value(1) = 0
                        End If
                        If MapData(X, Y).LV(2) > 0 And MapData(X, Y).LV(2) < 9 Then 'Luces normales
                            MapData(X, Y).light_value(2) = ambient_light((Luz * 9) + MapData(X, Y).LV(2) + 1)
                        ElseIf MapData(X, Y).LV(2) = 9 Then
                            MapData(X, Y).light_value(2) = extra_light(eE_Light.Oscuridad)
                        ElseIf MapData(X, Y).LV(2) = 11 Then
                            MapData(X, Y).light_value(2) = extra_light(eE_Light.Azul1)
                        ElseIf MapData(X, Y).LV(2) = 12 Then
                            MapData(X, Y).light_value(2) = extra_light(eE_Light.Azul2)
                        ElseIf MapData(X, Y).LV(2) = 13 Then
                            MapData(X, Y).light_value(2) = extra_light(eE_Light.Azul3)
                        Else
                            MapData(X, Y).light_value(2) = 0
                        End If
                        If MapData(X, Y).LV(3) > 0 And MapData(X, Y).LV(3) < 9 Then 'Luces normales
                            MapData(X, Y).light_value(3) = ambient_light((Luz * 9) + MapData(X, Y).LV(3) + 1)
                        ElseIf MapData(X, Y).LV(3) = 9 Then
                            MapData(X, Y).light_value(3) = extra_light(eE_Light.Oscuridad)
                        ElseIf MapData(X, Y).LV(3) = 11 Then
                            MapData(X, Y).light_value(3) = extra_light(eE_Light.Azul1)
                        ElseIf MapData(X, Y).LV(3) = 12 Then
                            MapData(X, Y).LV(3) = extra_light(eE_Light.Azul2)
                        ElseIf MapData(X, Y).LV(3) = 13 Then
                            MapData(X, Y).light_value(3) = extra_light(eE_Light.Azul3)
                                                Else
                            MapData(X, Y).light_value(3) = 0
                        End If

                    
                    End If


Next Y
Next X
base_light = ambient_light((Luz * 9) + 1)

End Sub

Public Sub DXEngine_iTextureRender(ByVal texture_index As Long, ByVal dest_x As Long, ByVal dest_y As Long, ByVal src_width As Long, _
                                            ByVal src_height As Long, ByRef rgb_list() As Long, ByVal Src_X As Long, _
                                            ByVal Src_Y As Long, ByVal dest_width As Long, ByVal dest_height As Long, _
                                            Optional ByVal alpha_blend As Boolean, Optional ByVal Angle As Single)
'**************************************************************
'This sub doesnt allow texture resizing
'
'**************************************************************
    Dim src_rect As RECT
    Dim dest_rect As RECT
    Dim temp_verts(3) As TLVERTEX
    Dim texture_height As Integer
    Dim texture_width As Integer
    Dim Texture As Direct3DTexture8
    Dim R(3) As Long
    'Set up the source rectangle
    With src_rect
        .Bottom = Src_Y + src_height - 1
        .left = Src_X
        .Right = Src_X + src_width - 1
        .top = Src_Y
    End With
        
    'Set up the destination rectangle
    With dest_rect
        .Bottom = dest_y + dest_height
        .left = dest_x
        .Right = dest_x + dest_width
        .top = dest_y
    End With
    If rgb_list(0) = 0 Then R(0) = base_light Else R(0) = rgb_list(0)
    If rgb_list(1) = 0 Then R(1) = base_light Else R(1) = rgb_list(1)
    If rgb_list(2) = 0 Then R(2) = base_light Else R(2) = rgb_list(2)
    If rgb_list(3) = 0 Then R(3) = base_light Else R(3) = rgb_list(3)
    
    'ESTO NO ME GUSTA
    Set Texture = DXPool.GetTexture(texture_index)
    Call DXPool.Texture_Dimension_Get(texture_index, texture_width, texture_height)
    If texture_width = 0 Then
        Set Texture = DXPool.GetTexture(1)
        Call DXPool.Texture_Dimension_Get(texture_index, texture_width, texture_height)
    End If
    
    'If texture_index = 15030 Then Stop
    'Set up the TempVerts(3) vertices
    Geometry_Create_iBox temp_verts(), dest_rect, src_rect, R(), texture_height, texture_width, Angle
    'Set Texture
    ddevice.SetTexture 0, Texture
    
    'Enable alpha-blending
    alpha_blend = False
    ddevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
'Call ddevice.SetTextureStageState(0, D3DTSS_ALPHAOP, D3DTOP_MODULATE)
'Call ddevice.SetTextureStageState(0, D3DTSS_ALPHAARG1, D3DTA_DIFFUSE Or D3DTA_TEXTURE)
'Call ddevice.SetTextureStageState(0, D3DTSS_ALPHAARG2, D3DTA_CURRENT)
'Call ddevice.SetTextureStageState(0, D3DTSS_COLORARG1, D3DTA_TEXTURE)
'Call ddevice.SetTextureStageState(0, D3DTSS_COLORARG2, D3DTA_TFACTOR)
'    Call ddevice.SetRenderState(D3DRS_TEXTUREFACTOR, D3DColorARGB(255, 0, 0, 0))
'Call ddevice.SetTextureStageState(0, D3DTSS_COLOROP, 18)
'CON EL 18  D3DTOP_MODULATEALPHA_ADDCOLOR hace las dos cosas, agrega alpha y da color. la intesidad del color se da por el valor
'no logro aplicar el GRADO DE ALPHA a las cosas.

    If alpha_blend Then
       'Set Rendering for alphablending
ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
        'ddevice.SetRenderState D3DRS_COLORVERTEX, 2
    End If
    
    'Draw the triangles that make up our square texture
    ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0))
    
    If alpha_blend Then
        'Set Rendering for colokeying
        ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    End If
    
    'Turn off alphablending after we're done
    'ddevice.SetRenderState D3DRS_ALPHABLENDENABLE, 0
End Sub
Private Function Geometry_Create_TLVertex(ByVal X As Single, ByVal Y As Single, ByVal z As Single, _
                                            ByVal rhw As Single, ByVal Color As Long, ByVal Specular As Long, tu As Single, _
                                            ByVal tv As Single) As TLVERTEX
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'**************************************************************
    Geometry_Create_TLVertex.X = X
    Geometry_Create_TLVertex.Y = Y
    Geometry_Create_TLVertex.z = z
    Geometry_Create_TLVertex.rhw = rhw
    Geometry_Create_TLVertex.Color = Color 'D3DColorARGB(155, 100, 100, 100)
    Geometry_Create_TLVertex.tu = tu
    Geometry_Create_TLVertex.tv = tv
End Function

Private Sub Geometry_Create_Box(ByRef verts() As TLVERTEX, ByRef Dest As RECT, ByRef src As RECT, ByRef rgb_list() As Long, _
                                Optional ByRef texture_width As Integer, Optional ByRef texture_height As Integer, Optional ByVal Angle As Single)
'**************************************************************
'Authors: Aaron Perkins;
'Last Modify Date: 5/07/2002
'
' * v1 *    v3
' |     \   |
' |     \   |
' |     \   |
' |     \   |
' |     \   |
' * v0 *    v2
'**************************************************************
    Dim x_center As Single
    Dim y_center As Single
    Dim radius As Single
    Dim x_Cor As Single
    Dim y_Cor As Single
    Dim left_point As Single
    Dim right_point As Single
    Dim temp As Single

    If Angle > 0 Then
        'Center coordinates on screen of the square
        x_center = Dest.left + (Dest.Right - Dest.left - 1) / 2
        y_center = Dest.top + (Dest.Bottom - Dest.top - 1) / 2
        
        'Calculate radius
        radius = Sqr((Dest.Right - x_center) ^ 2 + (Dest.Bottom - y_center) ^ 2)
        
        'Calculate left and right points
        temp = (Dest.Right - x_center) / radius
        right_point = Atn(temp / Sqr(-temp * temp + 1))
        left_point = PI - right_point
    End If
    
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If Angle = 0 Then
        x_Cor = Dest.left
        y_Cor = Dest.Bottom
    Else
        x_Cor = x_center + Cos(-left_point - Angle) * radius
        y_Cor = y_center - Sin(-left_point - Angle) * radius
    End If
    
    
    
    
    '0 - Bottom left vertex
    If texture_width And texture_height Then
        verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, src.left / texture_width, (src.Bottom + 1) / texture_height)
    Else
        verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, 0, 1)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If Angle = 0 Then
        x_Cor = Dest.left
        y_Cor = Dest.top
    Else
        x_Cor = x_center + Cos(left_point - Angle) * radius
        y_Cor = y_center - Sin(left_point - Angle) * radius
    End If
    
    
    '1 - Top left vertex
    If texture_width And texture_height Then
        verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, src.left / texture_width, src.top / texture_height)
    Else
        verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, 0, 0)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If Angle = 0 Then
        x_Cor = Dest.Right
        y_Cor = Dest.Bottom
    Else
        x_Cor = x_center + Cos(-right_point - Angle) * radius
        y_Cor = y_center - Sin(-right_point - Angle) * radius
    End If
    
    
    '2 - Bottom right vertex
    If texture_width And texture_height Then
        verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, (src.Right + 1) / texture_width, (src.Bottom + 1) / texture_height)
    Else
        verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, 1, 1)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If Angle = 0 Then
        x_Cor = Dest.Right
        y_Cor = Dest.top
    Else
        x_Cor = x_center + Cos(right_point - Angle) * radius
        y_Cor = y_center - Sin(right_point - Angle) * radius
    End If
    
    
    '3 - Top right vertex
    If texture_width And texture_height Then
        verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, (src.Right + 1) / texture_width, src.top / texture_height)
    Else
        verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, 1, 0)
    End If
End Sub
Private Sub Geometry_Create_iBox(ByRef verts() As TLVERTEX, ByRef Dest As RECT, ByRef src As RECT, ByRef rgb_list() As Long, _
                                Optional ByRef texture_width As Integer, Optional ByRef texture_height As Integer, Optional ByVal Angle As Single)
'**************************************************************
'Authors: Aaron Perkins;
'Last Modify Date: 5/07/2002
'
' * v1 *    v3
' |     \   |
' |     \   |
' |     \   |
' |     \   |
' |     \   |
' * v0 *    v2
'**************************************************************
    Dim x_center As Single
    Dim y_center As Single
    Dim radius As Single
    Dim x_Cor As Single
    Dim y_Cor As Single
    Dim left_point As Single
    Dim right_point As Single
    Dim temp As Single

    If Angle > 0 Then
        'Center coordinates on screen of the square
        x_center = Dest.left + (Dest.Right - Dest.left - 1) / 2
        y_center = Dest.top + (Dest.Bottom - Dest.top - 1) / 2
        
        'Calculate radius
        radius = Sqr((Dest.Right - x_center) ^ 2 + (Dest.Bottom - y_center) ^ 2)
        
        'Calculate left and right points
        temp = (Dest.Right - x_center) / radius
        right_point = Atn(temp / Sqr(-temp * temp + 1))
        left_point = PI - right_point
    End If
    
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If Angle = 0 Then
        x_Cor = Dest.left
        y_Cor = Dest.Bottom
    Else
        x_Cor = x_center + Cos(-left_point - Angle) * radius
        y_Cor = y_center - Sin(-left_point - Angle) * radius
    End If
    
    
    
    
    '0 - Bottom left vertex
    If texture_width And texture_height Then
        verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, src.left / texture_width, (src.Bottom + 1) / texture_height)
    Else
        verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, 0, 1)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If Angle = 0 Then
        x_Cor = Dest.left
        y_Cor = Dest.top
    Else
        x_Cor = x_center + Cos(left_point - Angle) * radius
        y_Cor = y_center - Sin(left_point - Angle) * radius
    End If
    
    
    '1 - Top left vertex
    If texture_width And texture_height Then
        verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, src.left / texture_width, src.top / texture_height)
    Else
        verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, 0, 0)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If Angle = 0 Then
        x_Cor = Dest.Right
        y_Cor = Dest.Bottom
    Else
        x_Cor = x_center + Cos(-right_point - Angle) * radius
        y_Cor = y_center - Sin(-right_point - Angle) * radius
    End If
    
    
    '2 - Bottom right vertex
    If texture_width And texture_height Then
        verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, (src.Right + 1) / texture_width, (src.Bottom + 1) / texture_height)
    Else
        verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, 1, 1)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If Angle = 0 Then
        x_Cor = Dest.Right
        y_Cor = Dest.top
    Else
        x_Cor = x_center + Cos(right_point - Angle) * radius
        y_Cor = y_center - Sin(right_point - Angle) * radius
    End If
    
    
    '3 - Top right vertex
    If texture_width And texture_height Then
        verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, (src.Right + 1) / texture_width, src.top / texture_height)
    Else
        verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, 1, 0)
    End If
End Sub
Public Sub DXEngine_GraphicTextRender(Font_Index As Integer, ByVal Text As String, ByVal top As Long, ByVal left As Long, _
                                  ByVal Color As Long)

    If Len(Text) > 255 Then Exit Sub
    
    Dim i As Byte
    Dim X As Integer
    Dim rgb_list(3) As Long
    
    For i = 0 To 3
        rgb_list(i) = Color
    Next i
    
    X = -1
    Dim Char As Integer
    For i = 1 To Len(Text)
        Char = AscB(mid$(Text, i, 1)) - 32
        
        If Char = 0 Then
            X = X + 1
        Else
            X = X + 1
            Call DXEngine_TextureRenderAdvance(gfont_list(Font_Index).texture_index, left + X * gfont_list(Font_Index).Char_Size, _
                                                        top, gfont_list(Font_Index).Caracteres(Char).Src_X, gfont_list(Font_Index).Caracteres(Char).Src_Y, _
                                                            gfont_list(Font_Index).Char_Size, gfont_list(Font_Index).Char_Size, gfont_list(Font_Index).Char_Size, gfont_list(Font_Index).Char_Size, _
                                                                rgb_list(), False)
        End If
    Next i
    
    
    
End Sub

Public Sub DXEngine_Deinitialize()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
On Error Resume Next

    'El manager de texturas es ahora independiente del engine.
    Call DXPool.Texture_Remove_All
    
    Set d3dx = Nothing
    Set ddevice = Nothing
    Set d3d = Nothing
    Set dx = Nothing
    Set DXPool = Nothing
End Sub

Private Sub LoadChars(ByVal Font_Index As Integer)
    Dim i As Integer
    Dim X As Integer
    Dim Y As Integer
    
    For i = 0 To 255
        With gfont_list(Font_Index).Caracteres(i)
            X = (i Mod 16) * gfont_list(Font_Index).Char_Size
            If X = 0 Then '16 chars per line
                Y = Y + 1
            End If
            .Src_X = X
            .Src_Y = (Y * gfont_list(Font_Index).Char_Size) - gfont_list(Font_Index).Char_Size
        End With
    Next i
End Sub
Public Sub LoadGraphicFonts()
    Dim i As Byte
    Dim file_path As String

    file_path = DirIndex & "GUIFonts.ini"

    If General_File_Exist(file_path, vbArchive) Then
        gfont_count = general_var_get(file_path, "INIT", "FontCount")
        If gfont_count > 0 Then
            ReDim gfont_list(1 To gfont_count) As tGraphicFont
            For i = 1 To gfont_count
                With gfont_list(i)
                    .Char_Size = general_var_get(file_path, "FONT" & i, "Size")
                    .texture_index = general_var_get(file_path, "FONT" & i, "Graphic")
                    If .texture_index > 0 Then Call DXPool.Texture_Load(.texture_index, 0)
                    LoadChars (i)
                End With
            Next i
        End If
    End If
End Sub

Public Sub DXEngine_StatsRender()
    'fps
    'Call DXEngine_TextRender(1, FPS & " FPS", 0, 0, D3DColorXRGB(255, 255, 255))
    
    modDXEngine.DrawText 0, 0, FPS & " FPS", D3DWHITE
    
    
End Sub

Private Sub Device_Flip()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    'Draw the graphics to the front buffer.
    ddevice.Present ByVal 0&, ByVal 0&, screen_hwnd, ByVal 0&
    
    

End Sub

Private Sub DeviceRenderStates()
    With ddevice
        'Set the vertex shader to an FVF that contains texture coords,
        'and transformed and lit vertex coords.
        .SetVertexShader FVF
        .SetRenderState D3DRS_LIGHTING, False
        .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        .SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        .SetRenderState D3DRS_ALPHABLENDENABLE, True
        
        'No se para q mierda sera esto.
        .SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
        .SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
        .SetRenderState D3DRS_ZENABLE, True
        .SetRenderState D3DRS_ZWRITEENABLE, False
        
        'Particle engine settings
        '.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
        '.SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
        '.SetRenderState D3DRS_POINTSCALE_ENABLE, 0


    End With
End Sub

Private Sub Font_Make(ByVal Style As String, ByVal size As Long, ByVal italic As Boolean, ByVal bold As Boolean)
    font_count = font_count + 1
    ReDim Preserve font_list(1 To font_count)
    
    Dim font_desc As IFont
    Dim fnt As New StdFont
    fnt.Name = Style
    fnt.size = size
    fnt.bold = bold
    fnt.italic = italic
    Set font_desc = fnt
    font_list(font_count).size = size
    Set font_list(font_count).dFont = d3dx.CreateFont(ddevice, font_desc.hFont)
End Sub

Private Sub LoadFonts()
    Dim num_fonts As Integer
    Dim i As Integer
    Dim file_path As String
    
    file_path = DirIndex & "fonts.ini"
    
    If Not General_File_Exist(file_path, vbArchive) Then Exit Sub
    
    num_fonts = general_var_get(file_path, "INIT", "FontCount")
    
    For i = 1 To num_fonts
        Call Font_Make(general_var_get(file_path, "FONT" & i, "Name"), general_var_get(file_path, "FONT" & i, "Size"), general_var_get(file_path, "FONT" & i, "Cursiva"), general_var_get(file_path, "FONT" & i, "Negrita"))
    Next i
End Sub
Public Sub DXEngine_TextRender(ByVal Font_Index As Integer, ByVal Text As String, ByVal left As Integer, ByVal top As Integer, ByVal Color As Long, Optional ByVal Alingment As Byte = DT_LEFT, Optional ByVal Width As Integer = 0, Optional ByVal Height As Integer = 0)
    If Not Font_Check(Font_Index) Then Exit Sub
    
    Dim TextRect As RECT 'This defines where it will be
    'Dim BorderColor As Long
    
    'Set width and height if no specified
    If Width = 0 Then Width = Len(Text) * (font_list(Font_Index).size + 1)
    If Height = 0 Then Height = font_list(Font_Index).size * 2
    
    'DrawBorder
    
    'BorderColor = D3DColorXRGB(0, 0, 0)
    
    'TextRect.top = top - 1
    'TextRect.left = left
    'TextRect.bottom = top + height
    'TextRect.Right = left + width
    'd3dx.DrawText font_list(Font_Index).dFont, BorderColor, Text, TextRect, Alingment
    'TextRect.top = top
    'TextRect.left = left - 1
    'TextRect.bottom = top + height
    'TextRect.Right = left + width
    'd3dx.DrawText font_list(Font_Index).dFont, BorderColor, Text, TextRect, Alingment
    'TextRect.top = top + 1
    'TextRect.left = left
    'TextRect.bottom = top + height
    'TextRect.Right = left + width
    'd3dx.DrawText font_list(Font_Index).dFont, BorderColor, Text, TextRect, Alingment
    'TextRect.top = top
    'TextRect.left = left + 1
    'TextRect.bottom = top + height
    'TextRect.Right = left + width
    'd3dx.DrawText font_list(Font_Index).dFont, BorderColor, Text, TextRect, Alingment
    
    TextRect.top = top
    TextRect.left = left
    TextRect.Bottom = top + Height
    TextRect.Right = left + Width
    d3dx.DrawText font_list(Font_Index).dFont, Color, Text, TextRect, Alingment

End Sub
Private Function Font_Check(ByVal Font_Index As Long) As Boolean
    If Font_Index > 0 And Font_Index <= font_count Then
        Font_Check = True
    End If
End Function

Private Sub Fonts_Destroy()
    Dim i As Integer
    
    For i = 1 To font_count
        Set font_list(i).dFont = Nothing
        font_list(i).size = 0
    Next i
    font_count = 0
End Sub

Public Function D3DColorValueGet(ByVal A As Byte, ByVal R As Byte, ByVal G As Byte, ByVal B As Byte) As D3DCOLORVALUE
    D3DColorValueGet.A = A
    D3DColorValueGet.R = R
    D3DColorValueGet.G = G
    D3DColorValueGet.B = B
End Function


Public Sub DibujareEnHwnd(ByVal PIC As Long, ByVal GrhIndex As Integer, ByRef src_rect As RECT, ByVal X As Integer, ByVal Y As Integer, ByVal PRESENTO As Boolean)

Dim DestRect As RECT
Dim tX As Byte
Dim tY As Byte
    
  X = X
  Y = Y
  
   DestRect.top = Y
   DestRect.left = X
   DestRect.Bottom = Y + src_rect.Bottom - src_rect.top
   DestRect.Right = X + src_rect.Right - src_rect.left
   If src_rect.Bottom <= 0 Or src_rect.Right <= 0 Or src_rect.left = src_rect.Right Or src_rect.top = src_rect.Bottom Then Exit Sub
   
   
   ddevice.Clear 1, DestRect, D3DCLEAR_TARGET, &H0, ByVal 0, 0
   ddevice.BeginScene
   Draw_RAWGrhindex GrhIndex, src_rect, X, Y

   ddevice.EndScene
   

   
   If PRESENTO Then ddevice.Present src_rect, DestRect, PIC, ByVal 0


End Sub
Public Sub DibujareEnHwnd2(ByVal PIC As Long, ByVal nIndex As Integer, ByRef src_rect As RECT, ByVal X As Integer, ByVal Y As Integer, ByVal PRESENTO As Boolean, Optional ByVal ForceSize As Boolean, Optional ByVal ForceW As Integer, Optional ByVal ForceH As Integer)

Dim DestRect As RECT
Dim tX As Byte
Dim tY As Byte
    
  X = X
  Y = Y
  
   DestRect.top = Y
   DestRect.left = X
   
   If ForceSize Then
    If Y + (src_rect.Bottom - src_rect.top) > ForceH Then
     DestRect.Bottom = ForceH
    Else
    DestRect.Bottom = Y + src_rect.Bottom - src_rect.top
    End If
    If X + (src_rect.Right - src_rect.left) > ForceW Then
     DestRect.Right = ForceW
    Else
    DestRect.Right = Y + src_rect.Right - src_rect.left
    End If
      
   
   Else
   DestRect.Bottom = Y + src_rect.Bottom - src_rect.top
   DestRect.Right = X + src_rect.Right - src_rect.left
   End If
   If src_rect.Bottom <= 0 Or src_rect.Right <= 0 Or src_rect.left = src_rect.Right Or src_rect.top = src_rect.Bottom Then Exit Sub
   
   

   ddevice.BeginScene
   If PRESENTO Then ddevice.Clear 1, src_rect, D3DCLEAR_TARGET, 0, 0, ByVal 0
   Draw_RAWnIndex nIndex, X, Y

   ddevice.EndScene
   

   
   If PRESENTO Then ddevice.Present src_rect, DestRect, PIC, ByVal 0


End Sub
Public Function DameWidthTextura(ByVal t As Integer) As Integer

Dim d3dTextures As D3D8Textures
DXPool.Texture_Dimension_Get t, d3dTextures.texwidth, d3dTextures.texheight

DameWidthTextura = d3dTextures.texwidth


End Function
Public Function DameHeightTextura(ByVal t As Integer) As Integer

Dim d3dTextures As D3D8Textures
DXPool.Texture_Dimension_Get t, d3dTextures.texwidth, d3dTextures.texheight

DameHeightTextura = d3dTextures.texheight


End Function

Public Sub DibujareEnHwnd3(ByVal PIC As Long, ByVal Graf As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal PRESENTO As Boolean, Optional ByVal ForceSize As Boolean, Optional ByVal ForceW As Integer, Optional ByVal ForceH As Integer)

Dim DestRect As RECT
Dim tX As Byte
Dim tY As Byte
Dim src_rect As RECT
Dim d3dTextures As D3D8Textures
Dim light_value(0 To 3) As Long
Dim verts(3) As TLVERTEX

Set d3dTextures.Texture = DXPool.GetTexture(Graf)
DXPool.Texture_Dimension_Get Graf, d3dTextures.texwidth, d3dTextures.texheight
ddevice.SetTexture 0, d3dTextures.Texture
    


  X = X
  Y = Y

   src_rect.top = 0
   src_rect.left = 0
   src_rect.Right = d3dTextures.texwidth
   src_rect.Bottom = d3dTextures.texheight
    
    
   DestRect.top = Y
   DestRect.left = X
   
   If ForceSize = False Then
        DestRect.Bottom = Y + src_rect.Bottom - src_rect.top
        DestRect.Right = X + src_rect.Right - src_rect.left
   Else
       If X + (src_rect.Right - src_rect.left) >= ForceW Then
        DestRect.Right = X + (ForceW - X)
       Else
        DestRect.Right = X + src_rect.Right - src_rect.left
       End If
       If Y + (src_rect.Bottom - src_rect.top) >= ForceH Then
        DestRect.Bottom = Y + (ForceH - Y)
       Else
        DestRect.Bottom = Y + src_rect.Bottom - src_rect.top
       End If
       
   End If
   
   If src_rect.Bottom <= 0 Or src_rect.Right <= 0 Or src_rect.left = src_rect.Right Or src_rect.top = src_rect.Bottom Then Exit Sub
   
   
   ddevice.Clear 1, DestRect, D3DCLEAR_TARGET, &H0, ByVal 0, 0
   ddevice.BeginScene



        With verts(2)
            .X = X
            .Y = Y + d3dTextures.texheight
            .tu = 0 / (d3dTextures.texwidth)
            .tv = (0 + d3dTextures.texheight) / (d3dTextures.texheight)
            .rhw = 1
            .Color = -1
        End With
        With verts(0)
            .X = X
            .Y = Y
            .tu = 0 / (d3dTextures.texwidth)
            .tv = 0 / (d3dTextures.texheight)
            .rhw = 1
            .Color = -1

        End With
        
        With verts(3)
            .X = X + d3dTextures.texwidth
            .Y = Y + d3dTextures.texheight
            .tu = (0 + d3dTextures.texwidth) / (d3dTextures.texwidth)
            .tv = (0 + d3dTextures.texheight) / (d3dTextures.texheight)
            .rhw = 1
            .Color = -1

        End With
        
        With verts(1)
            .X = X + d3dTextures.texwidth
            .Y = Y
            .tu = (0 + d3dTextures.texwidth) / (d3dTextures.texwidth)
            .tv = 0 / (d3dTextures.texheight)
            .rhw = 1
            .Color = -1

        End With
    
   

  
  

    ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, verts(0), 28



   ddevice.EndScene
   

   
   If PRESENTO Then ddevice.Present src_rect, DestRect, PIC, ByVal 0


End Sub
Public Sub PRESENTAR_PREVIEW(ByRef R As RECT, ByVal PIC As Long)



ddevice.Present R, R, PIC, ByVal 0
End Sub
Private Sub Draw_RAWGrhindex(ByRef CurrentGrhIndex As Integer, ByRef src_rect As RECT, ByVal X As Integer, ByVal Y As Integer)
    Dim d3dTextures As D3D8Textures

    Dim light_value(0 To 3) As Long
    Dim verts(3) As TLVERTEX
    Set d3dTextures.Texture = DXPool.GetTexture(grh_list(CurrentGrhIndex).texture_index)
    DXPool.Texture_Dimension_Get grh_list(CurrentGrhIndex).texture_index, d3dTextures.texwidth, d3dTextures.texheight
    ddevice.SetTexture 0, d3dTextures.Texture
    


      If ((d3dTextures.texwidth - 1)) And ((d3dTextures.texheight - 1)) Then
    
        With verts(2)
            .X = X
            .Y = Y + grh_list(CurrentGrhIndex).src_height
            .tu = grh_list(CurrentGrhIndex).Src_X / (d3dTextures.texwidth - 1)
            .tv = (grh_list(CurrentGrhIndex).Src_Y + grh_list(CurrentGrhIndex).src_height) / (d3dTextures.texheight - 1)
            .rhw = 1
            .Color = -1
        End With
        With verts(0)
            .X = X
            .Y = Y
            .tu = grh_list(CurrentGrhIndex).Src_X / (d3dTextures.texwidth - 1)
            .tv = grh_list(CurrentGrhIndex).Src_Y / (d3dTextures.texheight - 1)
            .rhw = 1
            .Color = -1

        End With
        
        With verts(3)
            .X = X + grh_list(CurrentGrhIndex).src_width
            .Y = Y + grh_list(CurrentGrhIndex).src_height
            .tu = (grh_list(CurrentGrhIndex).Src_X + grh_list(CurrentGrhIndex).src_width) / (d3dTextures.texwidth - 1)
            .tv = (grh_list(CurrentGrhIndex).Src_Y + grh_list(CurrentGrhIndex).src_height) / (d3dTextures.texheight - 1)
            .rhw = 1
            .Color = -1

        End With
        
        With verts(1)
            .X = X + grh_list(CurrentGrhIndex).src_width
            .Y = Y
            .tu = (grh_list(CurrentGrhIndex).Src_X + grh_list(CurrentGrhIndex).src_width) / (d3dTextures.texwidth - 1)
            .tv = grh_list(CurrentGrhIndex).Src_Y / (d3dTextures.texheight - 1)
            .rhw = 1
            .Color = -1

        End With
    
   
   Else
        With verts(0)
            .X = X
            .Y = Y + grh_list(CurrentGrhIndex).src_height
            .tu = 0
            .tv = 0
            .rhw = 1
            .Color = -1

        End With
        With verts(1)
            .X = X
            .Y = Y
            .tu = 0
            .tv = 0
            .rhw = 1
            .Color = -1

        End With
        
        With verts(2)
            .X = X + grh_list(CurrentGrhIndex).src_width
            .Y = Y + grh_list(CurrentGrhIndex).src_height
            .tu = 0
            .tv = 0
            .rhw = 1
            .Color = -1

        End With
        
        With verts(3)
            .X = X + grh_list(CurrentGrhIndex).src_width
            .Y = Y + grh_list(CurrentGrhIndex).src_height
            .tu = 0
            .tv = 0
            .rhw = 1
            .Color = -1

        End With
    
    End If
  

    ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, verts(0), 28
End Sub
Private Sub Draw_RAWnIndex(ByVal nIndex As Integer, ByVal X As Integer, ByVal Y As Integer)
    Dim d3dTextures As D3D8Textures

    Dim light_value(0 To 3) As Long
    Dim verts(3) As TLVERTEX
    
    Set d3dTextures.Texture = DXPool.GetTexture(NewIndexData(nIndex).OverWriteGrafico)
    DXPool.Texture_Dimension_Get NewIndexData(nIndex).OverWriteGrafico, d3dTextures.texwidth, d3dTextures.texheight
    ddevice.SetTexture 0, d3dTextures.Texture
    

    Dim jx As Integer
    Dim jy As Integer
    Dim jw As Integer
    Dim jh As Integer
    If NewIndexData(nIndex).Estatic > 0 Then
    With EstaticData(NewIndexData(nIndex).Estatic)
    jx = .L
    jy = .t
    jw = .W
    jh = .H
    
    End With
    ElseIf NewIndexData(nIndex).Dinamica > 0 Then
        With NewAnimationData(NewIndexData(nIndex).Dinamica)
            jx = .Indice(1).X
            jy = .Indice(1).Y
            jw = .Width
            jh = .Height
        
        End With
    End If
        With verts(2)
            .X = X
            .Y = Y + jh
            .tu = jx / (d3dTextures.texwidth)
            .tv = (jy + jh) / (d3dTextures.texheight)
            .rhw = 1
            .Color = -1
        End With
        With verts(0)
            .X = X
            .Y = Y
            .tu = jx / (d3dTextures.texwidth)
            .tv = jy / (d3dTextures.texheight)
            .rhw = 1
            .Color = -1

        End With
        
        With verts(3)
            .X = X + jw
            .Y = Y + jh
            .tu = (jx + jw) / (d3dTextures.texwidth)
            .tv = (jy + jh) / (d3dTextures.texheight)
            .rhw = 1
            .Color = -1

        End With
        
        With verts(1)
            .X = X + jw
            .Y = Y
            .tu = (jx + jw) / (d3dTextures.texwidth)
            .tv = jy / (d3dTextures.texheight)
            .rhw = 1
            .Color = -1

        End With
    
  
  

    ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, verts(0), 28
End Sub

Public Sub DXEngine_BeginSecondaryRender()
    Device_Clear
    ddevice.BeginScene
End Sub
Public Sub DXEngine_EndSecondaryRender(ByVal hWnd As Long, ByVal Width As Integer, ByVal Height As Integer)
    Dim DR As RECT
    DR.left = 0
    DR.top = 0
    DR.Bottom = Height
    DR.Right = Width
    
    ddevice.EndScene
    ddevice.Present DR, ByVal 0&, hWnd, ByVal 0&
End Sub

Public Sub DXEngine_DrawBox(ByVal X As Integer, ByVal Y As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Color As Long, Optional ByVal border_width = 1)
    Dim VertexB(3) As TLVERTEX
    Dim box_rect As RECT
    
    With box_rect
        .Bottom = Y + Height
        .left = X
        .Right = X + Width
        .top = Y
    End With
    
    ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
    ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
        
    ddevice.SetTexture 0, Nothing
    
    'Upper Line
    VertexB(0) = Geometry_Create_TLVertex(box_rect.left, box_rect.top, 0, 1, Color, 0, 0, 0)
    VertexB(1) = Geometry_Create_TLVertex(box_rect.Right, box_rect.top, 0, 1, Color, 0, 0, 0)
    VertexB(2) = Geometry_Create_TLVertex(box_rect.left, box_rect.top + border_width, 0, 1, Color, 0, 0, 0)
    VertexB(3) = Geometry_Create_TLVertex(box_rect.Right, box_rect.top + border_width, 0, 1, Color, 0, 0, 0)
    ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexB(0), Len(VertexB(0))
    'Left Line
    VertexB(0) = Geometry_Create_TLVertex(box_rect.left + border_width, box_rect.top, 0, 1, Color, 0, 0, 0)
    VertexB(1) = Geometry_Create_TLVertex(box_rect.left + border_width, box_rect.Bottom, 0, 1, Color, 0, 0, 0)
    VertexB(2) = Geometry_Create_TLVertex(box_rect.left, box_rect.top, 0, 2, Color, 0, 0, 0)
    VertexB(3) = Geometry_Create_TLVertex(box_rect.left, box_rect.Bottom, 0, 2, Color, 0, 0, 0)
    ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexB(0), Len(VertexB(0))
    'Right Border
    VertexB(0) = Geometry_Create_TLVertex(box_rect.Right, box_rect.top, 0, 1, Color, 0, 0, 0)
    VertexB(1) = Geometry_Create_TLVertex(box_rect.Right, box_rect.Bottom, 0, 1, Color, 0, 0, 0)
    VertexB(2) = Geometry_Create_TLVertex(box_rect.Right - border_width, box_rect.top, 0, 3, Color, 0, 0, 0)
    VertexB(3) = Geometry_Create_TLVertex(box_rect.Right - border_width, box_rect.Bottom, 0, 3, Color, 0, 0, 0)
    ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexB(0), Len(VertexB(0))
    'Bottom Border
    VertexB(0) = Geometry_Create_TLVertex(box_rect.left, box_rect.Bottom - border_width, 0, 1, Color, 0, 0, 0)
    VertexB(1) = Geometry_Create_TLVertex(box_rect.Right, box_rect.Bottom - border_width, 0, 1, Color, 0, 0, 0)
    VertexB(2) = Geometry_Create_TLVertex(box_rect.left, box_rect.Bottom, 0, 1, Color, 0, 0, 0)
    VertexB(3) = Geometry_Create_TLVertex(box_rect.Right, box_rect.Bottom, 0, 1, Color, 0, 0, 0)
    ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexB(0), Len(VertexB(0))
    
    ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
End Sub
Public Sub D3DColorToRgbList(rgb_list() As Long, Color As D3DCOLORVALUE)
    rgb_list(0) = D3DColorARGB(Color.A, Color.R, Color.G, Color.B)
    rgb_list(1) = rgb_list(0)
    rgb_list(2) = rgb_list(0)
    rgb_list(3) = rgb_list(0)
End Sub


Private Sub Engine_Render_Text(ByRef UseFont As CustomFont, ByVal Text As String, ByVal X As Long, ByVal Y As Long, ByVal Color As Long, Optional ByVal center As Boolean = False, Optional ByVal Alpha As Byte = 255)
Dim TempVA(0 To 3) As TLVERTEX
Dim tempstr() As String
Dim Count As Integer
Dim ascii() As Byte
Dim Row As Integer
Dim u As Single
Dim v As Single
Dim i As Long
Dim j As Long
Dim KeyPhrase As Byte
Dim TempColor As Long
Dim ResetColor As Byte
Dim SrcRect As RECT
Dim v2 As D3DVECTOR2
Dim v3 As D3DVECTOR2
Dim YOffset As Single
 
    ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    'D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, True
   
    'Check if we have the device
    If ddevice.TestCooperativeLevel <> D3D_OK Then Exit Sub
 
    'Check for valid text to render
    If LenB(Text) = 0 Then Exit Sub
   
    'Get the text into arrays (split by vbCrLf)
    tempstr = Split(Text, vbCrLf)
   
    'Set the temp color (or else the first character has no color)
    TempColor = Color
 
    'Set the texture
    ddevice.SetTexture 0, UseFont.Texture
   
    If center Then
        X = X - Engine_GetTextWidth(cfonts(1), Text) * 0.5
    End If
   
    'Loop through each line if there are line breaks (vbCrLf)
    For i = 0 To UBound(tempstr)
        If Len(tempstr(i)) > 0 Then
            YOffset = i * UseFont.CharHeight
            Count = 0
       
            'Convert the characters to the ascii value
            ascii() = StrConv(tempstr(i), vbFromUnicode)
       
            'Loop through the characters
            For j = 1 To Len(tempstr(i))
 
                'Check for a key phrase
                'If ascii(j - 1) = 124 Then 'If Ascii = "|"
                '    KeyPhrase = (Not KeyPhrase)  'TempColor = ARGB 255/255/0/0
                '    If KeyPhrase Then TempColor = ARGB(255, 0, 0, alpha) Else ResetColor = 1
                'Else
 
                    'Render with triangles
                    'If AlternateRender = 0 Then
 
                        'Copy from the cached vertex array to the temp vertex array
                        CopyMemory TempVA(0), UseFont.HeaderInfo.CharVA(ascii(j - 1)).Vertex(0), 32 * 4
 
                        'Set up the verticies
                        TempVA(0).X = X + Count
                        TempVA(0).Y = Y + YOffset
                       
                        TempVA(1).X = TempVA(1).X + X + Count
                        TempVA(1).Y = TempVA(0).Y
 
                        TempVA(2).X = TempVA(0).X
                        TempVA(2).Y = TempVA(2).Y + TempVA(0).Y
 
                        TempVA(3).X = TempVA(1).X
                        TempVA(3).Y = TempVA(2).Y
                       
                        'Set the colors
                        TempVA(0).Color = TempColor
                        TempVA(1).Color = TempColor
                        TempVA(2).Color = TempColor
                        TempVA(3).Color = TempColor
                       
                        'Draw the verticies
                        ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, TempVA(0), Len(TempVA(0))
                       
                     
                    'Shift over the the position to render the next character
                    Count = Count + UseFont.HeaderInfo.CharWidth(ascii(j - 1))
               
                'End If
               
                'Check to reset the color
                If ResetColor Then
                    ResetColor = 0
                    TempColor = Color
                End If
               
            Next j
           
        End If
    Next i
   
End Sub
Private Function Engine_GetTextWidth(ByRef UseFont As CustomFont, ByVal Text As String) As Integer
'***************************************************
'Returns the width of text
'More info: [url=http://www.vbgore.com/GameClient.TileEngine.Engine_GetTextWidth]http://www.vbgore.com/GameClient.TileEn ... tTextWidth[/url]
'***************************************************
Dim i As Integer
 
    'Make sure we have text
    If LenB(Text) = 0 Then Exit Function
   
    'Loop through the text
    For i = 1 To Len(Text)
       
        'Add up the stored character widths
        Engine_GetTextWidth = Engine_GetTextWidth + UseFont.HeaderInfo.CharWidth(Asc(mid$(Text, i, 1)))
       
    Next i
 
End Function
 
Sub Engine_Init_FontTextures()
On Error GoTo eDebug:
'*****************************************************************
'Init the custom font textures
'More info: [url=http://www.vbgore.com/GameClient.TileEngine.Engine_Init_FontTextures]http://www.vbgore.com/GameClient.TileEn ... ntTextures[/url]
'*****************************************************************
Dim TexInfo As D3DXIMAGE_INFO_A
 
    'Check if we have the device
    If ddevice.TestCooperativeLevel <> D3D_OK Then Exit Sub
 
    '*** Default font ***
   
    'Set the texture
    Set cfonts(1).Texture = d3dx.CreateTextureFromFileEx(ddevice, App.PATH & "\Resources\init\tahoma.bmp", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFF000000, ByVal 0, ByVal 0)
    'Store the size of the texture
    cfonts(1).TextureSize.X = TexInfo.Width
    cfonts(1).TextureSize.Y = TexInfo.Height
   
    Exit Sub
eDebug:
    If Err.Number = "-2005529767" Then
        MsgBox "Error en la textura utilizada de DirectX 8", vbCritical
        End
    End If
    End
 
End Sub
 
Sub Engine_Init_FontSettings()
'*********************************************************
'****** Coded by Dunkan ([email=emanuel.m@dunkancorp.com]emanuel.m@dunkancorp.com[/email]) *******
'*********************************************************
Dim FileNum As Byte
Dim LoopChar As Long
Dim Row As Single
Dim u As Single
Dim v As Single
 
    '*** Default font ***
 
    'Load the header information
    FileNum = FreeFile
    Open App.PATH & "\RESOURCES\INIT\tahoma.dat" For Binary As #FileNum
        Get #FileNum, , cfonts(1).HeaderInfo
    Close #FileNum
   
    'Calculate some common values
    cfonts(1).CharHeight = cfonts(1).HeaderInfo.CellHeight - 4
    cfonts(1).RowPitch = cfonts(1).HeaderInfo.BitmapWidth \ cfonts(1).HeaderInfo.CellWidth
    cfonts(1).ColFactor = cfonts(1).HeaderInfo.CellWidth / cfonts(1).HeaderInfo.BitmapWidth
    cfonts(1).RowFactor = cfonts(1).HeaderInfo.CellHeight / cfonts(1).HeaderInfo.BitmapHeight
   
    'Cache the verticies used to draw the character (only requires setting the color and adding to the X/Y values)
    For LoopChar = 0 To 255
       
        'tU and tV value (basically tU = BitmapXPosition / BitmapWidth, and height for tV)
        Row = (LoopChar - cfonts(1).HeaderInfo.BaseCharOffset) \ cfonts(1).RowPitch
        u = ((LoopChar - cfonts(1).HeaderInfo.BaseCharOffset) - (Row * cfonts(1).RowPitch)) * cfonts(1).ColFactor
        v = Row * cfonts(1).RowFactor
 
        'Set the verticies
        With cfonts(1).HeaderInfo.CharVA(LoopChar)
            .Vertex(0).Color = D3DColorARGB(255, 0, 0, 0)   'Black is the most common color
            .Vertex(0).rhw = 1
            .Vertex(0).tu = u
            .Vertex(0).tv = v
            .Vertex(0).X = 0
            .Vertex(0).Y = 0
            .Vertex(0).z = 0
           
            .Vertex(1).Color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(1).rhw = 1
            .Vertex(1).tu = u + cfonts(1).ColFactor
            .Vertex(1).tv = v
            .Vertex(1).X = cfonts(1).HeaderInfo.CellWidth
            .Vertex(1).Y = 0
            .Vertex(1).z = 0
           
            .Vertex(2).Color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(2).rhw = 1
            .Vertex(2).tu = u
            .Vertex(2).tv = v + cfonts(1).RowFactor
            .Vertex(2).X = 0
            .Vertex(2).Y = cfonts(1).HeaderInfo.CellHeight
            .Vertex(2).z = 0
           
            .Vertex(3).Color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(3).rhw = 1
            .Vertex(3).tu = u + cfonts(1).ColFactor
            .Vertex(3).tv = v + cfonts(1).RowFactor
            .Vertex(3).X = cfonts(1).HeaderInfo.CellWidth
            .Vertex(3).Y = cfonts(1).HeaderInfo.CellHeight
            .Vertex(3).z = 0
        End With
       
    Next LoopChar
 
End Sub
Public Function ARGBtoD3DCOLORVALUE(ByVal ARGB As Long, ByRef Color As D3DCOLORVALUE)
Dim Dest(3) As Byte
CopyMemory Dest(0), ARGB, 4
Color.A = Dest(3)
Color.R = Dest(2)
Color.G = Dest(1)
Color.B = Dest(0)
End Function
Public Sub DrawText(ByVal left As Long, ByVal top As Long, ByVal Text As String, ByVal Colorx As Long, Optional ByVal Alpha As Byte = 255, Optional ByVal center As Boolean = False)
Dim Color As Long
If Text = "0" Then Exit Sub

If top <= 0 Or left <= 0 Or Colorx = 0 Or LenB(Text) = 0 Then Exit Sub
    If Alpha <> 255 Then
        Dim aux As D3DCOLORVALUE
        ARGBtoD3DCOLORVALUE Colorx, aux
        Color = D3DColorARGB(Alpha, aux.R, aux.G, aux.B)
        
    Else
        Color = Colorx
    End If
        Engine_Render_Text cfonts(1), Text, left, top, Color, center, Alpha
End Sub

Public Sub SPOTLIGHTS_LOADDAT()
Dim S As String
Dim i As Long

Dim A As Byte
Dim R As Byte
Dim G As Byte
Dim B As Byte


S = App.PATH & "\RESOURCES\INIT\SPOTLIGHTS.DAT"

NUM_SPOTLIGHTS_COLORES = Val(GetVar(S, "COLORES", "NUM_COLORES"))

frmMain.COLOREXTRA.AddItem "NINGUNO"
If NUM_SPOTLIGHTS_COLORES > 0 Then
ReDim SPOTLIGHTS_COLORES(1 To NUM_SPOTLIGHTS_COLORES)
    For i = 1 To NUM_SPOTLIGHTS_COLORES
        A = Val(GetVar(S, "COLOR" & i, "A"))
        R = Val(GetVar(S, "COLOR" & i, "R"))
        G = Val(GetVar(S, "COLOR" & i, "G"))
        B = Val(GetVar(S, "COLOR" & i, "B"))
        SPOTLIGHTS_COLORES(i) = D3DColorARGB(A, R, G, B)
        frmMain.COLORSPOT.AddItem GetVar(S, "COLOR" & i, "Nombre")
        frmMain.COLOREXTRA.AddItem GetVar(S, "COLOR" & i, "Nombre")
        
    Next i
End If
frmMain.COLOREXTRA.AddItem "CUSTOM"
frmMain.COLORSPOT.AddItem "CUSTOM"

frmMain.COLOREXTRA.ListIndex = 0
frmMain.COLORSPOT.ListIndex = 0


frmMain.SPOT_ANIM.AddItem "INANIMADA"

NUM_SPOTLIGHTS_ANIMATION = Val(GetVar(S, "ANIMACIONES", "NUM_ANIM"))
If NUM_SPOTLIGHTS_ANIMATION > 0 Then
ReDim SPOTLIGHTS_ANIMATION(1 To NUM_SPOTLIGHTS_ANIMATION)
For i = 1 To NUM_SPOTLIGHTS_ANIMATION

    SPOTLIGHTS_ANIMATION(i) = Val(GetVar(S, "ANIM" & i, "Indice"))
    frmMain.SPOT_ANIM.AddItem GetVar(S, "ANIM" & i, "Nombre")
Next i
End If
frmMain.SPOT_ANIM.ListIndex = 0

End Sub
Public Sub Load_NewAnimation()
Dim S As String
Dim i As Long
Dim p As Long
Dim k As Long
Dim GrafCounter As Integer
S = App.PATH & "\RESOURCES\INIT\NewAnim.dat"


Num_NwAnim = Val(GetVar(S, "NW_ANIM", "NUM"))

If Num_NwAnim < 1 Then Exit Sub

ReDim NewAnimationData(1 To Num_NwAnim)

For i = 1 To Num_NwAnim

With NewAnimationData(i)
    .Grafico = Val(GetVar(S, "ANIMACION" & i, "Grafico"))
    .Columnas = Val(GetVar(S, "ANIMACION" & i, "Columnas"))
    .Filas = Val(GetVar(S, "ANIMACION" & i, "Filas"))
    .Height = Val(GetVar(S, "ANIMACION" & i, "Alto"))
    .Width = Val(GetVar(S, "ANIMACION" & i, "Ancho"))
    .NumFrames = Val(GetVar(S, "ANIMACION" & i, "NumeroFrames"))
    .Velocidad = Val(GetVar(S, "ANIMACION" & i, "Velocidad"))
    .TileWidth = .Width / 32
    .TileHeight = .Height / 32
    .Romboidal = Val(GetVar(S, "ANIMACION" & i, "AnimacionRomboidal"))
    ReDim .Indice(1 To .NumFrames) As tNewIndice
    GrafCounter = .Grafico
    k = Val(GetVar(S, "ANIMACION" & i, "INICIAL"))
    If k = 0 Then k = 1
    For p = 1 To .NumFrames
        .Indice(p).X = (((k - 1) Mod .Columnas) * .Width)
        .Indice(p).Y = ((Int((k - 1) / .Columnas)) * .Height)
        .Indice(p).Grafico = GrafCounter
        If p = CInt(.Columnas) * CInt(.Filas) And p < .NumFrames Then
            GrafCounter = GrafCounter + 1
            k = 0
        End If
        k = k + 1
    Next p
End With
Next i


End Sub

Public Sub SPOTLIGHTS_CREAR(ByVal Tipo As Byte, ByVal SPOT_COLOR_BASE As Byte, ByVal SPOT_COLOR_EXTRA, ByVal SPOT_INTENSITY As Byte, ByVal BIND_TO As Byte, ByVal Grafico As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal CHarIndex As Integer, Optional ByVal Color As Long, Optional ByVal COLOR_EXTRA As Long, Optional ByVal EXTRA_GRAFICO As Integer)
On Error GoTo erz
Dim PUEDE_CREAR As Boolean
If MapData(X, Y).SPOTLIGHT.INTENSITY > 0 Then Exit Sub
If SPOT_INTENSITY < 1 Or SPOT_INTENSITY = 2 Or SPOT_INTENSITY = 4 Then Exit Sub
Num_SPOTLIGHTS = Num_SPOTLIGHTS + 1
ReDim Preserve SPOT_LIGHTS(1 To Num_SPOTLIGHTS) As tSPOT_LIGHTS

With SPOT_LIGHTS(Num_SPOTLIGHTS)

    .SPOT_TIPO = Tipo
    
    .BIND_TO = BIND_TO
    
    .Mx = X
    .My = Y
                With MapData(X, Y).SPOTLIGHT
                
                    .INTENSITY = SPOT_INTENSITY
                    .SPOT_TIPO = Tipo
                    .SPOT_COLOR_BASE = SPOT_COLOR_BASE
                    .SPOT_COLOR_EXTRA = SPOT_COLOR_EXTRA
                    .Grafico = Grafico
                    .EXTRA_GRAFICO = EXTRA_GRAFICO
                    .COLOR_EXTRA = COLOR_EXTRA
                    .Color = Color
                    .OffsetX = Val(frmMain.SPOT_OFFSETX.Text)
                    .OffsetY = Val(frmMain.SPOT_OFFSETY.Text)
                    .index = Num_SPOTLIGHTS
                
                
                End With

    .EXTRA_GRAFICO = EXTRA_GRAFICO
                    .OffsetX = Val(frmMain.SPOT_OFFSETX.Text)
                    .OffsetY = Val(frmMain.SPOT_OFFSETY.Text)
    
    
    .SPOT_COLOR_BASE = SPOT_COLOR_BASE
    .INTENSITY = SPOT_INTENSITY
    If SPOT_COLOR_BASE = frmMain.COLORSPOT.ListCount Then
        .Color = Color
    Else
        .Color = SPOTLIGHTS_COLORES(.SPOT_COLOR_BASE)
    End If
    .SPOT_COLOR_EXTRA = SPOT_COLOR_EXTRA
    If SPOT_COLOR_EXTRA > 0 Then
        If SPOT_COLOR_EXTRA = frmMain.COLOREXTRA.ListCount - 1 Then
            .COLOR_EXTRA = COLOR_EXTRA
        Else
            .COLOR_EXTRA = SPOTLIGHTS_COLORES(.SPOT_COLOR_EXTRA)
        End If
    End If
    
    If .SPOT_TIPO > 0 Then
        'Sino es estática iniciamos la animación.
        nwAnimInit .Anim, SPOTLIGHTS_ANIMATION(.SPOT_TIPO)
    Else
        .Grafico = Grafico
    End If
LUZ_SELECTA = Num_SPOTLIGHTS
End With
Exit Sub
erz:
MsgBox "ERROR SPOTLIGHTS_CREAR: " & Err.Description

End Sub
Public Sub SPOTLIGHTS_BORRAR(ByVal Indice As Integer)
Dim i  As Long
On Error GoTo errz
If Indice > UBound(SPOT_LIGHTS) Or Indice = 0 Then Exit Sub


        
MapData(SPOT_LIGHTS(Indice).Mx, SPOT_LIGHTS(Indice).My).SPOTLIGHT.index = 0


If Num_SPOTLIGHTS > Indice Then
    'Tenemos que hacer resize.
    'Resort
    SPOTLIGHTS_RESORT Indice

End If
    
Num_SPOTLIGHTS = Num_SPOTLIGHTS - 1
If Num_SPOTLIGHTS > 0 Then ReDim Preserve SPOT_LIGHTS(1 To Num_SPOTLIGHTS) As tSPOT_LIGHTS
Exit Sub
errz:
MsgBox "ERROR SPOTLIGHTSBORRAR: " & Err.Description
End Sub

Public Sub SPOTLIGHTS_LIMPIARTODOS()

Num_SPOTLIGHTS = 0
Erase SPOT_LIGHTS

Set SCREEN_SPOTS = Nothing

End Sub

Private Sub SPOTLIGHTS_RESORT(ByVal Start As Integer)
On Error GoTo erz
Dim i As Long

For i = Start To Num_SPOTLIGHTS - 1

    SPOT_LIGHTS(i) = SPOT_LIGHTS(i + 1)
    
    Select Case SPOT_LIGHTS(i).BIND_TO
    
        Case 0 'Screen
            'SCREEN_SPOTS.Item(SPOT_LIGHTS(i).INDEX_IN_COL) = i
        
        Case 1 'Mapa
            MapData(SPOT_LIGHTS(i).Mx, SPOT_LIGHTS(i).My).SPOTLIGHT.index = i
        
        Case 2 'Char
            'CharList(SPOT_LIGHTS(i).CHarIndex).SPOTLIGHTS(SPOT_LIGHTS(i).INDEX_IN_COL) = i
    End Select
Next i
Exit Sub
erz:
MsgBox "ERROR SPOTS_RESORT: " & Err.Description
End Sub
Public Sub SPOTLIGHTS_LOADDATA(ByVal FF As Integer)
    'CARGA BINARIA DE SPOTLIGHTS DESDE EFECTOS.BIN
    

End Sub


Private Sub SPOTLIGHTS_DRAW(ByVal SPOT As Integer)
Dim CurrentIndex As Integer
Dim CurrentGrafico As Integer
Dim light_value(0 To 3) As Long
Dim Width As Integer
Dim Height As Integer
Dim sx As Integer
Dim sy As Integer
Dim X As Integer
Dim Y As Integer
Dim d3dTextures As D3D8Textures
Dim z As Integer
Dim verts(3) As TLVERTEX

    With SPOT_LIGHTS(SPOT)
    

        
    X = .X + .OffsetX
    Y = .Y + .OffsetY

    
    'Primero vemos si es animada
    If .SPOT_TIPO > 0 Then
        With SPOT_LIGHTS(SPOT).Anim
        
            If (.Romboidal = 1 And .Direction = 1) Or .Romboidal = 0 Then
                .IndiceCounter = .IndiceCounter + (((GetTickCount - fps_last_time) * 0.0001) * .NumFrames * (.Velocidad * 0.5))
            
                CurrentIndex = .IndiceCounter
                If CurrentIndex >= .NumFrames + 1 Then
                     If .Romboidal = 1 Then
                        .Direction = -1
                        .IndiceCounter = .NumFrames
                        CurrentIndex = .IndiceCounter
                    Else
                        .IndiceCounter = (.IndiceCounter Mod .NumFrames)
                        CurrentIndex = .IndiceCounter
                    End If
                End If
            ElseIf .Romboidal = 1 And .Direction = -1 Then
                .IndiceCounter = .IndiceCounter - (((GetTickCount - fps_last_time) * 0.0001) * .NumFrames * (.Velocidad * 0.5))
                CurrentIndex = .IndiceCounter
                If CurrentIndex <= 0 Then
                
                    .IndiceCounter = 1
                     CurrentIndex = .IndiceCounter
                     .Direction = 1
                End If
            
            
            End If
                If .TileWidth <> 1 Then
                    z = -.TileWidth * 16 + 16
                    X = X + z
                End If
                If .TileHeight <> 1 Then
                    z = -.TileHeight * 16 + 16
                    Y = Y + z
                End If
            If CurrentIndex = 0 Then Exit Sub
            CurrentGrafico = .Indice(CurrentIndex).Grafico
            Width = .Width
            Height = .Height
            sx = .Indice(CurrentIndex).X
            sy = .Indice(CurrentIndex).Y
        End With
    
    Else
        CurrentGrafico = .Grafico
    End If

    Set d3dTextures.Texture = DXPool.GetTexture(CurrentGrafico)
    Call DXPool.Texture_Dimension_Get(CurrentGrafico, d3dTextures.texwidth, d3dTextures.texheight)
    
   ddevice.SetTexture 0, d3dTextures.Texture
    
    If Width = 0 Then
        Width = d3dTextures.texwidth
        Height = d3dTextures.texheight
        
        z = -(Width / 32) * 16 + 19
        X = X + z
        
        z = -(Height / 32) * 16 + 16
        Y = Y + z
    End If
    
    
      If ((d3dTextures.texwidth - 1)) And ((d3dTextures.texheight - 1)) Then
    

            verts(2).X = X
            verts(2).Y = Y + Height
            verts(2).tu = sx / (d3dTextures.texwidth - 1)
            verts(2).tv = (sy + Height) / (d3dTextures.texheight - 1)
            verts(2).rhw = 1
            verts(2).Color = .Color


            verts(0).X = X
            verts(0).Y = Y
            verts(0).tu = sx / (d3dTextures.texwidth - 1)
            verts(0).tv = sy / (d3dTextures.texheight - 1)
            verts(0).rhw = 1
            verts(0).Color = .Color

        

            verts(3).X = X + Width
            verts(3).Y = Y + Height
            verts(3).tu = (sx + Width) / (d3dTextures.texwidth - 1)
            verts(3).tv = (sy + Height) / (d3dTextures.texheight - 1)
            verts(3).rhw = 1
            verts(3).Color = .Color

        

            verts(1).X = X + Width
            verts(1).Y = Y
            verts(1).tu = (sx + Width) / (d3dTextures.texwidth - 1)
            verts(1).tv = sy / (d3dTextures.texheight - 1)
            verts(1).rhw = 1
            verts(1).Color = .Color

    
   
   Else
            verts(0).X = X
            verts(0).Y = Y + Height
            verts(0).tu = 0
            verts(0).tv = 0
            verts(0).rhw = 1
            verts(0).Color = .Color


            verts(1).X = X
            verts(1).Y = Y
            verts(1).tu = 0
            verts(1).tv = 0
            verts(1).rhw = 1
            verts(1).Color = .Color

        

            verts(2).X = X + Width
            verts(2).Y = Y + Height
            verts(2).tu = 0
            verts(2).tv = 0
            verts(2).rhw = 1
            verts(2).Color = .Color

        

            verts(3).X = X + Width
            verts(3).Y = Y
            verts(3).tu = 0
            verts(3).tv = 0
            verts(3).rhw = 1
            verts(3).Color = .Color

    
    End If
        



ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_DESTCOLOR
ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_SRCCOLOR + D3DBLEND_INVSRCCOLOR


   Select Case .INTENSITY
        Case 1
            ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, verts(0), 28
        Case 3
            ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, verts(0), 28
            ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, verts(0), 28
        Case 5
            ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, verts(0), 28
            ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, verts(0), 28
            ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, verts(0), 28
    End Select
   
ddevice.SetRenderState D3DRS_SRCBLEND, 5
ddevice.SetRenderState D3DRS_DESTBLEND, 6

If .COLOR_EXTRA <> 0 Then

    verts(0).Color = .COLOR_EXTRA  '
    verts(1).Color = .COLOR_EXTRA  '
    verts(2).Color = .COLOR_EXTRA  '
    verts(3).Color = .COLOR_EXTRA  '
    
    If .EXTRA_GRAFICO > 0 Then
        Set d3dTextures.Texture = DXPool.GetTexture(.EXTRA_GRAFICO)
        ddevice.SetTexture 0, d3dTextures.Texture
    End If
        
            



ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    ddevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
        Call ddevice.SetTextureStageState(0, D3DTSS_ALPHAARG1, 2)
        Call ddevice.SetTextureStageState(0, D3DTSS_ALPHAARG2, 1)


    ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, verts(0), 28

ddevice.SetRenderState D3DRS_SRCBLEND, 5
ddevice.SetRenderState D3DRS_DESTBLEND, 6


        Call ddevice.SetTextureStageState(0, D3DTSS_ALPHAARG1, D3DTOP_SELECTARG1)
        Call ddevice.SetTextureStageState(0, D3DTSS_ALPHAARG2, D3DTOP_DISABLE)

End If

If Not .BIND_TO = 0 Then .Mustbe_Render = False
End With
End Sub
Public Sub SPOTLIGHTS_RENDER()
Dim i As Long

For i = 1 To Num_SPOTLIGHTS
    If SPOT_LIGHTS(i).Mustbe_Render Then
        SPOTLIGHTS_DRAW i
    End If
Next i
End Sub
Public Sub nwAnimInit(ByRef Animacion As tNewAnimation, ByVal Numero As Integer)
Dim t As Long
If Numero <= Num_NwAnim Then
With NewAnimationData(Numero)
    Animacion.Columnas = .Columnas
    Animacion.Filas = .Filas
    Animacion.Grafico = .Grafico
    Animacion.Width = .Width
    Animacion.Height = .Height
    Animacion.NumFrames = .NumFrames
    Animacion.Velocidad = .Velocidad
    Animacion.TileHeight = .TileHeight
    Animacion.TileWidth = .TileWidth
    Animacion.Romboidal = .Romboidal
    Animacion.Direction = 1
    ReDim Animacion.Indice(1 To .NumFrames) As tNewIndice
    For t = 1 To .NumFrames
        Animacion.Indice(t).X = .Indice(t).X
        Animacion.Indice(t).Y = .Indice(t).Y
        Animacion.Indice(t).Grafico = .Indice(t).Grafico
    Next t
End With
End If
End Sub

