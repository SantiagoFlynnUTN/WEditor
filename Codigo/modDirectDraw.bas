Attribute VB_Name = "modTileEngine"
'**************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'**************************************************************

''
' modTileEngine Nothing to do with DD
'
' @remarks Funciones de DirectDraw y Visualizacion
' @author unkwown
' @version 0.0.20
' @date 20061015

Option Explicit
Public Agua_Tex As Direct3DTexture8
Public Agua_Sur As Direct3DSurface8
Public Back_Sur As Direct3DSurface8
Public Back_St As Direct3DSurface8
Public Stencil As Direct3DSurface8
Public Agua_St As Direct3DSurface8

Public Indice_X(256) As Integer
Public Indice_Y(256) As Integer

Public Enum eTipoTerreno
    Nada = 0
    Agua = 1
    Lava = 2
    Costa = 4
End Enum
Public Type Particle
    friction As Single
    X As Single
    y As Single
    vector_x As Single
    vector_y As Single
    Angle As Byte
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
End Type

 
'Modified by: Ryan Cain (Onezero)
'Last modify date: 5/14/2003
Public Type particle_group
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
   
    'Added by Juan Martín Sotuyo Dodero
    Speed As Single
    life_counter As Long
End Type

'Particle system
 
Public particle_group_list() As particle_group
Public particle_group_count As Long
Public particle_group_last As Long

Public ma(1) As Single

Public Type TLVERTEX
    X As Single
    y As Single
    z As Single
    rhw As Single
    Color As Long
    tu As Single
    tv As Single
End Type

Sub ConvertTPtoCP(StartPixelLeft As Integer, StartPixelTop As Integer, CX As Single, CY As Single, ByVal tX As Integer, ByVal tY As Integer)
'*************************************************
'converts a tile pos to cursor pos
'*************************************************

Dim HWindowX As Integer
Dim HWindowY As Integer

Dim lx As Integer
Dim ly As Integer


HWindowX = (WindowTileWidth \ 2)
HWindowY = (WindowTileHeight \ 2)



Dim iX As Integer
Dim iY As Integer
iX = UserPos.X - HWindowX
lx = tX - iX
CX = (lx * TilePixelWidth) + StartPixelLeft

iY = UserPos.y - HWindowY
ly = tY - iY
CY = (ly * TilePixelHeight) + StartPixelTop


End Sub
Sub ConvertCPtoTP(StartPixelLeft As Integer, StartPixelTop As Integer, ByVal CX As Single, ByVal CY As Single, tX As Integer, tY As Integer)
'*************************************************

'converts cursor pos to tile pos

'*************************************************
Dim HWindowX As Integer
Dim HWindowY As Integer

CX = CX - StartPixelLeft
CY = CY - StartPixelTop

HWindowX = (WindowTileWidth \ 2)
HWindowY = (WindowTileHeight \ 2)

'Figure out X and Y tiles
CX = (CX \ TilePixelWidth)
CY = (CY \ TilePixelHeight)

If CX > HWindowX Then
    CX = (CX - HWindowX)

Else
    If CX < HWindowX Then
        CX = (0 - (HWindowX - CX))
    Else
        CX = 0
    End If
End If

If CY > HWindowY Then
    CY = (0 - (HWindowY - CY))
Else
    If CY < HWindowY Then
        CY = (CY - HWindowY)
    Else
        CY = 0
    End If
End If

tX = UserPos.X + CX
tY = UserPos.y + CY

End Sub

Sub MakeChar(CHarIndex As Integer, Body As Integer, Head As Integer, Heading As Byte, X As Integer, y As Integer)
'*************************************************
'Author: Unkwown
'Last modified: 28/05/06 by GS
'*************************************************
On Error Resume Next

'Update LastChar
If CHarIndex > LastChar Then LastChar = CHarIndex
NumChars = NumChars + 1

'Update head, body, ect.
CharList(CHarIndex).iHead = Head
CharList(CHarIndex).iBody = Body
CharList(CHarIndex).Body(1).index = BodyData(Body).mMovement(1)
CharList(CHarIndex).Body(2).index = BodyData(Body).mMovement(2)
CharList(CHarIndex).Body(3).index = BodyData(Body).mMovement(3)
CharList(CHarIndex).Body(4).index = BodyData(Body).mMovement(4)
If Head > 0 Then
CharList(CHarIndex).Head(1).index = HeadData(Head).Frame(1)
CharList(CHarIndex).Head(2).index = HeadData(Head).Frame(2)
CharList(CHarIndex).Head(3).index = HeadData(Head).Frame(3)
CharList(CHarIndex).Head(4).index = HeadData(Head).Frame(4)
Else
CharList(CHarIndex).Head(1).index = 0
CharList(CHarIndex).Head(2).index = 0
CharList(CHarIndex).Head(3).index = 0
CharList(CHarIndex).Head(4).index = 0
End If
If Heading = 0 Then Heading = 3
CharList(CHarIndex).Heading = Heading

'Reset moving stats
CharList(CHarIndex).Moving = 0
CharList(CHarIndex).MoveOffset.X = 0
CharList(CHarIndex).MoveOffset.y = 0

'Update position
CharList(CHarIndex).Pos.X = X
CharList(CHarIndex).Pos.y = y

'Make active
CharList(CHarIndex).Active = 1

'Plot on map
MapData(X, y).CHarIndex = CHarIndex

bRefreshRadar = True ' GS

End Sub

Sub EraseChar(CHarIndex As Integer)
'*************************************************
'Author: Unkwown
'Last modified: 28/05/06 by GS
'*************************************************
If CHarIndex = 0 Then Exit Sub
'Make un-active
CharList(CHarIndex).Active = 0

'Update lastchar
If CHarIndex = LastChar Then
    Do Until CharList(LastChar).Active = 1
        LastChar = LastChar - 1
        If LastChar = 0 Then Exit Do
    Loop
End If

MapData(CharList(CHarIndex).Pos.X, CharList(CHarIndex).Pos.y).CHarIndex = 0

'Update NumChars
NumChars = NumChars - 1

bRefreshRadar = True ' GS

End Sub

Sub MoveCharbyPos(CHarIndex As Integer, nX As Integer, nY As Integer)
'*************************************************
'Author: Unkwown
'Last modified: 28/05/06 by GS
'*************************************************
Dim X As Integer
Dim y As Integer
Dim addx As Integer
Dim addy As Integer
Dim nHeading As Byte

X = CharList(CHarIndex).Pos.X
y = CharList(CHarIndex).Pos.y

addx = nX - X
addy = nY - y

If Sgn(addx) = 1 Then
    nHeading = EAST
End If

If Sgn(addx) = -1 Then
    nHeading = WEST
End If

If Sgn(addy) = -1 Then
    nHeading = NORTH
End If

If Sgn(addy) = 1 Then
    nHeading = SOUTH
End If

MapData(nX, nY).CHarIndex = CHarIndex
CharList(CHarIndex).Pos.X = nX
CharList(CHarIndex).Pos.y = nY
MapData(X, y).CHarIndex = 0

CharList(CHarIndex).MoveOffset.X = -1 * (TilePixelWidth * addx)
CharList(CHarIndex).MoveOffset.y = -1 * (TilePixelHeight * addy)

CharList(CHarIndex).Moving = 1
CharList(CHarIndex).Heading = nHeading

bRefreshRadar = True ' GS

End Sub

Function NextOpenChar() As Integer
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************
Dim LoopC As Integer

LoopC = 1
Do While CharList(LoopC).Active
    LoopC = LoopC + 1
Loop

NextOpenChar = LoopC

End Function

Function LegalPos(X As Integer, y As Integer) As Boolean
'*************************************************
'Author: Unkwown
'Last modified: 28/05/06 - GS
'*************************************************

LegalPos = True

'Check to see if its out of bounds
If X - 8 < 1 Or X + 8 > 100 Or y - 6 < 1 Or y + 6 > 100 Then
    LegalPos = False
    Exit Function
End If

'Check to see if its blocked
If MapData(X, y).Blocked = 1 Then
    LegalPos = False
    Exit Function
End If

'Check for character
If MapData(X, y).CHarIndex > 0 Then
    LegalPos = False
    Exit Function
End If

End Function

Function InMapLegalBounds(X As Integer, y As Integer) As Boolean
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************

If X < MinXBorder Or X > MaxXBorder Or y < MinYBorder Or y > MaxYBorder Then
    InMapLegalBounds = False
    Exit Function
End If

InMapLegalBounds = True

End Function

Function InMapBounds(ByVal X As Long, ByVal y As Long) As Boolean
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************

If X < XMinMapSize Or X > XMaxMapSize Or y < YMinMapSize Or y > YMaxMapSize Then
    InMapBounds = False
    Exit Function
End If

InMapBounds = True

End Function

' [Loopzer]
Public Sub DePegar()
'*************************************************
'Author: Loopzer
'Last modified: 21/11/07
'*************************************************
    Dim X As Integer
    Dim y As Integer

    For X = 0 To DeSeleccionAncho - 1
        For y = 0 To DeSeleccionAlto - 1
             MapData(X + DeSeleccionOX, y + DeSeleccionOY) = DeSeleccionMap(X, y)
        Next
    Next
End Sub
Public Sub PegarSeleccion() '(mx As Integer, my As Integer)
'*************************************************
'Author: Loopzer
'Last modified: 21/11/07
'*************************************************
    'podria usar copy mem , pero por las dudas no XD
    Static UltimoX As Integer
    Static UltimoY As Integer
    If UltimoX = SobreX And UltimoY = SobreY Then Exit Sub
    UltimoX = SobreX
    UltimoY = SobreY
    Dim X As Integer
    Dim y As Integer
    DeSeleccionAncho = SeleccionAncho
    DeSeleccionAlto = SeleccionAlto
    DeSeleccionOX = SobreX
    DeSeleccionOY = SobreY
    ReDim DeSeleccionMap(DeSeleccionAncho, DeSeleccionAlto) As MapBlock
    
    If Mx + UBound(DeSeleccionMap, 1) > 100 Then
        DeSeleccionAncho = 100 - Mx + 1
    End If
    If My + UBound(DeSeleccionMap, 2) > 100 Then
        DeSeleccionAlto = 100 - My + 1
    End If
    
    For X = 0 To DeSeleccionAncho - 1
        For y = 0 To DeSeleccionAlto - 1
            DeSeleccionMap(X, y) = MapData(X + SobreX, y + SobreY)
        Next
    Next
    For X = 0 To DeSeleccionAncho - 1
        For y = 0 To DeSeleccionAlto - 1
             MapData(X + SobreX, y + SobreY) = SeleccionMap(X, y)
        Next
    Next
    Seleccionando = False
End Sub
Public Sub AccionSeleccion()
'*************************************************
'Author: Loopzer
'Last modified: 21/11/07
'*************************************************
    Dim X As Integer
    Dim y As Integer
    SeleccionAncho = Abs(SeleccionIX - SeleccionFX) + 1
    SeleccionAlto = Abs(SeleccionIY - SeleccionFY) + 1
    DeSeleccionAncho = SeleccionAncho
    DeSeleccionAlto = SeleccionAlto
    DeSeleccionOX = SeleccionIX
    DeSeleccionOY = SeleccionIY
    ReDim DeSeleccionMap(DeSeleccionAncho, DeSeleccionAlto) As MapBlock
    
    For X = 0 To SeleccionAncho - 1
        For y = 0 To SeleccionAlto - 1
            DeSeleccionMap(X, y) = MapData(X + SeleccionIX, y + SeleccionIY)
        Next
    Next
    For X = 0 To SeleccionAncho - 1
        For y = 0 To SeleccionAlto - 1
           ClickEdit vbLeftButton, SeleccionIX + X, SeleccionIY + y
        Next
    Next
    Seleccionando = False
    UsingUndoSelection = False
    
End Sub

Public Sub BlockearSeleccion()
'*************************************************
'Author: Loopzer
'Last modified: 21/11/07
'*************************************************
    Dim X As Integer
    Dim y As Integer
    SeleccionAncho = Abs(SeleccionIX - SeleccionFX) + 1
    SeleccionAlto = Abs(SeleccionIY - SeleccionFY) + 1
    DeSeleccionAncho = SeleccionAncho
    DeSeleccionAlto = SeleccionAlto
    DeSeleccionOX = SeleccionIX
    DeSeleccionOY = SeleccionIY
    ReDim DeSeleccionMap(DeSeleccionAncho, DeSeleccionAlto) As MapBlock
    
    For X = 0 To SeleccionAncho - 1
        For y = 0 To SeleccionAlto - 1
            DeSeleccionMap(X, y) = MapData(X + SeleccionIX, y + SeleccionIY)
        Next
    Next
    For X = 0 To SeleccionAncho - 1
        For y = 0 To SeleccionAlto - 1
             If MapData(X + SeleccionIX, y + SeleccionIY).Blocked = 1 Then
                MapData(X + SeleccionIX, y + SeleccionIY).Blocked = 0
             Else
                MapData(X + SeleccionIX, y + SeleccionIY).Blocked = 1
            End If
        Next
    Next
    Seleccionando = False
End Sub
Public Sub CortarSeleccion()
'*************************************************
'Author: Loopzer
'Last modified: 21/11/07
'*************************************************
    CopiarSeleccion
    Dim X As Integer
    Dim y As Integer
    Dim Vacio As MapBlock
    DeSeleccionAncho = SeleccionAncho
    DeSeleccionAlto = SeleccionAlto
    DeSeleccionOX = SeleccionIX
    DeSeleccionOY = SeleccionIY
    ReDim DeSeleccionMap(DeSeleccionAncho, DeSeleccionAlto) As MapBlock
    
    For X = 0 To SeleccionAncho - 1
        For y = 0 To SeleccionAlto - 1
            DeSeleccionMap(X, y) = MapData(X + SeleccionIX, y + SeleccionIY)
        Next
    Next
    For X = 0 To SeleccionAncho - 1
        For y = 0 To SeleccionAlto - 1
             MapData(X + SeleccionIX, y + SeleccionIY) = Vacio
        Next
    Next
    Seleccionando = False
End Sub
Public Sub CopiarSeleccionSinTraslados()
'*************************************************
'Author: Loopzer
'Last modified: 21/11/07
'*************************************************
    'podria usar copy mem , pero por las dudas no XD
    Dim X As Integer
    Dim y As Integer
    Seleccionando = False
    SeleccionAncho = Abs(SeleccionIX - SeleccionFX) + 1
    SeleccionAlto = Abs(SeleccionIY - SeleccionFY) + 1
    If SeleccionIX = 0 Or SeleccionFX = 0 Or SeleccionIY = 0 Or SeleccionFY = 0 Then Exit Sub
    
    ReDim SeleccionMap(SeleccionAncho, SeleccionAlto) As MapBlock
    For X = 0 To SeleccionAncho - 1
        For y = 0 To SeleccionAlto - 1
            SeleccionMap(X, y) = MapData(X + SeleccionIX, y + SeleccionIY)
            SeleccionMap(X, y).TileExit.Map = 0
            SeleccionMap(X, y).TileExit.X = 0
            SeleccionMap(X, y).TileExit.y = 0
        Next
    Next
End Sub
Public Sub CopiarSeleccion()
'*************************************************
'Author: Loopzer
'Last modified: 21/11/07
'*************************************************
    'podria usar copy mem , pero por las dudas no XD
    Dim X As Integer
    Dim y As Integer
    Seleccionando = False
    SeleccionAncho = Abs(SeleccionIX - SeleccionFX) + 1
    SeleccionAlto = Abs(SeleccionIY - SeleccionFY) + 1
    If SeleccionIX = 0 Or SeleccionFX = 0 Or SeleccionIY = 0 Or SeleccionFY = 0 Then Exit Sub
    
    ReDim SeleccionMap(SeleccionAncho, SeleccionAlto) As MapBlock
    For X = 0 To SeleccionAncho - 1
        For y = 0 To SeleccionAlto - 1
            SeleccionMap(X, y) = MapData(X + SeleccionIX, y + SeleccionIY)
        Next
    Next
End Sub
Public Sub GenerarVista()
'*************************************************
'Author: Loopzer
'Last modified: 21/11/07
'*************************************************
   ' hacer una llamada a un seter o geter , es mas lento q una variable
   ' con esto hacemos q no este preguntando a el objeto cadavez
   ' q dibuja , Render mas rapido ;)
    VerBlockeados = frmMain.cVerBloqueos.value
    VerTriggers = frmMain.cVerTriggers.value
    VerCapa1 = frmMain.mnuVerCapa1.Checked
    VerCapa2 = frmMain.mnuVerCapa2.Checked
    VerCapa3 = frmMain.mnuVerCapa3.Checked
    VerCapa4 = frmMain.mnuVerCapa4.Checked
    VerCapa5 = frmMain.MnuVerCapa5.Checked
    VerCapa9 = frmMain.MnuVerCapa9.Checked
    VerTranslados = frmMain.mnuVerTranslados.Checked
    VerObjetos = frmMain.mnuVerObjetos.Checked
    VerNpcs = frmMain.mnuVerNPCs.Checked
    VerDecors = frmMain.mVerDecors.Checked
    
End Sub
' [/Loopzer]
Public Sub RenderScreen(TileX As Integer, TileY As Integer, PixelOffsetX As Integer, PixelOffsetY As Integer)
      '*************************************************
      'Author: Unkwown
      'Last modified: 31/05/06 by GS
      'Last modified: 21/11/07 By Loopzer
      'Last modifier: 24/11/08 by GS
      '*************************************************

10    On Error GoTo errs
      Dim y       As Integer              'Keeps track of where on map we are
      Dim X       As Integer
      Dim MinY    As Integer              'Start Y pos on current map
      Dim MaxY    As Integer              'End Y pos on current map
      Dim MinX    As Integer              'Start X pos on current map
      Dim MaxX    As Integer              'End X pos on current map
      Dim ScreenX As Integer              'Keeps track of where to place tile on screen
      Dim ScreenY As Integer
      Dim Sobre   As Integer
      Dim iPPx    As Integer              'Usado en el Layer de Chars
      Dim iPPy    As Integer              'Usado en el Layer de Chars
      Dim Grh     As Grh
      Dim nGrh As tnGrh 'Temp Grh for show tile and blocked
      Dim bCapa    As Byte                 'cCapas ' 31/05/2006 - GS, control de Capas
      Dim iGrhIndex           As Integer  'Usado en el Layer 1
      Dim PixelOffsetXTemp    As Integer  'For centering grhs
      Dim PixelOffsetYTemp    As Integer
      Dim TempChar            As Char
      Dim tiempo As Byte
      Dim colorlist(3) As Long
      Dim Polygon_Ignore_Top As Byte
      Dim Polygon_Ignore_Right As Byte
      Dim Polygon_Ignore_Left As Byte
      Dim Polygon_Ignore_lower As Byte
      Dim Corner As Byte

20    tiempo = 255
30    colorlist(0) = D3DColorXRGB(255, 200, 0)
40    colorlist(1) = D3DColorXRGB(255, 200, 0)
50    colorlist(2) = D3DColorXRGB(255, 200, 0)
60    colorlist(3) = D3DColorXRGB(255, 200, 0)

70    Map_LightsRender
80    If Not guardobmp Then
90    MinY = (TileY - (WindowTileHeight \ 2)) - TileBufferSize
100   MaxY = (TileY + (WindowTileHeight \ 2)) + TileBufferSize
110   MinX = (TileX - (WindowTileWidth \ 2)) - TileBufferSize
120   MaxX = (TileX + (WindowTileWidth \ 2)) + TileBufferSize

130   Else
140   MinY = TileY - 8
150   MaxY = TileY + 16
160   MinX = TileX - 8
170   MaxX = TileX + 16


180   End If




      ' 31/05/2006 - GS, control de Capas
190   If cCapaSel >= 1 And cCapaSel <= 5 Then
200       bCapa = cCapaSel
210   Else
220       bCapa = 1
230   End If
240   GenerarVista 'Loopzer
250   ScreenY = -8
260   tiempo = 254

          Dim jx As Integer
          Dim jy As Integer
          Dim jh As Integer
          Dim jw As Integer
          Dim jg As Integer
          Dim jtw As Single
          Dim jth As Single

          Dim VertexArray(0 To 3) As TLVERTEX
          Dim Tex As Direct3DTexture8
          Dim SrcWidth As Integer
          Dim Width As Integer
          Dim SrcHeight As Integer
          Dim Height As Integer
          Dim SrcBitmapWidth As Long
          Dim SrcBitmapHeight As Long
          Dim xb As Integer
          Dim yb As Integer
          'Dim iGrhIndex As Integer
          Dim srdesc As D3DSURFACE_DESC
                                  Dim aux As Integer
                              Dim dy As Integer
                              Dim dX As Integer
          
270   For y = (MinY) To (MaxY)
280       ScreenX = -8
290       For X = (MinX) To (MaxX)

300           If InMapBounds(X, y) Then

310               If VerCapa1 Then


320               If MapData(X, y).Graphic(1).index <> 0 And VerCapa1 Then
330                   If MapData(X, y).Luz >= 202 And MapData(X, y).Luz <= 217 Then
340                   modGrh.Grh_iRenderN MapData(X, y).Graphic(1), ((32 * ScreenX) - 32) + PixelOffsetX, ((32 * ScreenY) - 32) + PixelOffsetY, MapData(X, y).light_value, True
                     
350                   Else
360                   modGrh.Grh_RenderN MapData(X, y).Graphic(1), ((32 * ScreenX) - 32) + PixelOffsetX, ((32 * ScreenY) - 32) + PixelOffsetY, MapData(X, y).light_value, True
370                   End If
380               End If
                  
390           End If
400           End If
              
410           ScreenX = ScreenX + 1
420       Next X
430       ScreenY = ScreenY + 1
440       If y > 100 Then Exit For
450   Next y
460   ScreenY = -8


470   ddevice.SetRenderTarget Agua_Sur, Agua_St, ByVal 0
480   ddevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, vbBlack, ByVal 0, ByVal 0
      'Call ddevice.SetTextureStageState(0, D3DTSS_ALPHAOP, D3DTOP_MODULATE)
490   Call ddevice.SetTextureStageState(0, D3DTSS_ALPHAARG1, D3DTA_TEXTURE)
500   Call ddevice.SetTextureStageState(0, D3DTSS_ALPHAARG2, D3DTA_CURRENT)
510   ddevice.SetRenderState D3DRS_SRCBLEND, 5
520   ddevice.SetRenderState D3DRS_DESTBLEND, 1

530   For y = (MinY) To (MaxY)
540       ScreenX = -8
550       For X = (MinX) To (MaxX)

560           If InMapBounds(X, y) Then
570             If MapData(X, y).Graphic(2).index > 0 And VerCapa2 Then

580       xb = (ScreenX - 1) * 32 + PixelOffsetX
590       yb = (ScreenY - 1) * 32 + PixelOffsetY
         
600       If NewIndexData(MapData(X, y).Graphic(2).index).Dinamica > 0 Then
610           With NewAnimationData(NewIndexData(MapData(X, y).Graphic(2).index).Dinamica)
              
             
                  'MapData(X, y).Graphic(2).fC = MapData(X, y).Graphic(2).fC + ((timer_elapsed_time * 0.1) * .NumFrames / .Velocidad)
620               If Not MapData(X, y).TipoTerreno And eTipoTerreno.Agua Then
630               MapData(X, y).Graphic(2).fC = MapData(X, y).Graphic(2).fC + (.NumFrames * (MEE * 0.0011) * Rnd)
640               Else
650               MapData(X, y).Graphic(2).fC = MapData(X, y).Graphic(2).fC + (.NumFrames * (MEE * 0.0005) * Rnd)
                  
660               End If
670               If MapData(X, y).Graphic(2).fC > .NumFrames Then
680                   MapData(X, y).Graphic(2).fC = (MapData(X, y).Graphic(2).fC Mod .NumFrames) + 1
690               End If
700   tiempo = 1
710               If MapData(X, y).Graphic(2).fC < 1 Then MapData(X, y).Graphic(2).fC = 1
                  
720               jx = .Indice(MapData(X, y).Graphic(2).fC).X
730               jy = .Indice(MapData(X, y).Graphic(2).fC).y
740               jw = .Width
750               jh = .Height
760               jtw = .TileWidth
770               jth = .TileHeight
780               jg = (.Indice(MapData(X, y).Graphic(2).fC).Grafico - .Indice(2).Grafico) + NewIndexData(MapData(X, y).Graphic(2).index).OverWriteGrafico
790           End With
800       Else
810           With EstaticData(NewIndexData(MapData(X, y).Graphic(2).index).Estatic)
820               jx = .L
830               jy = .t
840               jw = .W
850               jh = .H
860               jth = .th
870               jtw = .tw
880               jg = NewIndexData(MapData(X, y).Graphic(2).index).OverWriteGrafico
          
890           End With
900       End If
                      
910       Set Tex = DXPool.GetTexture(jg)
          'Call DXPool.Texture_Dimension_Get(.texture_index, texture_width, texture_height)
          
920       Tex.GetLevelDesc 0, srdesc
          
930           SrcWidth = 32 'd3dtextures.texwidth
940           Width = 32 'd3dtextures.texwidth
             
950           Height = 32 'd3dtextures.texheight
960           SrcHeight = 32 'd3dtextures.texheight
970           SrcBitmapWidth = srdesc.Width
980           SrcBitmapHeight = srdesc.Height
          'Set the RHWs (must always be 1)
         
990       VertexArray(0).rhw = 1
1000      VertexArray(1).rhw = 1
1010      VertexArray(2).rhw = 1
1020      VertexArray(3).rhw = 1
              
1030          If MapData(X, y).Luz <= 201 Or MapData(X, y).Luz >= 218 Then
              
              
              'Find the left side of the rectangle
1040          VertexArray(0).X = xb
1050          VertexArray(0).tu = (jx / SrcBitmapWidth)
       
              'Find the top side of the rectangle
1060          VertexArray(0).y = yb
1070          VertexArray(0).tv = (jy / SrcBitmapHeight)
         
              'Find the right side of the rectangle
1080          VertexArray(1).X = xb + jw
1090          VertexArray(1).tu = (jx + jw) / SrcBitmapWidth
       
              'These values will only equal each other when not a shadow
1100          VertexArray(2).X = VertexArray(0).X
1110          VertexArray(3).X = VertexArray(1).X
       
          'Find the bottom of the rectangle
1120      VertexArray(2).y = yb + jh
1130      VertexArray(2).tv = (jy + jh) / SrcBitmapHeight
       
          'Because this is a perfect rectangle, all of the values below will equal one of the values we already got
1140      VertexArray(1).y = VertexArray(0).y
1150      VertexArray(1).tv = VertexArray(0).tv
1160      VertexArray(2).tu = VertexArray(0).tu
1170      VertexArray(3).y = VertexArray(2).y
1180      VertexArray(3).tu = VertexArray(1).tu
1190      VertexArray(3).tv = VertexArray(2).tv
1200                              If ((MapData(X, y).TipoTerreno And eTipoTerreno.Agua) Or (MapData(X, y).TipoTerreno And eTipoTerreno.Lava)) Then

             
1210                              Polygon_Ignore_Right = 0
1220                              Polygon_Ignore_Left = 0
1230                              Polygon_Ignore_Top = 0
1240                              Polygon_Ignore_lower = 0
1250                              Corner = 0
                                  
1260                              If y <> 1 Then
1270                                If Not MapData(X, y - 1).TipoTerreno And eTipoTerreno.Agua And Not (MapData(X, y - 1).TipoTerreno And eTipoTerreno.Lava) Then Polygon_Ignore_Top = 1
1280                              End If
                                  
1290                              If y <> 100 Then
1300                                If Not MapData(X, y + 1).TipoTerreno And eTipoTerreno.Agua And Not (MapData(X, y + 1).TipoTerreno And eTipoTerreno.Lava) Then Polygon_Ignore_lower = 1
1310                              End If
                                  
1320                              If X <> 100 Then
1330                                If Not MapData(X + 1, y).TipoTerreno And eTipoTerreno.Agua And Not (MapData(X + 1, y).TipoTerreno And eTipoTerreno.Lava) Then Polygon_Ignore_Right = 1
1340                              End If
                                  
1350                              If X <> 1 Then
1360                                If Not MapData(X - 1, y).TipoTerreno And eTipoTerreno.Agua And Not (MapData(X - 1, y).TipoTerreno And eTipoTerreno.Lava) Then Polygon_Ignore_Left = 1
1370                              End If
                                  
1380                            If Polygon_Ignore_Left = 0 Then
1390                                  If X > 1 And y > 1 Then
1400                                  If MapData(X, y - 1).TipoTerreno And eTipoTerreno.Agua And (Not MapData(X - 1, y - 1).TipoTerreno And eTipoTerreno.Agua) Then
1410                                      Corner = 2
1420                                  End If
1430                                  End If
1440                                  If X > 1 And y < 100 Then
1450                                  If MapData(X, y + 1).TipoTerreno And eTipoTerreno.Agua And Not (MapData(X - 1, y + 1).TipoTerreno And eTipoTerreno.Agua) Then
1460                                      Corner = 1
1470                                  End If
1480                                  End If
1490                              End If
1500                              If Polygon_Ignore_Right = 0 Then
1510                                  If X < 100 And y > 1 Then
1520                                  If MapData(X, y - 1).TipoTerreno And eTipoTerreno.Agua And Not (MapData(X + 1, y - 1).TipoTerreno And eTipoTerreno.Agua) Then
1530                                      Corner = 4
1540                                  End If
1550                                  End If
1560                                  If X < 100 And y < 100 Then
1570                                  If MapData(X, y + 1).TipoTerreno And eTipoTerreno.Agua And Not (MapData(X + 1, y + 1).TipoTerreno And eTipoTerreno.Agua) Then
1580                                      Corner = 3
1590                                  End If
1600                                  End If
1610                              End If
                                  


                    
                                  
                                  
1620                              VertexArray(1).X = VertexArray(1).X + PolygonX
1630                              VertexArray(3).X = VertexArray(3).X + PolygonX


1640                              If Polygon_Ignore_Top <> 1 Then
1650                                  VertexArray(0).y = VertexArray(0).y + polygonCount(1)
1660                                  VertexArray(1).y = VertexArray(1).y - polygonCount(1)
1670                              End If

1680                              If Polygon_Ignore_lower <> 1 Then
1690                                  VertexArray(2).y = VertexArray(2).y + polygonCount(1)
1700                                  VertexArray(3).y = VertexArray(3).y - polygonCount(1)
1710                              End If
                                  
                                  
                               


                  
1720                      End If
         
1730      If MapData(X, y).light_value(0) <> 0 Then
1740          VertexArray(0).Color = MapData(X, y).light_value(0)
1750      Else
1760          VertexArray(0).Color = base_light
1770      End If
1780        If MapData(X, y).light_value(1) <> 0 Then
1790          VertexArray(1).Color = MapData(X, y).light_value(1)
1800      Else
1810          VertexArray(1).Color = base_light
1820      End If
1830      If MapData(X, y).light_value(2) <> 0 Then
1840          VertexArray(2).Color = MapData(X, y).light_value(2)
1850      Else
1860          VertexArray(2).Color = base_light
1870      End If
1880      If MapData(X, y).light_value(3) <> 0 Then
1890          VertexArray(3).Color = MapData(X, y).light_value(3)
1900      Else
1910          VertexArray(3).Color = base_light
1920      End If
         
1930     Else
         
              'Find the left side of the rectangle
1940          VertexArray(1).X = xb
1950          VertexArray(1).tu = (jx / SrcBitmapWidth)
       
              'Find the top side of the rectangle
1960          VertexArray(1).y = yb
1970          VertexArray(1).tv = (jy / SrcBitmapHeight)
         
              'Find the right side of the rectangle
1980          VertexArray(3).X = xb + jw
1990          VertexArray(3).tu = (jx + jw) / SrcBitmapWidth
       
              'These values will only equal each other when not a shadow
2000          VertexArray(0).X = VertexArray(1).X
2010          VertexArray(2).X = VertexArray(3).X
       
          'Find the bottom of the rectangle
2020      VertexArray(0).y = yb + jh
2030      VertexArray(0).tv = (jy + jh) / SrcBitmapHeight
       
          'Because this is a perfect rectangle, all of the values below will equal one of the values we already got
2040      VertexArray(3).y = VertexArray(1).y
2050      VertexArray(3).tv = VertexArray(1).tv
2060      VertexArray(0).tu = VertexArray(1).tu
2070      VertexArray(2).y = VertexArray(0).y
2080      VertexArray(2).tu = VertexArray(3).tu
2090      VertexArray(2).tv = VertexArray(0).tv
         
         
2100                             If (MapData(X, y).TipoTerreno And eTipoTerreno.Agua Or MapData(X, y).TipoTerreno And eTipoTerreno.Lava) Then

             
2110                              Polygon_Ignore_Right = 0
2120                              Polygon_Ignore_Left = 0
2130                              Polygon_Ignore_Top = 0
2140                              Polygon_Ignore_lower = 0
2150                              Corner = 0
                                  
2160                              If y <> 1 Then
2170                                If Not MapData(X, y - 1).TipoTerreno And eTipoTerreno.Agua And Not (MapData(X, y - 1).TipoTerreno And eTipoTerreno.Lava) Then Polygon_Ignore_Top = 1
2180                              End If
                                  
2190                              If y <> 100 Then
2200                                If Not MapData(X, y + 1).TipoTerreno And eTipoTerreno.Agua And Not (MapData(X, y + 1).TipoTerreno And eTipoTerreno.Lava) Then Polygon_Ignore_lower = 1
2210                              End If
                                  
2220                              If X <> 100 Then
2230                                If Not MapData(X + 1, y).TipoTerreno And eTipoTerreno.Agua And Not (MapData(X + 1, y).TipoTerreno And eTipoTerreno.Lava) Then Polygon_Ignore_Right = 1
2240                              End If
                                  
2250                              If X <> 1 Then
2260                                If Not MapData(X - 1, y).TipoTerreno And eTipoTerreno.Agua And Not (MapData(X - 1, y).TipoTerreno And eTipoTerreno.Lava) Then Polygon_Ignore_Left = 1
2270                              End If
                                  
2280                            If Polygon_Ignore_Left = 0 Then
2290                                  If X > 1 And y > 1 Then
2300                                  If Not MapData(X, y - 1).TipoTerreno And eTipoTerreno.Agua And MapData(X - 1, y - 1).TipoTerreno And eTipoTerreno.Agua Then
2310                                      Polygon_Ignore_Left = 1
2320                                      Corner = 1
2330                                  End If
2340                                  End If
2350                                  If X > 1 And y < 100 Then
2360                                  If Not MapData(X, y + 1).TipoTerreno And eTipoTerreno.Agua And MapData(X - 1, y + 1).TipoTerreno And eTipoTerreno.Agua Then
2370                                      Polygon_Ignore_Left = 1
2380                                      Corner = 1
2390                                  End If
2400                                  End If
2410                              End If
2420                              If Polygon_Ignore_Right = 0 Then
2430                                  If X < 100 And y > 1 Then
2440                                  If Not MapData(X, y - 1).TipoTerreno And eTipoTerreno.Agua And MapData(X + 1, y - 1).TipoTerreno And eTipoTerreno.Agua Then
2450                                      Polygon_Ignore_Right = 1
2460                                      Corner = 1
2470                                  End If
2480                                  End If
2490                                  If X < 100 And y < 100 Then
2500                                  If Not MapData(X, y + 1).TipoTerreno And eTipoTerreno.Agua And MapData(X + 1, y + 1).TipoTerreno And eTipoTerreno.Agua Then
2510                                      Polygon_Ignore_Right = 1
2520                                      Corner = 1
2530                                  End If
2540                                  End If
2550                              End If
                                  

2560                              If Polygon_Ignore_Left <> 1 Then
2570                                  VertexArray(1).X = VertexArray(1).X + PolygonX
2580                                  VertexArray(0).X = VertexArray(0).X + PolygonX
2590                              End If
                                  

2600                                  If Corner = 1 Then
2610                                      VertexArray(1).y = VertexArray(1).y - 1
2620                                      VertexArray(3).y = VertexArray(3).y - 1
2630                                      VertexArray(0).y = VertexArray(0).y + 1
2640                                      VertexArray(2).y = VertexArray(2).y + 1
2650                                  End If

                                  
2660                          If Polygon_Ignore_Right <> 1 Then
2670                              VertexArray(3).X = VertexArray(3).X + PolygonX
2680                              VertexArray(2).X = VertexArray(2).X + PolygonX
2690                          End If
                              
2700                              If Polygon_Ignore_Top <> 1 Then
2710                                  VertexArray(3).y = VertexArray(3).y - polygonCount(1)
2720                                  VertexArray(1).y = VertexArray(1).y + polygonCount(1)
                                  
2730                              End If

2740                              If Polygon_Ignore_lower <> 1 Then
2750                                  VertexArray(2).y = VertexArray(2).y - polygonCount(1)
2760                                  VertexArray(0).y = VertexArray(0).y + polygonCount(1)
2770                              End If
                          
                                

                             


2780              End If
          
2790      If MapData(X, y).light_value(0) <> 0 Then
2800          VertexArray(0).Color = MapData(X, y).light_value(0)
2810      Else
2820          VertexArray(0).Color = base_light
2830      End If
2840        If MapData(X, y).light_value(1) <> 0 Then
2850          VertexArray(1).Color = MapData(X, y).light_value(1)
2860      Else
2870          VertexArray(1).Color = base_light
2880      End If
2890      If MapData(X, y).light_value(2) <> 0 Then
2900          VertexArray(2).Color = MapData(X, y).light_value(2)
2910      Else
2920          VertexArray(2).Color = base_light
2930      End If
2940      If MapData(X, y).light_value(3) <> 0 Then
2950          VertexArray(3).Color = MapData(X, y).light_value(3)
2960      Else
2970          VertexArray(3).Color = base_light
2980      End If
         
2990     End If


          
          'VertexArray(0).y = VertexArray(0).y - MapData(X, y).AlturaPoligonos(0)
          'VertexArray(1).y = VertexArray(1).y - MapData(X, y).AlturaPoligonos(1)
          'VertexArray(2).y = VertexArray(2).y - MapData(X, y).AlturaPoligonos(2)
          'VertexArray(3).y = VertexArray(3).y - MapData(X, y).AlturaPoligonos(3)
          
3000      ddevice.SetTexture 0, Tex
          


3010      ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexArray(0), 28
          

          

          
3020  End If


                  'Layer 2 **********************************
3030  End If
3040          ScreenX = ScreenX + 1
3050      Next X
3060      ScreenY = ScreenY + 1
3070      If y > 100 Then Exit For
3080  Next y

3090  ddevice.SetRenderTarget Back_Sur, Stencil, ByVal 0

3100  ddevice.SetTexture 0, Agua_Tex

3110  VertexArray(0).X = 0
3120  VertexArray(0).y = 0
3130  VertexArray(0).tu = 0
3140  VertexArray(0).tv = 0 '

3150  VertexArray(1).X = ClienteWidth * 32
3160  VertexArray(1).y = 0
3170  VertexArray(1).tu = 1
3180  VertexArray(1).tv = 0

3190  VertexArray(2).X = 0
3200  VertexArray(2).y = ClienteHeight * 32
3210  VertexArray(2).tu = 0
3220  VertexArray(2).tv = 1

3230  VertexArray(3).X = ClienteWidth * 32
3240  VertexArray(3).y = ClienteHeight * 32
3250  VertexArray(3).tu = 1
3260  VertexArray(3).tv = 1
      '
3270  ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA

3280  ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA

3290  Call ddevice.SetTextureStageState(0, D3DTSS_ALPHAARG1, D3DTA_TEXTURE)
3300  Call ddevice.SetTextureStageState(0, D3DTSS_ALPHAARG2, D3DTA_CURRENT)



3310  ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexArray(0), 28

3320  ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
3330  ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA

3340  Call ddevice.SetTextureStageState(0, D3DTSS_ALPHAARG1, D3DTA_TEXTURE Or D3DTA_DIFFUSE)
3350  Call ddevice.SetTextureStageState(0, D3DTSS_ALPHAARG2, D3DTA_CURRENT)


3360            tiempo = 2
3370              ScreenY = -8
3380  For y = (MinY) To (MaxY)
3390      ScreenX = -8
3400      For X = (MinX) To (MaxX)

3410          If InMapBounds(X, y) Then
                  

                  
                  'Layer 5
3420              If MapData(X, y).Graphic(5).index <> 0 And VerCapa5 Then
3430                  If MapData(X, y).Luz >= 202 And MapData(X, y).Luz <= 217 Then
3440                  modGrh.Grh_iRenderN MapData(X, y).Graphic(5), ((32 * ScreenX) - 32) + PixelOffsetX, ((32 * ScreenY) - 32) + PixelOffsetY, MapData(X, y).light_value, True
                     
3450                  Else
3460                  modGrh.Grh_RenderN MapData(X, y).Graphic(5), ((32 * ScreenX) - 32) + PixelOffsetX, ((32 * ScreenY) - 32) + PixelOffsetY, MapData(X, y).light_value, True
3470                  End If
3480              End If
3490  End If

              
3500          ScreenX = ScreenX + 1
3510      Next X
3520      ScreenY = ScreenY + 1
3530      If y > 100 Then Exit For
3540  Next y

3550  ScreenY = -8
3560  tiempo = 3


3570  For y = (MinY) To (MaxY)   '- 8+ 8
3580      ScreenX = -8
3590      For X = (MinX) To (MaxX)   '- 8 + 8
3600          If InMapBounds(X, y) Then
3610              If X > 100 Or X < -3 Then Exit For ' 30/05/2006

3620              iPPx = ((32 * ScreenX) - 32) + PixelOffsetX
3630              iPPy = ((32 * ScreenY) - 32) + PixelOffsetY
                   'Object Layer **********************************

3640               If MapData(X, y).OBJInfo.objindex <> 0 And VerObjetos Then
3650                  If MapData(X, y).Luz >= 202 And MapData(X, y).Luz <= 217 Then
3660                      modGrh.Grh_iRenderN MapData(X, y).ObjGrh, iPPx, iPPy, MapData(X, y).light_value, True
3670                  Else
3680                      modGrh.Grh_RenderN MapData(X, y).ObjGrh, iPPx, iPPy, MapData(X, y).light_value, True
3690                  End If
3700               End If
3710               If MapData(X, y).DecorI > 0 And MapData(X, y).DecorGrh.index > 0 And VerDecors Then
3720                  If TipoSeleccionado = 1 Then
3730                      If ObjetoSeleccionado.X = X And ObjetoSeleccionado.y = y Then
3740                          If MapData(X, y).Luz >= 202 And MapData(X, y).Luz <= 217 Then

3750                          modGrh.Grh_iRenderN SeleccionnGrh, iPPx, iPPy + (EstaticData(NewIndexData(SeleccionIndex).Estatic).H * 0.5), SeleccionadoArrayColor, True
3760                          Else
3770                          modGrh.Grh_RenderN SeleccionnGrh, iPPx, iPPy + (EstaticData(NewIndexData(SeleccionIndex).Estatic).H * 0.5), SeleccionadoArrayColor, True
                              
3780                          End If
3790                      End If
3800                  End If
3810                  If MapData(X, y).Luz >= 202 And MapData(X, y).Luz <= 217 Then
3820                      modGrh.Grh_iRenderN MapData(X, y).DecorGrh, iPPx, iPPy, MapData(X, y).light_value, True
                          
3830                  Else
3840                      modGrh.Grh_RenderN MapData(X, y).DecorGrh, iPPx, iPPy, MapData(X, y).light_value, True
3850                  End If

3860               End If
3870            tiempo = 4

                        'Char layer **********************************
3880                   If MapData(X, y).CHarIndex <> 0 And VerNpcs Then
                       
3890                       TempChar = CharList(MapData(X, y).CHarIndex)

3900                       PixelOffsetXTemp = PixelOffsetX
3910                       PixelOffsetYTemp = PixelOffsetY
                          
3920                      If TipoSeleccionado = 2 Then
3930                      If ObjetoSeleccionado.X = X And ObjetoSeleccionado.y = y Then
3940                          If MapData(X, y).Luz >= 202 And MapData(X, y).Luz <= 217 Then

3950                          modGrh.Grh_iRenderN SeleccionnGrh, iPPx, iPPy + (EstaticData(NewIndexData(SeleccionIndex).Estatic).H * 0.5), SeleccionadoArrayColor, True
3960                          Else
3970                          modGrh.Grh_RenderN SeleccionnGrh, iPPx, iPPy + (EstaticData(NewIndexData(SeleccionIndex).Estatic).H * 0.5), SeleccionadoArrayColor, True
                              
3980                          End If
3990                      End If
4000                  End If
                          
                          
                         'Dibuja solamente players
4010                     If TempChar.Head(TempChar.Heading).index <> 0 Then
4020                  If MapData(X, y).Luz >= 202 And MapData(X, y).Luz <= 217 Then
4030                       modGrh.Anim_iRender TempChar.Body(TempChar.Heading), iPPx, iPPy, MapData(X, y).light_value, True, False
                           'Draw Head
4040                       modGrh.Grh_iRenderN TempChar.Head(TempChar.Heading), iPPx, iPPy + BodyData(TempChar.iBody).OffsetY + HeadData(TempChar.iHead).OffsetDibujoY, MapData(X, y).light_value, True
                         
4050                  Else
4060                        modGrh.Anim_Render TempChar.Body(TempChar.Heading), iPPx, iPPy, MapData(X, y).light_value, True, False, BodyData(TempChar.iBody).OverWriteGrafico
                           'Draw Head
4070                       modGrh.Grh_RenderN TempChar.Head(TempChar.Heading), iPPx, iPPy + BodyData(TempChar.iBody).OffsetY + HeadData(TempChar.iHead).OffsetDibujoY, MapData(X, y).light_value, True
                                        
4080                  End If
4090                     Else
                         
4100                  If MapData(X, y).Luz >= 202 And MapData(X, y).Luz <= 217 Then
4110                       modGrh.Anim_iRender TempChar.Body(TempChar.Heading), iPPx, iPPy, MapData(X, y).light_value, True, False, BodyData(TempChar.iBody).OverWriteGrafico
4120                  Else
4130                        modGrh.Anim_Render TempChar.Body(TempChar.Heading), iPPx, iPPy, MapData(X, y).light_value, True, False, BodyData(TempChar.iBody).OverWriteGrafico
4140                  End If
4150              End If
                  
4160                   End If


4170                   tiempo = 5

                   'Layer 3 *****************************************
4180               If MapData(X, y).Graphic(3).index <> 0 And VerCapa3 Then
4190                  If MapData(X, y).Luz >= 202 And MapData(X, y).Luz <= 217 Then
4200                   modGrh.Grh_iRenderN MapData(X, y).Graphic(3), ((32 * ScreenX) - 32) + PixelOffsetX, ((32 * ScreenY) - 32) + PixelOffsetY, MapData(X, y).light_value, True
4210              Else

4220                  modGrh.Grh_RenderN MapData(X, y).Graphic(3), ((32 * ScreenX) - 32) + PixelOffsetX, ((32 * ScreenY) - 32) + PixelOffsetY, MapData(X, y).light_value, True
4230              End If
                   
4240               End If
4181                If MapData(X, y).Graphic(2).index < 0 And VerCapa9 Then
4191                 If MapData(X, y).Luz >= 202 And MapData(X, y).Luz <= 217 Then
4201                   modGrh.Grh_iRenderN MapData(X, y).Graphic(2), ((32 * ScreenX) - 32) + PixelOffsetX, ((32 * ScreenY) - 32) + PixelOffsetY, MapData(X, y).light_value, True
4211              Else

4221                  modGrh.Grh_RenderN MapData(X, y).Graphic(2), ((32 * ScreenX) - 32) + PixelOffsetX, ((32 * ScreenY) - 32) + PixelOffsetY, MapData(X, y).light_value, True
4231              End If
                   
4241              End If
                                     
4250               If MapData(X, y).SPOTLIGHT.index > 0 Then
4260                  SPOT_LIGHTS(MapData(X, y).SPOTLIGHT.index).X = ((32 * ScreenX) - 32) + PixelOffsetX
4270                  SPOT_LIGHTS(MapData(X, y).SPOTLIGHT.index).y = ((32 * ScreenY) - 32) + PixelOffsetY
4280                  SPOT_LIGHTS(MapData(X, y).SPOTLIGHT.index).Mustbe_Render = True
4290                  If frmMain.MarcarsPOT.value Then
4300                      nGrh.index = 247
4310                      modGrh.Grh_RenderN nGrh, ((32 * ScreenX) - 32) + PixelOffsetX, ((32 * ScreenY) - 32) + PixelOffsetY, MapData(X, y).light_value, True
4320                  End If
4330              End If
                   
4340               tiempo = 6

4350              tiempo = 7

4360          End If
              

4370          ScreenX = ScreenX + 1
4380      Next X
4390      ScreenY = ScreenY + 1
4400  Next y

      'Tiles blokeadas, techos, triggers , seleccion
4410  ScreenY = -8
4420  For y = (MinY) To (MaxY)
4430      ScreenX = -8
4440      For X = (MinX) To (MaxX)
4450          If X < 101 And X > 0 And y < 101 And y > 0 Then ' 30/05/2006
4460              iPPx = ((32 * ScreenX) - 32) + PixelOffsetX
4470              iPPy = ((32 * ScreenY) - 32) + PixelOffsetY

                  
4480                           If MapData(X, y).particle_group Then
4490                    modDXEngine.Particle_Group_Render MapData(X, y).particle_group, iPPx, iPPy

4500               End If
4510              If frmMain.cVerLuces.value And MapData(X, y).Luz > 0 Then
                      'modDXEngine.DXEngine_TextRender 1, MapData(x, Y).Luz, iPPx, iPPy, D3DColorXRGB(255, 0, 0), DT_CENTER, 32, 32
4520                  modDXEngine.DrawText iPPx, iPPy, MapData(X, y).Luz, D3DRED
4530              ElseIf frmMain.chkParticle.value And MapData(X, y).particle_group Then
4540                  DrawText iPPx, iPPy, CStr(MapData(X, y).parti_index), D3DWHITE
4550              ElseIf frmMain.ChkInterior.value And MapData(X, y).InteriorVal > 0 Then
4560                  DrawText iPPx, iPPy, CStr(MapData(X, y).InteriorVal), D3DWHITE
4570              End If
                  
                  
4580              If MapData(X, y).Graphic(4).index <> 0 _
                  And (frmMain.mnuVerCapa4.Checked = True) Then
4590                  If MapData(X, y).Luz >= 202 And MapData(X, y).Luz <= 217 Then
4600                      modGrh.Grh_RenderN MapData(X, y).Graphic(4), iPPx, iPPy, MapData(X, y).light_value, True
4610              Else
                  
4620                  modGrh.Grh_iRenderN MapData(X, y).Graphic(4), iPPx, iPPy, MapData(X, y).light_value, True
4630  End If
4640              End If
4650              If MapData(X, y).TileExit.Map <> 0 And VerTranslados Then
4660                  nGrh.index = 63
4670                  modGrh.Grh_RenderN nGrh, iPPx, iPPy, MapData(X, y).light_value, True
4680              End If
                  
4690              If MapData(X, y).light_index Then
4700                  nGrh.index = 247
4710                  modGrh.Grh_RenderN nGrh, iPPx, iPPy, colorlist, True
4720              End If
                  
                  'Show blocked tiles
4730              If VerBlockeados And MapData(X, y).Blocked = 1 Then
4740                  nGrh.index = 65
4750                  modGrh.Grh_RenderN nGrh, iPPx, iPPy, MapData(X, y).light_value, True
4760              End If
4770              If VerGrilla Then
                      'Grilla 24/11/2008 by GS
4780                  modDXEngine.DXEngine_DrawBox ((32 * ScreenX) - 32) + PixelOffsetX, ((32 * ScreenY) - 32) + PixelOffsetY, 1, 32, RGB(255, 255, 255)
4790                  modDXEngine.DXEngine_DrawBox ((32 * ScreenX) - 32) + PixelOffsetX, ((32 * ScreenY) - 32) + PixelOffsetY, 32, 1, RGB(255, 255, 255)
4800              End If
4810              If VerTriggers Then
                      'Call DrawText(PixelPos(ScreenX), PixelPos(ScreenY), Str(MapData(X, Y).Trigger), vbRed)
4820                  If frmMain.lListado(8).Visible Then
4830                      If MapData(X, y).TipoTerreno <> 0 Then
4840                     modDXEngine.DrawText ((32 * ScreenX) - 32) + PixelOffsetX, ((32 * ScreenY) - 32) + PixelOffsetY, "T:" & CStr(MapData(X, y).TipoTerreno), D3DWHITE
4850                     End If
4860                  Else
4870                      If MapData(X, y).Trigger <> 0 Then
4880                      modDXEngine.DrawText ((32 * ScreenX) - 32) + PixelOffsetX, ((32 * ScreenY) - 32) + PixelOffsetY, "G:" & CStr(MapData(X, y).Trigger), D3DWHITE
4890                      End If
4900                  End If
4910              End If
4920              If Seleccionando Then
                      'If ScreenX >= SeleccionIX And ScreenX <= SeleccionFX And ScreenY >= SeleccionIY And ScreenY <= SeleccionFY Then
4930                      If X >= SeleccionIX And y >= SeleccionIY Then
4940                          If X <= SeleccionFX And y <= SeleccionFY Then
4950                              modDXEngine.DXEngine_DrawBox ((32 * ScreenX) - 32) + PixelOffsetX, ((32 * ScreenY) - 32) + PixelOffsetY, 32, 32, RGB(100, 255, 255)
4960                          End If
4970                      End If
4980              End If

4990          End If
5000          ScreenX = ScreenX + 1
5010      Next X
5020      ScreenY = ScreenY + 1
5030  Next y
      Dim Xx(0 To 3) As Long
5040  tiempo = 100
5050  If (frmMain.cSeleccionarSuperficie.value Or frmMain.cQuitarEnEstaCapa.value Or frmMain.cQuitarEnTodasLasCapas.value Or ((frmMain.cInsertarTrigger.value Or frmMain.cQuitarTrigger.value Or frmMain.cVerTriggers.value) And (frmMain.lListado(8).Visible Or frmMain.lListado(4).Visible))) And SobreIndex > 0 Then

5060      Xx(0) = 0
5070      Xx(1) = 0
5080      Xx(2) = 0
5090      Xx(3) = -1
          Dim o As tnGrh
5100      o.index = SobreIndex
          
5110     modGrh.Grh_RenderN o, ((SobreX - (MinX + 9)) * 32), ((SobreY - (MinY + 9)) * 32), Xx, IIf(cCapaSel <> 2, True, False)

5120  End If




5130  Exit Sub

errs:
5140  Debug.Print Err.Description & "_" & X & "_" & y & "_" & tiempo & "_" & Erl

End Sub



Public Sub DrawText(lngXPos As Integer, lngYPos As Integer, strText As String, lngColor As Long)
'*************************************************
'Author: Unkwown
'Last modified: 26/05/06
'*************************************************
    If LenB(strText) <> 0 And lngXPos > 0 And lngYPos > 0 And lngColor <> 0 Then
        'Call modDXEngine.DXEngine_TextRender(1, strText, lngXPos, lngYPos, D3DColorXRGB(255, 255, 255))
        modDXEngine.DrawText lngXPos, lngYPos, strText, lngColor
    End If
End Sub

Function PixelPos(X As Integer) As Integer
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************

PixelPos = (TilePixelWidth * X) - TilePixelWidth

End Function

Function InitTileEngine(ByRef setDisplayFormhWnd As Long, setMainViewTop As Integer, setMainViewLeft As Integer, setTilePixelHeight As Integer, setTilePixelWidth As Integer, setWindowTileHeight As Integer, setWindowTileWidth As Integer, setTileBufferSize As Integer) As Boolean
'*************************************************
'Author: Unkwown
'Last modified: 15/10/06 by GS
'*************************************************
    'Fill startup variables
    DisplayFormhWnd = setDisplayFormhWnd
    MainViewTop = setMainViewTop
    MainViewLeft = setMainViewLeft
    TilePixelWidth = setTilePixelWidth
    TilePixelHeight = setTilePixelHeight
    WindowTileHeight = setWindowTileHeight
    WindowTileWidth = setWindowTileWidth
    TileBufferSize = setTileBufferSize

    
    '[GS] 02/10/2006
    MinXBorder = XMinMapSize + (ClienteWidth \ 2)
    MaxXBorder = XMaxMapSize - (ClienteWidth \ 2)
    MinYBorder = YMinMapSize + (ClienteHeight \ 2)
    MaxYBorder = YMaxMapSize - (ClienteHeight \ 2)
    
    MainViewWidth = (TilePixelWidth * WindowTileWidth)
    MainViewHeight = (TilePixelHeight * WindowTileHeight)
    'frmMain.MainViewPic.Width = MainViewWidth
    'frmMain.MainViewPic.Height = MainViewHeight
    Set Back_Sur = ddevice.GetRenderTarget
    
    Set Stencil = ddevice.CreateDepthStencilSurface(MainViewWidth, MainViewHeight, D3DFMT_D16, D3DMULTISAMPLE_NONE)

    
    Set Agua_Tex = d3dx.CreateTexture(ddevice, MainViewWidth, MainViewHeight, 0, D3DUSAGE_RENDERTARGET, D3DFMT_A8R8G8B8, D3DPOOL_DEFAULT)
    'Set Agua_Tex = d3dx.CreateTextureFromFileEx(ddevice, vbNullString, 800, 600, 0, D3DUSAGE_RENDERTARGET, D3DFMT_A8R8G8B8, D3DPOOL_DEFAULT, 0, 0, vbBlack, ByVal 0, ByVal 0)
    
    Set Agua_Sur = Agua_Tex.GetSurfaceLevel(0)
    
   ' Set Agua_St = ddevice.CreateDepthStencilSurface(800, 600, D3DFMT_D16, D3DMULTISAMPLE_NONE)
    Set Agua_St = ddevice.GetDepthStencilSurface
    
    'Resize mapdata array
    ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    
    InitTileEngine = True
    EngineRun = True
    DoEvents
End Function

Public Sub LightSet(ByVal X As Byte, ByVal y As Byte, ByVal Rounded As Boolean, ByVal Range As Integer, ByVal R As Byte, ByVal G As Byte, ByVal B As Byte)
    Dim min_x As Integer
    Dim min_y As Integer
    Dim max_x As Integer
    Dim max_y As Integer
    Dim iX As Integer
    Dim iY As Integer
    Dim i As Integer
    
    If Rounded Then
        For i = 1 To Light_Count
            If Light_Count = 0 Then Exit For
            If Lights(i).Active = 0 Then
                Exit For
            End If
        Next i
        If i > Light_Count Then
            Light_Count = Light_Count + 1
            i = Light_Count
        End If
        MapData(X, y).light_index = i
        ReDim Preserve Lights(1 To Light_Count) As Light
        Lights(i).Active = True
        Lights(i).map_x = X
        Lights(i).map_y = y
        Lights(i).X = X * 32
        Lights(i).y = y * 32
        Lights(i).Range = Range
        Lights(i).RGBCOLOR.A = 255
        Lights(i).RGBCOLOR.R = R
        Lights(i).RGBCOLOR.G = G
        Lights(i).RGBCOLOR.B = B
    Else
        'Set up light borders
        min_x = X - Range
        min_y = y - Range
        max_x = X + Range
        max_y = y + Range
    
        If InMapBounds(min_x, min_y) Then
            MapData(min_x, min_y).base_light(2) = True
            MapData(min_x, min_y).light_base_value(2) = D3DColorXRGB(R, G, B)
        End If
        If InMapBounds(min_x, max_y) Then
            MapData(min_x, max_y).base_light(3) = True
            MapData(min_x, max_y).light_base_value(3) = D3DColorXRGB(R, G, B)
        End If
        If InMapBounds(max_x, min_y) Then
            MapData(max_x, min_y).base_light(0) = True
            MapData(max_x, min_y).light_base_value(0) = D3DColorXRGB(R, G, B)
        End If
        If InMapBounds(max_x, max_y) Then
            MapData(max_x, max_y).base_light(1) = True
            MapData(max_x, max_y).light_base_value(1) = D3DColorXRGB(R, G, B)
        End If
        
        'Upper Border
        For iX = min_x + 1 To max_x - 1
            If InMapBounds(iX, min_y) Then
                MapData(iX, min_y).base_light(0) = True
                MapData(iX, min_y).light_base_value(0) = D3DColorXRGB(R, G, B)
                MapData(iX, min_y).base_light(2) = True
                MapData(iX, min_y).light_base_value(2) = D3DColorXRGB(R, G, B)
            End If
        Next iX
        
        'Lower Border
        For iX = min_x + 1 To max_x - 1
            If InMapBounds(iX, max_y) Then
                MapData(iX, max_y).base_light(3) = True
                MapData(iX, max_y).light_base_value(3) = D3DColorXRGB(R, G, B)
                MapData(iX, max_y).base_light(1) = True
                MapData(iX, max_y).light_base_value(1) = D3DColorXRGB(R, G, B)
            End If
        Next iX
        
        'Right Border
        For iY = min_y + 1 To max_y - 1
            If InMapBounds(max_x, iY) Then
                MapData(max_x, iY).base_light(1) = True
                MapData(max_x, iY).light_base_value(1) = D3DColorXRGB(R, G, B)
                MapData(max_x, iY).base_light(0) = True
                MapData(max_x, iY).light_base_value(0) = D3DColorXRGB(R, G, B)
            End If
        Next iY
        
        'Left Border
        For iY = min_y + 1 To max_y - 1
            If InMapBounds(min_x, iY) Then
                MapData(min_x, iY).base_light(3) = True
                MapData(min_x, iY).light_base_value(3) = D3DColorXRGB(R, G, B)
                MapData(min_x, iY).base_light(2) = True
                MapData(min_x, iY).light_base_value(2) = D3DColorXRGB(R, G, B)
            End If
        Next iY
        
        'Left Border
        For iY = min_y + 1 To max_y - 1
            For iX = min_x + 1 To max_x - 1
                If InMapBounds(iX, iY) Then
                    MapData(iX, iY).base_light(3) = True
                    MapData(iX, iY).light_base_value(3) = D3DColorXRGB(R, G, B)
                    MapData(iX, iY).base_light(2) = True
                    MapData(iX, iY).light_base_value(2) = D3DColorXRGB(R, G, B)
                    MapData(iX, iY).base_light(1) = True
                    MapData(iX, iY).light_base_value(1) = D3DColorXRGB(R, G, B)
                    MapData(iX, iY).base_light(0) = True
                    MapData(iX, iY).light_base_value(0) = D3DColorXRGB(R, G, B)
                End If
            Next iX
        Next iY
    End If
End Sub


Public Sub Map_LightsRender()
    Dim i As Integer
    
    Call Map_LightsClear
    
    For i = 1 To Light_Count
        Map_LightRender (i)
    Next i
End Sub

Public Function Map_LightsClear()
On Error GoTo errx
    Dim X As Integer
    Dim y As Integer
    Dim Luz As Byte
    Dim AmbientColor As D3DCOLORVALUE
    Dim Color As Long
    
    
    
    Meteo.Get_AmbientLight AmbientColor
    Color = D3DColorXRGB(AmbientColor.R, AmbientColor.G, AmbientColor.B)
    
    
    Luz = HoraLuz
    For X = 1 To 100
        For y = 1 To 100
  '          If X = 90 And Y = 55 Then Stop
            If InMapBounds(X, y) Then
                With MapData(X, y)
                    'If X = 13 And y = 74 Then Stop
                    'If MapData(X, Y).Luz = 8 Then Stop
                    If MapData(X, y).Luz > 0 And MapData(X, y).Luz < 200 Then
                    If MapData(X, y).Luz <= 8 Then
                    .light_value(0) = ambient_light((HoraLuz * 9) + MapData(X, y).Luz + 1)
                    .light_value(1) = ambient_light((HoraLuz * 9) + MapData(X, y).Luz + 1)
                    .light_value(2) = ambient_light((HoraLuz * 9) + MapData(X, y).Luz + 1)
                    .light_value(3) = ambient_light((HoraLuz * 9) + MapData(X, y).Luz + 1)
                    
                    .LV(0) = MapData(X, y).Luz
                    .LV(1) = MapData(X, y).Luz
                    .LV(2) = MapData(X, y).Luz
                    .LV(3) = MapData(X, y).Luz
                    ElseIf MapData(X, y).Luz > 0 Then 'Bordes, cargamos los cosos.
                    
                        If MapData(X, y).LV(0) > 0 And MapData(X, y).LV(0) < 9 Then 'Luces normales
                            MapData(X, y).light_value(0) = ambient_light(((Luz * 9) + MapData(X, y).LV(0) + 1))

                        ElseIf MapData(X, y).LV(0) = 9 Then
                            MapData(X, y).light_value(0) = extra_light(eE_Light.Oscuridad)
                        ElseIf MapData(X, y).LV(0) = 11 Then
                            MapData(X, y).light_value(0) = extra_light(eE_Light.Azul1)
                        ElseIf MapData(X, y).LV(0) = 12 Then
                            MapData(X, y).light_value(0) = extra_light(eE_Light.Azul2)
                        ElseIf MapData(X, y).LV(0) = 13 Then
                            MapData(X, y).light_value(0) = extra_light(eE_Light.Azul3)
                                                Else
                            MapData(X, y).light_value(0) = 0
                        End If
                        If MapData(X, y).LV(1) > 0 And MapData(X, y).LV(1) < 9 Then 'Luces normales
                            MapData(X, y).light_value(1) = ambient_light((Luz * 9) + MapData(X, y).LV(1) + 1)
                        ElseIf MapData(X, y).LV(1) = 9 Then
                            MapData(X, y).light_value(1) = extra_light(eE_Light.Oscuridad)
                        ElseIf MapData(X, y).LV(1) = 11 Then
                            MapData(X, y).light_value(1) = extra_light(eE_Light.Azul1)
                        ElseIf MapData(X, y).LV(1) = 12 Then
                            MapData(X, y).light_value(1) = extra_light(eE_Light.Azul2)
                        ElseIf MapData(X, y).LV(1) = 13 Then
                            MapData(X, y).light_value(1) = extra_light(eE_Light.Azul3)
                                                Else
                            MapData(X, y).light_value(1) = 0
                        End If
                        If MapData(X, y).LV(2) > 0 And MapData(X, y).LV(2) < 9 Then 'Luces normales
                            MapData(X, y).light_value(2) = ambient_light((Luz * 9) + MapData(X, y).LV(2) + 1)
                        ElseIf MapData(X, y).LV(2) = 9 Then
                            MapData(X, y).light_value(2) = extra_light(eE_Light.Oscuridad)
                        ElseIf MapData(X, y).LV(2) = 11 Then
                            MapData(X, y).light_value(2) = extra_light(eE_Light.Azul1)
                        ElseIf MapData(X, y).LV(2) = 12 Then
                            MapData(X, y).light_value(2) = extra_light(eE_Light.Azul2)
                        ElseIf MapData(X, y).LV(2) = 13 Then
                            MapData(X, y).light_value(2) = extra_light(eE_Light.Azul3)
                        Else
                            MapData(X, y).light_value(2) = 0
                        End If
                        If MapData(X, y).LV(3) > 0 And MapData(X, y).LV(3) < 9 Then 'Luces normales
                            MapData(X, y).light_value(3) = ambient_light((Luz * 9) + MapData(X, y).LV(3) + 1)
                        ElseIf MapData(X, y).LV(3) = 9 Then
                            MapData(X, y).light_value(3) = extra_light(eE_Light.Oscuridad)
                        ElseIf MapData(X, y).LV(3) = 11 Then
                            MapData(X, y).light_value(3) = extra_light(eE_Light.Azul1)
                        ElseIf MapData(X, y).LV(3) = 12 Then
                            MapData(X, y).LV(3) = extra_light(eE_Light.Azul2)
                        ElseIf MapData(X, y).LV(3) = 13 Then
                            MapData(X, y).light_value(3) = extra_light(eE_Light.Azul3)
                                                Else
                            MapData(X, y).light_value(3) = 0
                        End If

                    
                    End If
                    End If
                    
                End With
            End If
        Next y
    Next X
    Exit Function
errx:
    '"ERRMAPALIGHTS:" & Err.Description & "_" & X & "_" & Y
    
End Function

Private Sub Map_LightRender(ByVal light_index As Integer)
    Dim min_x As Integer
    Dim min_y As Integer
    Dim max_x As Integer
    Dim max_y As Integer
    Dim Ya As Integer
    Dim Xa As Integer
    
    Dim AmbientColor As D3DCOLORVALUE
    Dim LightColor As D3DCOLORVALUE
    
    Dim XCoord As Integer
    Dim YCoord As Integer
        
        LightColor = Lights(light_index).RGBCOLOR
        Meteo.Get_AmbientLight AmbientColor
        
        If Not Lights(light_index).Active = True Then Exit Sub
        
        min_x = Lights(light_index).map_x - Lights(light_index).Range
        max_x = Lights(light_index).map_x + Lights(light_index).Range
        min_y = Lights(light_index).map_y - Lights(light_index).Range
        max_y = Lights(light_index).map_y + Lights(light_index).Range
        
        For Ya = min_y To max_y
            For Xa = min_x To max_x
                If InMapBounds(Xa, Ya) Then
                    XCoord = Xa * 32
                    YCoord = Ya * 32
                    'Color = LightCalculate(lights(light_index).range, lights(light_index).x, lights(light_index).y, XCoord, YCoord, mapdata(Xa, Ya).light_value(1), LightColor, AmbientColor)
                    MapData(Xa, Ya).light_value(1) = LightCalculate(Lights(light_index).Range, Lights(light_index).X, Lights(light_index).y, XCoord, YCoord, MapData(Xa, Ya).light_value(1), LightColor, AmbientColor)

                    XCoord = Xa * 32 + 32
                    YCoord = Ya * 32
                    MapData(Xa, Ya).light_value(3) = LightCalculate(Lights(light_index).Range, Lights(light_index).X, Lights(light_index).y, XCoord, YCoord, MapData(Xa, Ya).light_value(3), LightColor, AmbientColor)
                       
                    XCoord = Xa * 32
                    YCoord = Ya * 32 + 32
                    MapData(Xa, Ya).light_value(0) = LightCalculate(Lights(light_index).Range, Lights(light_index).X, Lights(light_index).y, XCoord, YCoord, MapData(Xa, Ya).light_value(0), LightColor, AmbientColor)
    
                    XCoord = Xa * 32 + 32
                    YCoord = Ya * 32 + 32
                    MapData(Xa, Ya).light_value(2) = LightCalculate(Lights(light_index).Range, Lights(light_index).X, Lights(light_index).y, XCoord, YCoord, MapData(Xa, Ya).light_value(2), LightColor, AmbientColor)
                End If
            Next Xa
        Next Ya
End Sub

Private Function LightCalculate(ByVal cRadio As Integer, ByVal LightX As Integer, ByVal LightY As Integer, ByVal XCoord As Integer, ByVal YCoord As Integer, TileLight As Long, LightColor As D3DCOLORVALUE, AmbientColor As D3DCOLORVALUE) As Long
    Dim XDist As Single
    Dim YDist As Single
    Dim VertexDist As Single
    Dim pRadio As Integer
    
    Dim CurrentColor As D3DCOLORVALUE
    
    pRadio = cRadio * 32
    
    XDist = LightX + 16 - XCoord
    YDist = LightY + 16 - YCoord
    
    VertexDist = Sqr(XDist * XDist + YDist * YDist)
    
    If VertexDist <= pRadio Then
        Call D3DXColorLerp(CurrentColor, LightColor, AmbientColor, VertexDist / pRadio)
        LightCalculate = D3DColorXRGB(CurrentColor.R, CurrentColor.G, CurrentColor.B)
        If TileLight > LightCalculate Then LightCalculate = TileLight
    Else
        LightCalculate = TileLight
    End If
End Function

Public Sub LightDestroy(ByVal X As Byte, ByVal y As Byte)
    If MapData(X, y).light_index Then
        Lights(MapData(X, y).light_index).Active = False
        MapData(X, y).light_index = 0
    Else
        MapData(X, y).base_light(0) = False
        MapData(X, y).base_light(1) = False
        MapData(X, y).base_light(2) = False
        MapData(X, y).base_light(3) = False
    End If
End Sub

Public Sub LightDestroyAll()
    Dim X As Integer
    Dim y As Integer
    For X = 1 To 100
        For y = 1 To 100
        Call LightDestroy(X, y)
        Next y
    Next X
End Sub

Sub Map_ResetMontañita()
Dim xb As Integer, yb As Integer, i As Byte

For xb = MinXBorder To MaxXBorder
    For yb = MinYBorder To MaxYBorder
        For i = 0 To 3
            MapData(xb, yb).AlturaPoligonos(i) = 0
        Next i
    Next yb
Next xb
End Sub
Sub Map_CreateMontañita(X As Integer, y As Integer, Radio As Byte, alturamaxima As Integer)
 
Dim xb As Integer, yb As Integer

For xb = X - Radio To X + Radio
    For yb = y - Radio To y + Radio
        'For i = 0 To 3

            MapData(xb, yb).AlturaPoligonos(0) = CalcularAlturaPoligono(xb * 32, yb * 32, X * 32, y * 32, Radio, alturamaxima)
            MapData(xb, yb).AlturaPoligonos(1) = CalcularAlturaPoligono(xb * 32 + 32, yb * 32, X * 32, y * 32, Radio, alturamaxima)
            MapData(xb, yb).AlturaPoligonos(2) = CalcularAlturaPoligono(xb * 32, yb * 32 + 32, X * 32, y * 32, Radio, alturamaxima)
            MapData(xb, yb).AlturaPoligonos(3) = CalcularAlturaPoligono(xb * 32 + 32, yb * 32 + 32, X * 32, y * 32, Radio, alturamaxima)
        
        'Next i
    Next yb
Next xb
 
    'Orden del poligono:
    '0---1
    '|  /|
    '| / |
    '|/  |
    '2---3
       
End Sub
 
Function CalcularAlturaPoligono(Mx As Integer, My As Integer, Xn As Integer, Yn As Integer, Radio As Byte, am As Integer) As Integer
       
       
    Dim Dp As Integer, Dm As Integer
    Dp = Abs(Mx - Xn) + Abs(My - Yn)
    Dm = Radio * 32
   
    CalcularAlturaPoligono = Val(am * (1 - (Dp / Dm)))
    If CalcularAlturaPoligono < 0 Then CalcularAlturaPoligono = 0

End Function

Function HayAgua(ByVal X As Integer, ByVal y As Integer) As Boolean

If X < 1 Or X > 100 Or y < 1 Or y > 100 Then Exit Function

    If MapData(X, y).Graphic(2).index > 0 Then
        If NewIndexData(MapData(X, y).Graphic(2).index).OverWriteGrafico = 20 Then
            HayAgua = True
        End If
    End If
           
End Function
Public Sub Resetalllights()
Dim X As Integer
Dim y As Integer
For X = 1 To 100
For y = 1 To 100
MapData(X, y).Luz = 0
MapData(X, y).LV(0) = 0
MapData(X, y).LV(1) = 0
MapData(X, y).LV(2) = 0
MapData(X, y).LV(3) = 0
Next y
Next X

End Sub
Public Sub AplicarBordeManual(ByVal X As Byte, ByVal y As Byte, ByVal Tipo As Byte)
'Aplicamos el Borde seleccionado.
Dim QueBorde As Byte
Dim Puede As Byte

'0          1
'
'
'2          3


If InMapBounds(X, y) Then
    'Cliqueo bien el putito.
    QueBorde = 255 - Tipo
    
        Select Case QueBorde
            
            Case eB_Light.Horizontal
                If InMapBounds(X, y - 1) Then Puede = 1
                If InMapBounds(X, y + 1) Then Puede = Puede + 2
                If Puede = 3 Then 'Ambos TILES existen.
                    MapData(X, y).Luz = eB_Light.Horizontal
                    'Esto podria ir fuera de este ifClause, pero en realidad seria mas optimo
                    'Para el cliente, que en los casos siguientes fueran TIPOS especiales
                    'Para evitar estos chequeos en el cliente donde el procesador nos importa.
                    'Lean1! Proxima revision.
                                 
                        
                    'Vertices Superiores
                    MapData(X, y).light_value(0) = MapData(X, y - 1).light_value(2)
                    MapData(X, y).light_value(1) = MapData(X, y - 1).light_value(3)
                     MapData(X, y).light_value(2) = MapData(X, y + 1).light_value(0)
                    MapData(X, y).light_value(3) = MapData(X, y + 1).light_value(1)
                    
                    MapData(X, y).LV(1) = MapData(X, y - 1).LV(3)
                    MapData(X, y).LV(0) = MapData(X, y - 1).LV(2)
                     MapData(X, y).LV(2) = MapData(X, y + 1).LV(0)
                    MapData(X, y).LV(3) = MapData(X, y + 1).LV(1)
                    
                ElseIf Puede = 2 Then 'No existe el Tile Inferior
                    MapData(X, y).Luz = eB_Light.HSoloUpper
                    'Lean! deberian ser distintos... EB_LIGHT.HSoloUPPER
                     'Vertices Superiores
                    MapData(X, y).light_value(0) = MapData(X, y - 1).light_value(2)
                    MapData(X, y).light_value(1) = MapData(X, y - 1).light_value(3)
                    
                    'Vertices Inferiores
                    MapData(X, y).light_value(2) = 0
                    MapData(X, y).light_value(3) = 0
                    
                    MapData(X, y).LV(1) = MapData(X, y - 1).LV(3)
                    MapData(X, y).LV(0) = MapData(X, y - 1).LV(2)
                    
                    'Vertices Inferiores
                    MapData(X, y).LV(3) = 0
                    MapData(X, y).LV(2) = 0
                ElseIf Puede = 1 Then 'No existe el Tile Superior
                    MapData(X, y).Luz = eB_Light.HSoloBottom 'Lean! ...EB_LIGHT.HSoloBOTTOM
                    'Vertices Superiores
                    MapData(X, y).light_value(0) = 0
                    MapData(X, y).light_value(1) = 0
                    
                    'Vertices Inferiores
                    MapData(X, y).light_value(2) = MapData(X, y + 1).light_value(0)
                    MapData(X, y).light_value(3) = MapData(X, y + 1).light_value(1)
                    
                    
                                    'Vertices Superiores
                    MapData(X, y).LV(1) = 0
                    MapData(X, y).LV(0) = 0
                    
                    'Vertices Inferiores
                    MapData(X, y).LV(3) = MapData(X, y + 1).LV(1)
                    MapData(X, y).LV(2) = MapData(X, y + 1).LV(0)
                End If
            Case eB_Light.Vertical
                If InMapBounds(X - 1, y) Then Puede = 1
                If InMapBounds(X + 1, y) Then Puede = Puede + 2
                
                If Puede = 3 Then ' Derecha e Izquierda
                    MapData(X, y).Luz = eB_Light.Vertical 'Lean! mismo que para arriba deberia ser distintos
                    
                    'Vertices Izquierda
                    MapData(X, y).light_value(0) = MapData(X - 1, y).light_value(1)
                    MapData(X, y).light_value(2) = MapData(X - 1, y).light_value(3)
    
                    
                    'Vertices Derecha
                    MapData(X, y).light_value(1) = MapData(X + 1, y).light_value(0)
                    MapData(X, y).light_value(3) = MapData(X + 1, y).light_value(2)
                    
                    
                    MapData(X, y).LV(2) = MapData(X - 1, y).LV(3)
                    MapData(X, y).LV(0) = MapData(X - 1, y).LV(1)
    
                    
                    'Vertices Derecha
                    MapData(X, y).LV(3) = MapData(X + 1, y).LV(2)
                    MapData(X, y).LV(1) = MapData(X + 1, y).LV(0)
                
                ElseIf Puede = 1 Then 'Solo Izquierda
                    MapData(X, y).Luz = eB_Light.VSoloLeft 'Lean! eb_light.VSoloLEFT
                    
                    'Vertices Izquierda
                    MapData(X, y).light_value(0) = MapData(X - 1, y).light_value(1)
                    MapData(X, y).light_value(2) = MapData(X - 1, y).light_value(3)
    
                    
                    'Vertices Derecha
                    MapData(X, y).light_value(1) = 0
                    MapData(X, y).light_value(3) = 0
                
                
                
                                    'Vertices Izquierda
                    MapData(X, y).LV(2) = MapData(X - 1, y).LV(3)
                    MapData(X, y).LV(0) = MapData(X - 1, y).LV(1)
    
                    
                    'Vertices Derecha
                    MapData(X, y).LV(3) = 0
                    MapData(X, y).LV(1) = 0
                
                ElseIf Puede = 2 Then 'Solo Derecha
                    MapData(X, y).Luz = eB_Light.VSoloRight 'Lean! eb_Light.VSoloRIGHT
                    
                    'Vertices Izquierda
                    MapData(X, y).light_value(0) = 0
                    MapData(X, y).light_value(2) = 0
    
                    
                    'Vertices Derecha
                    MapData(X, y).light_value(1) = MapData(X + 1, y).light_value(0)
                    MapData(X, y).light_value(3) = MapData(X + 1, y).light_value(2)
                    
                                        'Vertices Izquierda
                    MapData(X, y).LV(0) = 0
                    MapData(X, y).LV(2) = 0
    
                    
                    'Vertices Derecha
                    MapData(X, y).LV(3) = MapData(X + 1, y).LV(2)
                    MapData(X, y).LV(1) = MapData(X + 1, y).LV(0)
                    
                
                End If
            Case eB_Light.UpperLeft
                If InMapBounds(X - 1, y - 1) Then
                    'Si hay un tile con luz arriba y la izquierda.
                    If MapData(X - 1, y - 1).Luz <> 0 And MapData(X - 1, y - 1).Luz < EB_LIMITE_INFERIOR Then
                        Puede = 1
                    End If
                End If
                If Puede = 0 And InMapBounds(X - 1, y) Then
                    If MapData(X - 1, y).Luz <> 0 And MapData(X - 1, y).Luz < EB_LIMITE_INFERIOR Then
                        Puede = 2
                    End If
                End If
                If Puede = 0 And InMapBounds(X, y - 1) Then
                    If MapData(X, y - 1).Luz <> 0 And MapData(X, y - 1).Luz < EB_LIMITE_INFERIOR Then
                        Puede = 3
                    End If
                End If
                
                If Puede = 1 Then 'Buscamos la luz del tile izquierdo superior
                    MapData(X, y).Luz = eB_Light.UpperLeft
                    MapData(X, y).light_value(0) = MapData(X - 1, y - 1).light_value(3)
                    MapData(X, y).light_value(1) = 0
                    MapData(X, y).light_value(2) = 0
                    MapData(X, y).light_value(3) = 0
                    
                    MapData(X, y).LV(0) = MapData(X - 1, y - 1).LV(3)
                    MapData(X, y).LV(3) = 0
                    MapData(X, y).LV(1) = 0
                    MapData(X, y).LV(2) = 0
                
                ElseIf Puede = 2 Then 'Buscamos la luz del tile de la izquierda
                    MapData(X, y).Luz = eB_Light.HUpperLeft 'Lean! Deberia ser distinto
                    'eb_light.HUpperLeft
                    
                    MapData(X, y).light_value(0) = MapData(X - 1, y).light_value(1)
                    MapData(X, y).light_value(1) = 0
                    MapData(X, y).light_value(2) = 0
                    MapData(X, y).light_value(3) = 0
                    
                    
                    MapData(X, y).LV(0) = MapData(X - 1, y).LV(1)
                    MapData(X, y).LV(3) = 0
                    MapData(X, y).LV(1) = 0
                    MapData(X, y).LV(2) = 0
                    
                ElseIf Puede = 3 Then
                    MapData(X, y).Luz = eB_Light.VUpperLeft 'Lean! Deberia ser distinto
                    'eb_light.VUpperLeft
                    
                    MapData(X, y).light_value(0) = MapData(X, y - 1).light_value(2)
                    MapData(X, y).light_value(1) = 0
                    MapData(X, y).light_value(2) = 0
                    MapData(X, y).light_value(3) = 0
                    
                    MapData(X, y).LV(0) = MapData(X, y - 1).LV(2)
                    MapData(X, y).LV(3) = 0
                    MapData(X, y).LV(1) = 0
                    MapData(X, y).LV(2) = 0
                    
                End If
             Case eB_Light.UpperRight
                If InMapBounds(X + 1, y - 1) Then
                    'Si hay un tile con luz arriba y la derecha
                    If MapData(X + 1, y - 1).Luz <> 0 And MapData(X + 1, y - 1).Luz < EB_LIMITE_INFERIOR Then
                        Puede = 1
                    End If
                End If
                If Puede = 0 And InMapBounds(X + 1, y) Then
                    If MapData(X + 1, y).Luz <> 0 And MapData(X + 1, y).Luz < EB_LIMITE_INFERIOR Then
                        Puede = 2
                    End If
                End If
                If Puede = 0 And InMapBounds(X, y - 1) Then
                    If MapData(X, y - 1).Luz <> 0 And MapData(X, y - 1).Luz < EB_LIMITE_INFERIOR Then
                        Puede = 3
                    End If
                End If
                
                If Puede = 1 Then 'Buscamos la luz del tile derecho superior
                    MapData(X, y).Luz = eB_Light.UpperRight
                    MapData(X, y).light_value(1) = MapData(X + 1, y - 1).light_value(2)
                    MapData(X, y).light_value(0) = 0
                    MapData(X, y).light_value(2) = 0
                    MapData(X, y).light_value(3) = 0
                    
                    
                    MapData(X, y).LV(1) = MapData(X + 1, y - 1).LV(2)
                    MapData(X, y).LV(3) = 0
                    MapData(X, y).LV(0) = 0
                    MapData(X, y).LV(2) = 0
                
                ElseIf Puede = 2 Then 'Buscamos la luz del tile de la derecha
                    MapData(X, y).Luz = eB_Light.HUpperRight 'Lean! Deberia ser distinto
                    'eb_light.HUpperRight
                    
                    MapData(X, y).light_value(1) = MapData(X + 1, y).light_value(0)
                    MapData(X, y).light_value(0) = 0
                    MapData(X, y).light_value(2) = 0
                    MapData(X, y).light_value(3) = 0
                    
                    
                    MapData(X, y).LV(1) = MapData(X + 1, y).LV(0)
                    MapData(X, y).LV(3) = 0
                    MapData(X, y).LV(0) = 0
                    MapData(X, y).LV(2) = 0
                ElseIf Puede = 3 Then 'Buscamos la luz del tile de la derecha
                    MapData(X, y).Luz = eB_Light.VUpperRight  'Lean! Deberia ser distinto
                    'eb_light.VUpperRight
                    
                    MapData(X, y).light_value(1) = MapData(X, y - 1).light_value(3)
                    MapData(X, y).light_value(0) = 0
                    MapData(X, y).light_value(2) = 0
                    MapData(X, y).light_value(3) = 0
                    
                    
                    MapData(X, y).LV(1) = MapData(X, y - 1).LV(3)
                    MapData(X, y).LV(3) = 0
                    MapData(X, y).LV(0) = 0
                    MapData(X, y).LV(2) = 0
                End If
                
             Case eB_Light.BottomRight
                If InMapBounds(X + 1, y + 1) Then
                    'Si hay un tile con luz arriba y la derecha
                    If MapData(X + 1, y + 1).Luz <> 0 And MapData(X + 1, y + 1).Luz < EB_LIMITE_INFERIOR Then
                        Puede = 1
                    End If
                End If
                If Puede = 0 And InMapBounds(X + 1, y) Then
                    If MapData(X + 1, y).Luz <> 0 And MapData(X + 1, y).Luz < EB_LIMITE_INFERIOR Then
                        Puede = 2
                    End If
                End If
                If Puede = 0 And InMapBounds(X, y + 1) Then
                    If MapData(X, y + 1).Luz <> 0 And MapData(X, y + 1).Luz < EB_LIMITE_INFERIOR Then
                        Puede = 3
                    End If
                End If
                
                If Puede = 1 Then 'Buscamos la luz del tile derecho superior
                    MapData(X, y).Luz = eB_Light.BottomRight
                    MapData(X, y).light_value(3) = MapData(X + 1, y + 1).light_value(0)
                    MapData(X, y).light_value(0) = 0
                    MapData(X, y).light_value(2) = 0
                    MapData(X, y).light_value(1) = 0
                
                    MapData(X, y).LV(3) = MapData(X + 1, y + 1).LV(0)
                    MapData(X, y).LV(0) = 0
                    MapData(X, y).LV(2) = 0
                    MapData(X, y).LV(1) = 0
                ElseIf Puede = 2 Then 'Buscamos la luz del tile de la derecha
                    MapData(X, y).Luz = eB_Light.HBottomRight 'Lean! Deberia ser distinto
                    'eb_light.HBottomRight
                    
                    MapData(X, y).light_value(3) = MapData(X + 1, y).light_value(2)
                    MapData(X, y).light_value(0) = 0
                    MapData(X, y).light_value(2) = 0
                    MapData(X, y).light_value(1) = 0
                    
                    
                    MapData(X, y).LV(3) = MapData(X + 1, y).LV(2)
                    MapData(X, y).LV(0) = 0
                    MapData(X, y).LV(2) = 0
                    MapData(X, y).LV(1) = 0
                ElseIf Puede = 3 Then 'Buscamos la luz del tile de la derecha
                    MapData(X, y).Luz = eB_Light.VBottomRight 'Lean! Deberia ser distinto
                    'eb_light.VBottomRight
                    
                    MapData(X, y).light_value(3) = MapData(X, y + 1).light_value(1)
                    MapData(X, y).light_value(0) = 0
                    MapData(X, y).light_value(2) = 0
                    MapData(X, y).light_value(1) = 0
                    
                    
                    
                    MapData(X, y).LV(3) = MapData(X, y + 1).LV(1)
                    MapData(X, y).LV(1) = 0
                    MapData(X, y).LV(0) = 0
                    MapData(X, y).LV(2) = 0
                End If
             Case eB_Light.BottomLeft
                If InMapBounds(X - 1, y + 1) Then
                    'Si hay un tile con luz arriba y la derecha
                    If MapData(X - 1, y + 1).Luz <> 0 And MapData(X - 1, y + 1).Luz < EB_LIMITE_INFERIOR Then
                        Puede = 1
                    End If
                End If
                If Puede = 0 And InMapBounds(X + 1, y) Then
                    If MapData(X - 1, y).Luz <> 0 And MapData(X - 1, y).Luz < EB_LIMITE_INFERIOR Then
                        Puede = 2
                    End If
                End If
                If Puede = 0 And InMapBounds(X, y + 1) Then
                    If MapData(X, y + 1).Luz <> 0 And MapData(X, y + 1).Luz < EB_LIMITE_INFERIOR Then
                        Puede = 3
                    End If
                End If
                
                If Puede = 1 Then 'Buscamos la luz del tile derecho superior
                    MapData(X, y).Luz = eB_Light.BottomLeft
                    MapData(X, y).light_value(2) = MapData(X - 1, y + 1).light_value(1)
                    MapData(X, y).light_value(0) = 0
                    MapData(X, y).light_value(3) = 0
                    MapData(X, y).light_value(1) = 0
                    
                    
                    MapData(X, y).LV(2) = MapData(X - 1, y + 1).LV(1)
                    MapData(X, y).LV(1) = 0
                    MapData(X, y).LV(0) = 0
                    MapData(X, y).LV(3) = 0
                
                ElseIf Puede = 2 Then 'Buscamos la luz del tile de la derecha
                    MapData(X, y).Luz = eB_Light.HBottomLeft 'Lean! Deberia ser distinto
                    'eb_light.HBottomLeft
                    
                    MapData(X, y).light_value(2) = MapData(X - 1, y).light_value(3)
                    MapData(X, y).light_value(0) = 0
                    MapData(X, y).light_value(3) = 0
                    MapData(X, y).light_value(1) = 0
                    
                    
                    
                    MapData(X, y).LV(2) = MapData(X - 1, y).LV(3)
                    MapData(X, y).LV(1) = 0
                    MapData(X, y).LV(0) = 0
                    MapData(X, y).LV(3) = 0
                ElseIf Puede = 3 Then 'Buscamos la luz del tile de la derecha
                    MapData(X, y).Luz = eB_Light.VBottomLeft 'Lean! Deberia ser distinto
                    'eb_light.VBottomLeft
                    
                    MapData(X, y).light_value(2) = MapData(X, y + 1).light_value(1)
                    MapData(X, y).light_value(0) = 0
                    MapData(X, y).light_value(3) = 0
                    MapData(X, y).light_value(1) = 0
                    
                    MapData(X, y).LV(2) = MapData(X, y + 1).LV(1)
                    MapData(X, y).LV(1) = 0
                    MapData(X, y).LV(0) = 0
                    MapData(X, y).LV(3) = 0
                    
                End If
        Case eB_Light.CrossLeftUp
            'Izquierda Arriba , Derecha Abajo.
            If InMapBounds(X - 1, y - 1) Then
                If MapData(X - 1, y - 1).Luz <> 0 And MapData(X - 1, y - 1).Luz < EB_LIMITE_INFERIOR Then
                    If InMapBounds(X + 1, y + 1) Then
                        If MapData(X + 1, y + 1).Luz <> 0 And MapData(X + 1, y + 1).Luz < EB_LIMITE_INFERIOR Then
                            Puede = 1 'Diagonal PURO
                        End If
                    End If
                End If
            End If
            If InMapBounds(X - 1, y - 1) And Puede = 0 Then
                If MapData(X - 1, y - 1).Luz <> 0 And MapData(X - 1, y - 1).Luz < EB_LIMITE_INFERIOR Then
                    If InMapBounds(X + 1, y) Then
                        If MapData(X + 1, y).Luz <> 0 And MapData(X + 1, y).Luz < EB_LIMITE_INFERIOR Then
                            Puede = 2 'Diagonal Izquierda, Horizontal Derecha
                        End If
                    End If
                End If
            End If
            If InMapBounds(X - 1, y - 1) And Puede = 0 Then
                If MapData(X - 1, y - 1).Luz <> 0 And MapData(X - 1, y - 1).Luz < EB_LIMITE_INFERIOR Then
                    If InMapBounds(X, y + 1) Then
                        If MapData(X, y + 1).Luz <> 0 And MapData(X, y + 1).Luz < EB_LIMITE_INFERIOR Then
                            Puede = 3 'Diagonal Izquierda, Vertical Derecha
                        End If
                    End If
                End If
            End If
            
            If InMapBounds(X + 1, y + 1) And Puede = 0 Then
                If MapData(X + 1, y + 1).Luz > 0 And MapData(X + 1, y + 1).Luz < EB_LIMITE_INFERIOR Then
                    If InMapBounds(X - 1, y) Then
                        If MapData(X - 1, y).Luz > 0 And MapData(X - 1, y).Luz < EB_LIMITE_INFERIOR Then
                            Puede = 4 'Diagonal derecha, Horizontal Izquierda
                        End If
                    End If
                End If
            End If
            If InMapBounds(X + 1, y + 1) And Puede = 0 Then
                If MapData(X + 1, y + 1).Luz > 0 And MapData(X + 1, y + 1).Luz < EB_LIMITE_INFERIOR Then
                    If InMapBounds(X, y - 1) Then
                        If MapData(X, y - 1).Luz > 0 And MapData(X, y - 1).Luz < EB_LIMITE_INFERIOR Then
                            Puede = 5 'Diagonal derecha, Vertical Izquierda.
                        End If
                    End If
                End If
            End If
            
            If InMapBounds(X + 1, y) And Puede = 0 Then
                If MapData(X + 1, y).Luz > 0 And MapData(X + 1, y).Luz < EB_LIMITE_INFERIOR Then
                    If InMapBounds(X - 1, y) Then
                        If MapData(X - 1, y).Luz > 0 And MapData(X - 1, y).Luz < EB_LIMITE_INFERIOR Then
                            Puede = 6 'Horizontal Derecha, Horizontal Izquierda
                        End If
                    End If
                End If
            End If
            If InMapBounds(X + 1, y) And Puede = 0 Then
                If MapData(X + 1, y).Luz > 0 And MapData(X + 1, y).Luz < EB_LIMITE_INFERIOR Then
                    If InMapBounds(X, y - 1) Then
                        If MapData(X, y - 1).Luz > 0 And MapData(X, y - 1).Luz < EB_LIMITE_INFERIOR Then
                            Puede = 7 'Horizontal Derecha, vertical Izquierda
                        End If
                    End If
                End If
            End If
            If InMapBounds(X, y + 1) And Puede = 0 Then
                If MapData(X, y + 1).Luz > 0 And MapData(X, y + 1).Luz < EB_LIMITE_INFERIOR Then
                    If InMapBounds(X, y - 1) Then
                        If MapData(X, y - 1).Luz > 0 And MapData(X, y - 1).Luz < EB_LIMITE_INFERIOR Then
                            Puede = 8 'vertical Derecha, vertical Izquierda
                        End If
                    End If
                End If
            End If
            
            If Puede = 1 Then
                MapData(X, y).Luz = eB_Light.DIHRCrossLeftUp
                
                MapData(X, y).light_value(0) = MapData(X - 1, y - 1).light_value(3)
                MapData(X, y).light_value(3) = MapData(X + 1, y).light_value(0)
                
                MapData(X, y).light_value(1) = 0
                MapData(X, y).light_value(2) = 0
                
                
                MapData(X, y).LV(0) = MapData(X - 1, y - 1).LV(3)
                MapData(X, y).LV(3) = MapData(X + 1, y).LV(0)
                
                MapData(X, y).LV(2) = 0
                MapData(X, y).LV(1) = 0
                
            
            ElseIf Puede = 2 Then
                MapData(X, y).Luz = eB_Light.DIVRCrossLeftUp
                
                MapData(X, y).light_value(0) = MapData(X - 1, y - 1).light_value(3)
                MapData(X, y).light_value(3) = MapData(X, y + 1).light_value(2)
                
                MapData(X, y).light_value(1) = 0
                MapData(X, y).light_value(2) = 0
                
                MapData(X, y).LV(0) = MapData(X - 1, y - 1).LV(3)
                MapData(X, y).LV(3) = MapData(X, y + 1).LV(2)
                
                MapData(X, y).LV(2) = 0
                MapData(X, y).LV(1) = 0
            
            ElseIf Puede = 3 Then
                MapData(X, y).Luz = eB_Light.CrossLeftUp
                
                MapData(X, y).light_value(0) = MapData(X - 1, y - 1).light_value(3)
                MapData(X, y).light_value(3) = MapData(X + 1, y + 1).light_value(1)
                
                MapData(X, y).light_value(1) = 0
                MapData(X, y).light_value(2) = 0
                
                
                
                MapData(X, y).LV(0) = MapData(X - 1, y - 1).LV(3)
                MapData(X, y).LV(3) = MapData(X + 1, y + 1).LV(1)
                
                MapData(X, y).LV(1) = 0
                MapData(X, y).LV(2) = 0
            
            ElseIf Puede = 4 Then
                MapData(X, y).Luz = eB_Light.VIDRCrossLeftUp
                
                MapData(X, y).light_value(0) = MapData(X - 1, y).light_value(1)
                MapData(X, y).light_value(3) = MapData(X + 1, y + 1).light_value(2)
                
                MapData(X, y).light_value(1) = 0
                MapData(X, y).light_value(2) = 0
                
                
                MapData(X, y).LV(0) = MapData(X - 1, y).LV(1)
                MapData(X, y).LV(3) = MapData(X + 1, y + 1).LV(2)
                
                MapData(X, y).LV(1) = 0
                MapData(X, y).LV(2) = 0
            
            ElseIf Puede = 5 Then
                MapData(X, y).Luz = eB_Light.VIDRCrossLeftUp
                
                MapData(X, y).light_value(0) = MapData(X, y - 1).light_value(2)
                MapData(X, y).light_value(3) = MapData(X + 1, y + 1).light_value(2)
                
                MapData(X, y).light_value(1) = 0
                MapData(X, y).light_value(2) = 0
            
                MapData(X, y).LV(0) = MapData(X, y - 1).LV(2)
                MapData(X, y).LV(3) = MapData(X + 1, y + 1).LV(2)
                
                MapData(X, y).LV(1) = 0
                MapData(X, y).LV(2) = 0
            ElseIf Puede = 6 Then
                MapData(X, y).Luz = eB_Light.HIHRCrossLeftUp
                
                MapData(X, y).light_value(0) = MapData(X - 1, y).light_value(1)
                MapData(X, y).light_value(3) = MapData(X + 1, y).light_value(2)
                
                MapData(X, y).light_value(1) = 0
                MapData(X, y).light_value(2) = 0
            
                MapData(X, y).LV(0) = MapData(X - 1, y).LV(1)
                MapData(X, y).LV(3) = MapData(X + 1, y).LV(2)
                
                MapData(X, y).LV(1) = 0
                MapData(X, y).LV(2) = 0
            
            ElseIf Puede = 7 Then
                MapData(X, y).Luz = eB_Light.VIHRCrossLeftUp
                
                MapData(X, y).light_value(0) = MapData(X, y - 1).light_value(2)
                MapData(X, y).light_value(3) = MapData(X + 1, y).light_value(2)
                
                MapData(X, y).light_value(1) = 0
                MapData(X, y).light_value(2) = 0
                
                
                MapData(X, y).LV(0) = MapData(X, y - 1).LV(2)
                MapData(X, y).LV(3) = MapData(X + 1, y).LV(2)
                
                MapData(X, y).LV(1) = 0
                MapData(X, y).LV(2) = 0
            
            ElseIf Puede = 8 Then
                MapData(X, y).Luz = eB_Light.VIVRCrossLeftUp
                
                MapData(X, y).light_value(0) = MapData(X, y - 1).light_value(2)
                MapData(X, y).light_value(3) = MapData(X, y + 1).light_value(1)
                
                MapData(X, y).light_value(1) = 0
                MapData(X, y).light_value(2) = 0
                
                MapData(X, y).LV(0) = MapData(X, y - 1).LV(2)
                MapData(X, y).LV(3) = MapData(X, y + 1).LV(1)
                
                MapData(X, y).LV(1) = 0
                MapData(X, y).LV(2) = 0
            
            End If
        Case eB_Light.CrossRightUp
            'Derecha Arriba, Izquierda ABAJO
            If InMapBounds(X + 1, y - 1) Then
                If MapData(X + 1, y - 1).Luz <> 0 And MapData(X + 1, y - 1).Luz < EB_LIMITE_INFERIOR Then
                    If InMapBounds(X - 1, y + 1) Then
                        If MapData(X - 1, y + 1).Luz <> 0 And MapData(X - 1, y + 1).Luz < EB_LIMITE_INFERIOR Then
                            Puede = 1 'Diagonal PURO
                        End If
                    End If
                End If
            End If
            If InMapBounds(X - 1, y + 1) And Puede = 0 Then
                If MapData(X - 1, y + 1).Luz <> 0 And MapData(X - 1, y + 1).Luz < EB_LIMITE_INFERIOR Then
                    If InMapBounds(X + 1, y) Then
                        If MapData(X + 1, y).Luz <> 0 And MapData(X + 1, y).Luz < EB_LIMITE_INFERIOR Then
                            Puede = 2 'Diagonal Izquierda, Horizontal Derecha
                        End If
                    End If
                End If
            End If
            If InMapBounds(X - 1, y + 1) And Puede = 0 Then
                If MapData(X - 1, y + 1).Luz <> 0 And MapData(X - 1, y + 1).Luz < EB_LIMITE_INFERIOR Then
                    If InMapBounds(X, y - 1) Then
                        If MapData(X, y - 1).Luz <> 0 And MapData(X, y - 1).Luz < EB_LIMITE_INFERIOR Then
                            Puede = 3 'Diagonal Izquierda, Vertical Derecha
                        End If
                    End If
                End If
            End If
            
            If InMapBounds(X + 1, y - 1) And Puede = 0 Then
                If MapData(X + 1, y - 1).Luz > 0 And MapData(X + 1, y - 1).Luz < EB_LIMITE_INFERIOR Then
                    If InMapBounds(X - 1, y) Then
                        If MapData(X - 1, y).Luz > 0 And MapData(X - 1, y).Luz < EB_LIMITE_INFERIOR Then
                            Puede = 4 'Diagonal derecha, Horizontal Izquierda
                        End If
                    End If
                End If
            End If
            If InMapBounds(X + 1, y - 1) And Puede = 0 Then
                If MapData(X + 1, y - 1).Luz > 0 And MapData(X + 1, y - 1).Luz < EB_LIMITE_INFERIOR Then
                    If InMapBounds(X, y + 1) Then
                        If MapData(X, y + 1).Luz > 0 And MapData(X, y + 1).Luz < EB_LIMITE_INFERIOR Then
                            Puede = 5 'Diagonal derecha, Vertical Izquierda.
                        End If
                    End If
                End If
            End If
            
            If InMapBounds(X + 1, y) And Puede = 0 Then
                If MapData(X + 1, y).Luz > 0 And MapData(X + 1, y).Luz < EB_LIMITE_INFERIOR Then
                    If InMapBounds(X - 1, y) Then
                        If MapData(X - 1, y).Luz > 0 And MapData(X - 1, y).Luz < EB_LIMITE_INFERIOR Then
                            Puede = 6 'Horizontal Derecha, Horizontal Izquierda
                        End If
                    End If
                End If
            End If
            If InMapBounds(X + 1, y) And Puede = 0 Then
                If MapData(X + 1, y).Luz > 0 And MapData(X + 1, y).Luz < EB_LIMITE_INFERIOR Then
                    If InMapBounds(X, y + 1) Then
                        If MapData(X, y + 1).Luz > 0 And MapData(X, y + 1).Luz < EB_LIMITE_INFERIOR Then
                            Puede = 7 'Horizontal Derecha, vertical Izquierda
                        End If
                    End If
                End If
            End If
            If InMapBounds(X, y + 1) And Puede = 0 Then
                If MapData(X, y + 1).Luz > 0 And MapData(X, y + 1).Luz < EB_LIMITE_INFERIOR Then
                    If InMapBounds(X, y - 1) Then
                        If MapData(X, y - 1).Luz > 0 And MapData(X, y - 1).Luz < EB_LIMITE_INFERIOR Then
                            Puede = 8 'vertical Derecha, vertical Izquierda
                        End If
                    End If
                End If
            End If
            
            If Puede = 1 Then
                MapData(X, y).Luz = eB_Light.CrossRightUp
                
                MapData(X, y).light_value(0) = 0
                MapData(X, y).light_value(3) = 0
                
                MapData(X, y).light_value(1) = MapData(X + 1, y - 1).light_value(2)
                MapData(X, y).light_value(2) = MapData(X - 1, y + 1).light_value(1)
                
            
                MapData(X, y).LV(0) = 0
                MapData(X, y).LV(3) = 0
                
                MapData(X, y).LV(1) = MapData(X + 1, y - 1).LV(2)
                MapData(X, y).LV(2) = MapData(X - 1, y + 1).LV(1)
            
            ElseIf Puede = 2 Then
                MapData(X, y).Luz = eB_Light.HRDICrossRightUp
                
                MapData(X, y).light_value(0) = 0
                MapData(X, y).light_value(3) = 0
                
                MapData(X, y).light_value(1) = MapData(X + 1, y).light_value(0)
                MapData(X, y).light_value(2) = MapData(X - 1, y + 1).light_value(1)
                
                
                MapData(X, y).LV(0) = 0
                MapData(X, y).LV(3) = 0
                
                MapData(X, y).LV(1) = MapData(X + 1, y).LV(0)
                MapData(X, y).LV(2) = MapData(X - 1, y + 1).LV(1)
                
            ElseIf Puede = 3 Then
                MapData(X, y).Luz = eB_Light.VRDICrossRightUp
                
                MapData(X, y).light_value(0) = 0
                MapData(X, y).light_value(3) = 0
                
                MapData(X, y).light_value(1) = MapData(X, y - 1).light_value(3)
                MapData(X, y).light_value(2) = MapData(X - 1, y + 1).light_value(1)
                
                MapData(X, y).LV(0) = 0
                MapData(X, y).LV(3) = 0
                
                MapData(X, y).LV(1) = MapData(X, y - 1).LV(3)
                MapData(X, y).LV(2) = MapData(X - 1, y + 1).LV(1)
            ElseIf Puede = 4 Then
                MapData(X, y).Luz = eB_Light.DRHICrossRightUp
                
                MapData(X, y).light_value(0) = 0
                MapData(X, y).light_value(3) = 0
                
                MapData(X, y).light_value(1) = MapData(X + 1, y - 1).light_value(2)
                MapData(X, y).light_value(2) = MapData(X - 1, y).light_value(3)
            
                MapData(X, y).LV(0) = 0
                MapData(X, y).LV(3) = 0
                
                MapData(X, y).LV(1) = MapData(X + 1, y - 1).LV(2)
                MapData(X, y).LV(2) = MapData(X - 1, y).LV(3)
            
            ElseIf Puede = 5 Then
                MapData(X, y).Luz = eB_Light.DRVICrossRightUp
                
                MapData(X, y).light_value(0) = 0
                MapData(X, y).light_value(3) = 0
                
                MapData(X, y).light_value(1) = MapData(X + 1, y - 1).light_value(2)
                MapData(X, y).light_value(2) = MapData(X, y + 1).light_value(0)
                            
                MapData(X, y).LV(0) = 0
                MapData(X, y).LV(3) = 0
                
                MapData(X, y).LV(1) = MapData(X + 1, y - 1).LV(2)
                MapData(X, y).LV(2) = MapData(X, y + 1).LV(0)
            
            ElseIf Puede = 6 Then
                MapData(X, y).Luz = eB_Light.HRHICrossRightUp
                
                MapData(X, y).light_value(0) = 0
                MapData(X, y).light_value(3) = 0
                
                MapData(X, y).light_value(1) = MapData(X + 1, y).light_value(0)
                MapData(X, y).light_value(2) = MapData(X - 1, y).light_value(3)
            
                MapData(X, y).LV(0) = 0
                MapData(X, y).LV(3) = 0
                
                MapData(X, y).LV(1) = MapData(X + 1, y).LV(0)
                MapData(X, y).LV(2) = MapData(X - 1, y).LV(3)
            
            ElseIf Puede = 7 Then
          
            ElseIf Puede = 8 Then
            
            End If
        
    'HASTA ACA ESTA HECHO COMPLETO. COMO MUCHO HABRIA QUE VER CASOS DE VERTICALES Y HORIZONTALES
    'QUE TOMARAN SU VALOR A PARTIR DE OTRAS COSAS Q NO SEAN JUSTAMENTE ESOS LIMITES.
    'EN ADELANTE HAGO SOLO LOS RESUMIDOS.
        Case eB_Light.NotUpperLeft
            If InMapBounds(X + 1, y - 1) Then
                If MapData(X + 1, y - 1).Luz > 0 And MapData(X + 1, y - 1).Luz < EB_LIMITE_INFERIOR Then
                    Puede = 1
                End If
            End If
            If InMapBounds(X - 1, y + 1) Then
                If MapData(X - 1, y + 1).Luz > 0 And MapData(X - 1, y + 1).Luz < EB_LIMITE_INFERIOR Then
                    Puede = Puede + 2
                End If
            End If
            If InMapBounds(X + 1, y + 1) Then
                If MapData(X + 1, y + 1).Luz > 0 And MapData(X + 1, y + 1).Luz < EB_LIMITE_INFERIOR Then
                    Puede = Puede + 3
                End If
            End If
            
            If Puede = 6 Then 'Clasicos diagonales
                MapData(X, y).Luz = eB_Light.NotUpperLeft
                
                MapData(X, y).light_value(0) = 0
                
                MapData(X, y).light_value(1) = MapData(X + 1, y - 1).light_value(2)
                MapData(X, y).light_value(2) = MapData(X - 1, y + 1).light_value(1)
                MapData(X, y).light_value(3) = MapData(X + 1, y + 1).light_value(0)
                
                MapData(X, y).LV(0) = 0
                MapData(X, y).LV(1) = MapData(X + 1, y - 1).LV(2)
                MapData(X, y).LV(2) = MapData(X - 1, y + 1).LV(1)
                MapData(X, y).LV(3) = MapData(X + 1, y + 1).LV(0)
            End If
        Case eB_Light.NotUpperRight
            If InMapBounds(X - 1, y - 1) Then
                If MapData(X - 1, y - 1).Luz > 0 And MapData(X - 1, y - 1).Luz < EB_LIMITE_INFERIOR Then
                    Puede = 1
                End If
            End If
            If InMapBounds(X - 1, y + 1) Then
                If MapData(X - 1, y + 1).Luz > 0 And MapData(X - 1, y + 1).Luz < EB_LIMITE_INFERIOR Then
                    Puede = Puede + 2
                End If
            End If
            If InMapBounds(X + 1, y + 1) Then
                If MapData(X + 1, y + 1).Luz > 0 And MapData(X + 1, y + 1).Luz < EB_LIMITE_INFERIOR Then
                    Puede = Puede + 3
                End If
            End If
            
            If Puede = 6 Then 'Clasicos diagonales
                MapData(X, y).Luz = eB_Light.NotUpperRight
                
                MapData(X, y).light_value(0) = MapData(X - 1, y - 1).light_value(3)
                MapData(X, y).light_value(1) = 0
                MapData(X, y).light_value(2) = MapData(X - 1, y + 1).light_value(1)
                MapData(X, y).light_value(3) = MapData(X + 1, y + 1).light_value(0)
                

                MapData(X, y).LV(0) = MapData(X - 1, y - 1).LV(3)
                MapData(X, y).LV(1) = 0
                MapData(X, y).LV(2) = MapData(X - 1, y + 1).LV(2)
                MapData(X, y).LV(3) = MapData(X + 1, y + 1).LV(1)
                
            End If
        Case eB_Light.NotBottomLeft
            If InMapBounds(X - 1, y - 1) Then
                If MapData(X - 1, y - 1).Luz > 0 And MapData(X - 1, y - 1).Luz < EB_LIMITE_INFERIOR Then
                    Puede = 1
                End If
            End If
            If InMapBounds(X + 1, y - 1) Then
                If MapData(X + 1, y - 1).Luz > 0 And MapData(X + 1, y - 1).Luz < EB_LIMITE_INFERIOR Then
                    Puede = Puede + 2
                End If
            End If
            If InMapBounds(X + 1, y + 1) Then
                If MapData(X + 1, y + 1).Luz > 0 And MapData(X + 1, y + 1).Luz < EB_LIMITE_INFERIOR Then
                    Puede = Puede + 3
                End If
            End If
            
            If Puede = 6 Then 'Clasicos diagonales
                MapData(X, y).Luz = eB_Light.NotUpperLeft
                
                MapData(X, y).light_value(0) = MapData(X - 1, y - 1).light_value(3)
                MapData(X, y).light_value(1) = MapData(X + 1, y - 1).light_value(2)
                MapData(X, y).light_value(2) = 0
                MapData(X, y).light_value(3) = MapData(X + 1, y + 1).light_value(0)
                
                MapData(X, y).LV(0) = MapData(X - 1, y - 1).LV(3)
                MapData(X, y).LV(1) = MapData(X + 1, y - 1).LV(2)
                MapData(X, y).LV(2) = 0
                MapData(X, y).LV(3) = MapData(X + 1, y + 1).LV(0)
                
            End If
        Case eB_Light.NotBottomRight
            If InMapBounds(X - 1, y - 1) Then
                If MapData(X - 1, y - 1).Luz > 0 And MapData(X - 1, y - 1).Luz < EB_LIMITE_INFERIOR Then
                    Puede = 1
                End If
            End If
            If InMapBounds(X - 1, y + 1) Then
                If MapData(X - 1, y + 1).Luz > 0 And MapData(X - 1, y + 1).Luz < EB_LIMITE_INFERIOR Then
                    Puede = Puede + 2
                End If
            End If
            If InMapBounds(X + 1, y - 1) Then
                If MapData(X + 1, y - 1).Luz > 0 And MapData(X + 1, y - 1).Luz < EB_LIMITE_INFERIOR Then
                    Puede = Puede + 3
                End If
            End If
            
            If Puede = 6 Then 'Clasicos diagonales
                MapData(X, y).Luz = eB_Light.NotUpperLeft
                
                MapData(X, y).light_value(0) = MapData(X - 1, y - 1).light_value(3)
                MapData(X, y).light_value(1) = MapData(X + 1, y - 1).light_value(2)
                MapData(X, y).light_value(2) = MapData(X - 1, y + 1).light_value(1)
                MapData(X, y).light_value(3) = 0
            
                MapData(X, y).LV(0) = MapData(X - 1, y - 1).LV(3)
                MapData(X, y).LV(1) = MapData(X + 1, y - 1).LV(2)
                MapData(X, y).LV(2) = MapData(X - 1, y + 1).LV(1)
                MapData(X, y).LV(3) = 0
            
            End If
        Case eB_Light.AllCorner
            If InMapBounds(X - 1, y - 1) Then
                If MapData(X - 1, y - 1).Luz > 0 And MapData(X - 1, y - 1).Luz < EB_LIMITE_INFERIOR Then
                    Puede = 1
                ElseIf MapData(X - 1, y - 1).Luz > 0 And MapData(X - 1, y - 1).light_value(3) <> 0 Then
                    Puede = 1
                End If
            End If
            If InMapBounds(X - 1, y + 1) Then
                If MapData(X - 1, y + 1).Luz > 0 And MapData(X - 1, y + 1).Luz < EB_LIMITE_INFERIOR Then
                    Puede = Puede + 2
                ElseIf MapData(X - 1, y + 1).Luz > 0 And MapData(X - 1, y + 1).light_value(1) Then
                    Puede = Puede + 2
                End If
            End If
            If InMapBounds(X + 1, y + 1) Then
                If MapData(X + 1, y + 1).Luz > 0 And MapData(X + 1, y + 1).Luz < EB_LIMITE_INFERIOR Then
                    Puede = Puede + 3
                ElseIf MapData(X + 1, y + 1).Luz > 0 And MapData(X + 1, y + 1).light_value(0) Then
                    Puede = Puede + 3
                End If
            End If
            If InMapBounds(X + 1, y - 1) Then
                If MapData(X + 1, y - 1).Luz > 0 And MapData(X + 1, y - 1).Luz < EB_LIMITE_INFERIOR Then
                    Puede = Puede + 4
                ElseIf MapData(X + 1, y - 1).Luz > 0 And MapData(X + 1, y - 1).light_value(2) Then
                    Puede = Puede + 4
                End If
            End If
            
            If Puede = 10 Then 'Clasicos diagonales
                MapData(X, y).Luz = eB_Light.AllCorner
                MapData(X, y).light_value(0) = MapData(X - 1, y - 1).light_value(3)
                MapData(X, y).light_value(1) = MapData(X + 1, y - 1).light_value(2)
                MapData(X, y).light_value(2) = MapData(X - 1, y + 1).light_value(1)
                MapData(X, y).light_value(3) = MapData(X + 1, y + 1).light_value(0)
                
                
                MapData(X, y).LV(0) = MapData(X - 1, y - 1).LV(3)
                MapData(X, y).LV(1) = MapData(X + 1, y - 1).LV(2)
                MapData(X, y).LV(2) = MapData(X - 1, y + 1).LV(1)
                MapData(X, y).LV(3) = MapData(X + 1, y + 1).LV(0)
            End If
    End Select
End If
End Sub

Public Sub AplicarBorde(ByVal X As Byte, ByVal y As Byte)
Dim ul As Boolean
Dim ur As Boolean
Dim bl As Boolean
Dim br As Boolean
Dim lC As Byte
Dim SameColor As Boolean
Dim Color0 As Long
Dim Color1 As Long
Dim Color2 As Long
Dim Color3 As Long
Dim OldL As Byte




If frmMain.cHorizontal.value Then
    AplicarBordeManual X, y, 0
    Exit Sub
ElseIf frmMain.cVertical.value Then
    AplicarBordeManual X, y, 1
    Exit Sub
ElseIf frmMain.cUL.value Then
    AplicarBordeManual X, y, 2
    Exit Sub
ElseIf frmMain.cUR.value Then
    AplicarBordeManual X, y, 3
    Exit Sub
ElseIf frmMain.cBL.value Then
    AplicarBordeManual X, y, 4
    Exit Sub
ElseIf frmMain.cBR.value Then
    AplicarBordeManual X, y, 5
    Exit Sub
ElseIf frmMain.cCROSSUR.value Then
    AplicarBordeManual X, y, 6
    Exit Sub
ElseIf frmMain.cCROSSUL.value Then
    AplicarBordeManual X, y, 7
    Exit Sub
ElseIf frmMain.cNotUL.value Then
    AplicarBordeManual X, y, 8
    Exit Sub
ElseIf frmMain.cNotUR.value Then
    AplicarBordeManual X, y, 9
    Exit Sub
ElseIf frmMain.cNotBL.value Then
    AplicarBordeManual X, y, 10
    Exit Sub
ElseIf frmMain.cNotBR.value Then
    AplicarBordeManual X, y, 11
    Exit Sub
ElseIf frmMain.cALLC.value Then
    AplicarBordeManual X, y, 12
    Exit Sub
End If

    OldL = MapData(X, y).Luz
    'Insertamos un borde en el TILE.
    
    'Si son limites horizontales o verticales lo miramos primero...
    If InMapBounds(X + 1, y) Then
        If MapData(X + 1, y).Luz <> 0 Then
            If MapData(X + 1, y).light_value(0) = MapData(X + 1, y).light_value(2) Then
                'Mismo limite vertical
                If MapData(X - 1, y).light_value(1) = MapData(X - 1, y).light_value(3) Then
                    MapData(X, y).Luz = eB_Light.Vertical
                    MapData(X, y).light_value(0) = MapData(X - 1, y).light_value(1)
                    MapData(X, y).light_value(2) = MapData(X - 1, y).light_value(3)
                    MapData(X, y).light_value(1) = MapData(X + 1, y).light_value(0)
                    MapData(X, y).light_value(3) = MapData(X + 1, y).light_value(2)
                    MapData(X, y).LV(0) = MapData(X - 1, y).LV(1)
                    MapData(X, y).LV(2) = MapData(X - 1, y).LV(3)
                    MapData(X, y).LV(1) = MapData(X + 1, y).LV(0)
                    MapData(X, y).LV(3) = MapData(X + 1, y).LV(2)
                    
                    
                    Exit Sub
                Else 'Aca habria que hacer un VERTICAL CON DIF CORNER.
                End If
            End If
        End If
    End If
    If InMapBounds(X - 1, y) Then
        If MapData(X - 1, y).Luz <> 0 Then
            If MapData(X - 1, y).light_value(1) = MapData(X - 1, y).light_value(3) Then
                'Mismo limite vertical
                If MapData(X + 1, y).light_value(0) = MapData(X + 1, y).light_value(2) Then
                    MapData(X, y).Luz = eB_Light.Vertical
                    MapData(X, y).light_value(0) = MapData(X - 1, y).light_value(1)
                    MapData(X, y).light_value(2) = MapData(X - 1, y).light_value(3)
                    MapData(X, y).light_value(1) = MapData(X + 1, y).light_value(0)
                    MapData(X, y).light_value(3) = MapData(X + 1, y).light_value(2)
                    MapData(X, y).LV(0) = MapData(X - 1, y).LV(1)
                    MapData(X, y).LV(2) = MapData(X - 1, y).LV(3)
                    MapData(X, y).LV(1) = MapData(X + 1, y).LV(0)
                    MapData(X, y).LV(3) = MapData(X + 1, y).LV(2)
                    Exit Sub
                Else 'Aca habria que hacer un VERTICAL CON DIF CORNER.
                End If
            End If
        End If
    End If
    
    If InMapBounds(X, y + 1) Then
        If MapData(X, y + 1).Luz <> 0 Then
            If MapData(X, y + 1).light_value(0) = MapData(X, y + 1).light_value(1) Then
                'Mismo limite vertical
                If MapData(X, y - 1).light_value(2) = MapData(X, y - 1).light_value(3) Then
                    MapData(X, y).Luz = eB_Light.Horizontal
                    MapData(X, y).light_value(0) = MapData(X, y - 1).light_value(2)
                    MapData(X, y).light_value(1) = MapData(X, y - 1).light_value(3)
                    MapData(X, y).light_value(2) = MapData(X, y + 1).light_value(0)
                    MapData(X, y).light_value(3) = MapData(X, y + 1).light_value(1)
                        MapData(X, y).LV(0) = MapData(X, y - 1).LV(2)
                    MapData(X, y).LV(1) = MapData(X, y - 1).LV(3)
                    MapData(X, y).LV(2) = MapData(X, y + 1).LV(0)
                    MapData(X, y).LV(3) = MapData(X, y + 1).LV(1)
                    Exit Sub
                Else 'Aca habria que hacer un VERTICAL CON DIF CORNER.
                End If
            End If
        End If
    End If
        If InMapBounds(X, y - 1) Then
        If MapData(X, y - 1).Luz <> 0 Then
            If MapData(X, y - 1).light_value(2) = MapData(X, y - 1).light_value(3) Then
                'Mismo limite vertical
                If MapData(X, y + 1).light_value(2) = MapData(X, y + 1).light_value(0) Then
                    MapData(X, y).Luz = eB_Light.Horizontal
                    MapData(X, y).light_value(0) = MapData(X, y - 1).light_value(2)
                    MapData(X, y).light_value(1) = MapData(X, y - 1).light_value(3)
                    MapData(X, y).light_value(2) = MapData(X, y + 1).light_value(0)
                    MapData(X, y).light_value(3) = MapData(X, y + 1).light_value(1)
                    MapData(X, y).LV(0) = MapData(X, y - 1).LV(2)
                    MapData(X, y).LV(1) = MapData(X, y - 1).LV(3)
                    MapData(X, y).LV(2) = MapData(X, y + 1).LV(0)
                    MapData(X, y).LV(3) = MapData(X, y + 1).LV(1)
                    Exit Sub
                Else 'Aca habria que hacer un VERTICAL CON DIF CORNER.
                End If
            End If
        End If
    End If
    
    
    
    
    
    'Chequeamos el Vertice UPPERLEFT.
    'Para eso miramos el BOTTOM RIGHT del X-1 Y-1
    If InMapBounds(X - 1, y - 1) Then
        If MapData(X - 1, y - 1).Luz <> 0 Then
        If MapData(X - 1, y - 1).light_value(3) <> 0 Then
            'El bottomright del x-1y-1 esta iluminado then copiamos en el upperleft
            MapData(X, y).light_value(0) = MapData(X - 1, y - 1).light_value(3)
            MapData(X, y).LV(0) = MapData(X - 1, y - 1).LV(3)
            ul = True
            lC = lC + 1
            Color0 = MapData(X - 1, y - 1).light_value(3)
            End If
        End If
    End If
    
    'Chequeamos el UpperRight
    If InMapBounds(X + 1, y - 1) Then
    If MapData(X + 1, y - 1).Luz <> 0 Then
        'If MapData(X + 1, y - 1).light_value(3) <> 0 Then
        If MapData(X + 1, y - 1).light_value(2) <> 0 Then
            MapData(X, y).light_value(1) = MapData(X + 1, y - 1).light_value(2)
            MapData(X, y).LV(1) = MapData(X + 1, y - 1).LV(2)
            ur = True
            lC = lC + 1
            Color1 = MapData(X + 1, y - 1).light_value(2)
        End If
        End If
    End If
    
    
    'Chequeamos el BottomLeft
    If InMapBounds(X - 1, y + 1) Then
            If MapData(X - 1, y + 1).Luz <> 0 Then
        If MapData(X - 1, y + 1).light_value(1) <> 0 Then
             MapData(X, y).LV(2) = MapData(X - 1, y + 1).LV(1)
            MapData(X, y).light_value(2) = MapData(X - 1, y + 1).light_value(1)
            bl = True
            lC = lC + 1
            Color2 = MapData(X - 1, y + 1).light_value(1)
            End If
        End If
    End If
    
    
    'Chequeamos el BottomRight
    If InMapBounds(X + 1, y + 1) Then
        If MapData(X + 1, y + 1).Luz <> 0 Then
        If MapData(X + 1, y + 1).light_value(0) <> 0 Then
            MapData(X, y).light_value(3) = MapData(X + 1, y + 1).light_value(0)
            MapData(X, y).LV(3) = MapData(X + 1, y + 1).LV(0)
            br = True
            lC = lC + 1
            Color3 = MapData(X + 1, y + 1).light_value(0)
            End If
        End If
    End If
    
    
    If Not br Then MapData(X, y).light_value(3) = 0
    If Not bl Then MapData(X, y).light_value(2) = 0
    If Not ul Then MapData(X, y).light_value(0) = 0
    If Not ur Then MapData(X, y).light_value(1) = 0
    If Not br Then MapData(X, y).LV(3) = 0
    If Not bl Then MapData(X, y).LV(2) = 0
    If Not ul Then MapData(X, y).LV(0) = 0
    If Not ur Then MapData(X, y).LV(1) = 0

    
    
    
    'Dilucidamos que "luz" aplicamos
    If lC = 1 Then
        'Solo un Corner
        If ul Then
            MapData(X, y).Luz = eB_Light.UpperLeft
        ElseIf ur Then
            MapData(X, y).Luz = eB_Light.UpperRight
        ElseIf bl Then
            MapData(X, y).Luz = eB_Light.BottomLeft
        ElseIf br Then
            MapData(X, y).Luz = eB_Light.BottomRight
        End If
        
    
    
    
        
    ElseIf lC = 2 Then
        'Bordes verticales u horizontales
        If ul And ur And (Color0 = Color1 And Color2 = Color3) Then
            'Borde Horizontal Superior.
            MapData(X, y).Luz = eB_Light.Horizontal
        ElseIf bl And br And (Color0 = Color1 And Color2 = Color3) Then
            'Borde Horizontal Inferior
            MapData(X, y).Luz = eB_Light.Horizontal
        End If
        
        If ul And bl And (Color0 = Color2 And Color1 = Color3) Then
            'Borde Vertical Izquierdo
            MapData(X, y).Luz = eB_Light.Vertical
        ElseIf ur And br And (Color0 = Color2 And Color1 = Color3) Then
            'Borde Vertical Derecho
            MapData(X, y).Luz = eB_Light.Vertical
        End If
        
        'Cruzados
        If ul And br Then
                    MapData(X, y).Luz = eB_Light.CrossLeftUp
        ElseIf ur And bl Then
                    MapData(X, y).Luz = eB_Light.CrossRightUp
        
        End If
        
        If OldL = MapData(X, y).Luz Then
            ' No entro en ningun IF CLAUSE, por ahora se me ocurre que es un limite distinto.
            If ul And ur Then
                MapData(X, y).Luz = eB_Light.AllCorner
            ElseIf bl And br Then
                 MapData(X, y).Luz = eB_Light.AllCorner
            ElseIf ul And bl Then
                  MapData(X, y).Luz = eB_Light.AllCorner
            
            ElseIf ur And br Then
            
                   MapData(X, y).Luz = eB_Light.AllCorner
            End If
        End If
        
    ElseIf lC = 3 Then
        'Alguno no esta siendo utilizado...lo buscamos
        
        'NotUL
        If Not ul Then
                    MapData(X, y).Luz = eB_Light.NotUpperLeft
        ElseIf Not ur Then
                    MapData(X, y).Luz = eB_Light.NotUpperRight
        ElseIf Not bl Then
                        MapData(X, y).Luz = eB_Light.NotBottomLeft
        ElseIf Not br Then
                        MapData(X, y).Luz = eB_Light.NotBottomRight
        End If
    ElseIf lC = 4 Then
        'Todos los Corner
        'Es un ALL CORNER, hay que ver que en realidad no sea un limite horizontal entre dos luces.
        
        
        If (MapData(X, y - 1).Luz > 0 And MapData(X, y - 1).Luz < 243) And (MapData(X, y + 1).Luz > 0 And MapData(X, y + 1).Luz < 243) Then
            
            If Color0 = Color1 And Color2 = Color3 Then
                MapData(X, y).Luz = eB_Light.Horizontal
            Else
                MapData(X, y).Luz = eB_Light.AllCorner
            End If
        Else
                 
        MapData(X, y).Luz = eB_Light.AllCorner
        End If
    
    End If
        
        
    
    


End Sub
Public Sub AplicarLuz(ByVal X As Byte, ByVal y As Byte, ByVal Luz As Byte, ByVal Rango As Byte, ByVal Borde As Byte)

Dim nX As Byte
Dim Xx As Byte
Dim nY As Byte
Dim xY As Byte

Dim lx As Byte ' Looper
Dim ly As Byte 'Looper

'Calculamos extremos tiles
nX = X - Rango
Xx = X + Rango

nY = y - Rango
xY = y + Rango

'Verificamos que este en el mapa.

If nX < XMinMapSize Then
    nX = XMinMapSize
End If
If Xx > XMaxMapSize Then
    Xx = XMaxMapSize
End If
If nY < YMinMapSize Then
    nY = YMinMapSize
End If
If xY > YMaxMapSize Then
    xY = YMaxMapSize
End If

'Si esta fuera de rango lo reducimos.

If Rango = 0 And (frmMain.cCROSSUR.value Or frmMain.cCROSSUL.value Or frmMain.cVertical.value Or frmMain.cHorizontal Or frmMain.cBR Or frmMain.cUL Or frmMain.cUR Or frmMain.cBL Or frmMain.cNotBL Or frmMain.cNotBR Or frmMain.cNotUL Or frmMain.cNotUR) Then
    If frmMain.cCROSSUL Then

    MapData(X, y).Luz = eB_Light.DIAGONALUL
            
                MapData(X, y).light_value(1) = DAMELONGLUZ(Luz)
                MapData(X, y).light_value(2) = DAMELONGLUZ(Luz)
                
                MapData(X, y).light_value(0) = DAMELONGLUZ(0)
                MapData(X, y).light_value(3) = DAMELONGLUZ(0)
                
                
                MapData(X, y).LV(1) = Luz
                MapData(X, y).LV(2) = Luz
                
                MapData(X, y).LV(3) = 0
                MapData(X, y).LV(0) = 0

        ElseIf frmMain.cUL Then
            If frmMain.cINV Then
            MapData(X, y).Luz = eB_Light.iUpperLeft
    
                MapData(X, y).light_value(1) = DAMELONGLUZ(Luz)
                MapData(X, y).light_value(2) = DAMELONGLUZ(0)
                
                MapData(X, y).light_value(0) = DAMELONGLUZ(0)
                MapData(X, y).light_value(3) = DAMELONGLUZ(0)
                
                
                MapData(X, y).LV(1) = Luz
                MapData(X, y).LV(2) = 0
                
                MapData(X, y).LV(3) = 0
                MapData(X, y).LV(0) = 0
            Else
                        MapData(X, y).Luz = eB_Light.UpperLeft
    
                MapData(X, y).light_value(1) = DAMELONGLUZ(0)
                MapData(X, y).light_value(2) = DAMELONGLUZ(0)
                
                MapData(X, y).light_value(0) = DAMELONGLUZ(Luz)
                MapData(X, y).light_value(3) = DAMELONGLUZ(0)
                
                
                MapData(X, y).LV(1) = 0
                MapData(X, y).LV(2) = 0
                
                MapData(X, y).LV(3) = 0
                MapData(X, y).LV(0) = Luz
            End If
        ElseIf frmMain.cNotUL Then
            If frmMain.cINV Then
            MapData(X, y).Luz = eB_Light.iNotUpperLeft
    
                MapData(X, y).light_value(1) = DAMELONGLUZ(0)
                MapData(X, y).light_value(2) = DAMELONGLUZ(Luz)
                
                MapData(X, y).light_value(0) = DAMELONGLUZ(Luz)
                MapData(X, y).light_value(3) = DAMELONGLUZ(Luz)
                
                
                MapData(X, y).LV(1) = 0
                MapData(X, y).LV(2) = Luz
                
                MapData(X, y).LV(3) = Luz
                MapData(X, y).LV(0) = Luz
            Else
                MapData(X, y).Luz = eB_Light.NotUpperLeft
    
                MapData(X, y).light_value(1) = DAMELONGLUZ(Luz)
                MapData(X, y).light_value(2) = DAMELONGLUZ(Luz)
                
                MapData(X, y).light_value(0) = DAMELONGLUZ(0)
                MapData(X, y).light_value(3) = DAMELONGLUZ(Luz)
                
                
                MapData(X, y).LV(1) = Luz
                MapData(X, y).LV(2) = Luz
                
                MapData(X, y).LV(3) = Luz
                MapData(X, y).LV(0) = 0
            End If
        ElseIf frmMain.cNotUR Then
            If frmMain.cINV Then
            MapData(X, y).Luz = eB_Light.iNotUpperRight
    
                MapData(X, y).light_value(1) = DAMELONGLUZ(Luz)
                MapData(X, y).light_value(2) = DAMELONGLUZ(Luz)
                
                MapData(X, y).light_value(0) = DAMELONGLUZ(Luz)
                MapData(X, y).light_value(3) = DAMELONGLUZ(0)
                
                
                MapData(X, y).LV(1) = Luz
                MapData(X, y).LV(2) = Luz
                
                MapData(X, y).LV(3) = 0
                MapData(X, y).LV(0) = Luz
            Else
                MapData(X, y).Luz = eB_Light.NotUpperRight
    
                MapData(X, y).light_value(1) = DAMELONGLUZ(0)
                MapData(X, y).light_value(2) = DAMELONGLUZ(Luz)
                
                MapData(X, y).light_value(0) = DAMELONGLUZ(Luz)
                MapData(X, y).light_value(3) = DAMELONGLUZ(Luz)
                
                
                MapData(X, y).LV(1) = 0
                MapData(X, y).LV(2) = Luz
                
                MapData(X, y).LV(3) = Luz
                MapData(X, y).LV(0) = Luz
            End If
        ElseIf frmMain.cNotBR Then
            If frmMain.cINV Then
            MapData(X, y).Luz = eB_Light.iNotBottomRight
    
                MapData(X, y).light_value(1) = DAMELONGLUZ(Luz)
                MapData(X, y).light_value(2) = DAMELONGLUZ(0)
                
                MapData(X, y).light_value(0) = DAMELONGLUZ(Luz)
                MapData(X, y).light_value(3) = DAMELONGLUZ(Luz)
                
                
                MapData(X, y).LV(1) = Luz
                MapData(X, y).LV(2) = 0
                
                MapData(X, y).LV(3) = Luz
                MapData(X, y).LV(0) = Luz
            Else
                MapData(X, y).Luz = eB_Light.NotBottomRight
    
                MapData(X, y).light_value(1) = DAMELONGLUZ(Luz)
                MapData(X, y).light_value(2) = DAMELONGLUZ(Luz)
                
                MapData(X, y).light_value(0) = DAMELONGLUZ(Luz)
                MapData(X, y).light_value(3) = DAMELONGLUZ(0)
                
                
                MapData(X, y).LV(1) = Luz
                MapData(X, y).LV(2) = Luz
                
                MapData(X, y).LV(3) = 0
                MapData(X, y).LV(0) = Luz
            End If
        ElseIf frmMain.cNotBL Then
            If frmMain.cINV Then
            MapData(X, y).Luz = eB_Light.iNotBottomLeft
    
                MapData(X, y).light_value(1) = DAMELONGLUZ(Luz)
                MapData(X, y).light_value(2) = DAMELONGLUZ(Luz)
                
                MapData(X, y).light_value(0) = DAMELONGLUZ(0)
                MapData(X, y).light_value(3) = DAMELONGLUZ(Luz)
                
                
                MapData(X, y).LV(1) = Luz
                MapData(X, y).LV(2) = Luz
                
                MapData(X, y).LV(3) = Luz
                MapData(X, y).LV(0) = 0
            Else
                MapData(X, y).Luz = eB_Light.NotBottomLeft
    
                MapData(X, y).light_value(1) = DAMELONGLUZ(Luz)
                MapData(X, y).light_value(2) = DAMELONGLUZ(0)
                
                MapData(X, y).light_value(0) = DAMELONGLUZ(Luz)
                MapData(X, y).light_value(3) = DAMELONGLUZ(Luz)
                
                
                MapData(X, y).LV(1) = Luz
                MapData(X, y).LV(2) = 0
                
                MapData(X, y).LV(3) = Luz
                MapData(X, y).LV(0) = Luz
            End If
        ElseIf frmMain.cUR Then
            If frmMain.cINV Then
                   MapData(X, y).Luz = eB_Light.UpperRight
    
                MapData(X, y).light_value(1) = DAMELONGLUZ(Luz)
                MapData(X, y).light_value(2) = DAMELONGLUZ(0)
                
                MapData(X, y).light_value(0) = DAMELONGLUZ(0)
                MapData(X, y).light_value(3) = DAMELONGLUZ(0)
                
                
                MapData(X, y).LV(1) = Luz
                MapData(X, y).LV(2) = 0
                
                MapData(X, y).LV(3) = 0
                MapData(X, y).LV(0) = 0
            Else
                               MapData(X, y).Luz = eB_Light.iUpperRight
    
                MapData(X, y).light_value(1) = DAMELONGLUZ(0)
                MapData(X, y).light_value(2) = DAMELONGLUZ(0)
                
                MapData(X, y).light_value(0) = DAMELONGLUZ(0)
                MapData(X, y).light_value(3) = DAMELONGLUZ(Luz)
                
                
                MapData(X, y).LV(1) = 0
                MapData(X, y).LV(2) = 0
                
                MapData(X, y).LV(3) = Luz
                MapData(X, y).LV(0) = 0
            End If
            
            
        ElseIf frmMain.cBL Then
            If frmMain.cINV Then
                   MapData(X, y).Luz = eB_Light.BottomLeft
    
                MapData(X, y).light_value(1) = DAMELONGLUZ(0)
                MapData(X, y).light_value(2) = DAMELONGLUZ(Luz)
                
                MapData(X, y).light_value(0) = DAMELONGLUZ(0)
                MapData(X, y).light_value(3) = DAMELONGLUZ(0)
                
                
                MapData(X, y).LV(1) = 0
                MapData(X, y).LV(2) = Luz
                
                MapData(X, y).LV(3) = 0
                MapData(X, y).LV(0) = 0
            Else
                MapData(X, y).Luz = eB_Light.iBottomLeft
    
                MapData(X, y).light_value(1) = DAMELONGLUZ(0)
                MapData(X, y).light_value(2) = DAMELONGLUZ(0)
                
                MapData(X, y).light_value(0) = DAMELONGLUZ(Luz)
                MapData(X, y).light_value(3) = DAMELONGLUZ(0)
                
                
                MapData(X, y).LV(1) = 0
                MapData(X, y).LV(2) = 0
                
                MapData(X, y).LV(3) = 0
                MapData(X, y).LV(0) = Luz
            
            End If
        ElseIf frmMain.cBR Then
        
                If frmMain.cINV Then
                    MapData(X, y).Luz = eB_Light.iBottomRight
        
                    MapData(X, y).light_value(1) = DAMELONGLUZ(0)
                    MapData(X, y).light_value(2) = DAMELONGLUZ(Luz)
                    
                    MapData(X, y).light_value(0) = DAMELONGLUZ(0)
                    MapData(X, y).light_value(3) = DAMELONGLUZ(0)
                    
                    
                    MapData(X, y).LV(1) = 0
                    MapData(X, y).LV(2) = Luz
                    
                    MapData(X, y).LV(3) = 0
                    MapData(X, y).LV(0) = 0
                Else
                    MapData(X, y).Luz = eB_Light.BottomRight
        
                    MapData(X, y).light_value(1) = DAMELONGLUZ(0)
                    MapData(X, y).light_value(2) = DAMELONGLUZ(0)
                    
                    MapData(X, y).light_value(0) = DAMELONGLUZ(0)
                    MapData(X, y).light_value(3) = DAMELONGLUZ(Luz)
                    
                    
                    MapData(X, y).LV(1) = 0
                    MapData(X, y).LV(2) = 0
                    
                    MapData(X, y).LV(3) = Luz
                    MapData(X, y).LV(0) = 0
                End If
        ElseIf frmMain.cHorizontal Then
    
    If MapData(X, y).LV(1) = 0 Then
    MapData(X, y).Luz = eB_Light.Horizontal
    
                MapData(X, y).light_value(0) = DAMELONGLUZ(Luz)
                MapData(X, y).light_value(1) = DAMELONGLUZ(Luz)
                
                MapData(X, y).light_value(2) = DAMELONGLUZ(0)
                MapData(X, y).light_value(3) = DAMELONGLUZ(0)
                
                
                MapData(X, y).LV(1) = Luz
                MapData(X, y).LV(0) = Luz
                
                MapData(X, y).LV(3) = 0
                MapData(X, y).LV(2) = 0
    Else
    
        MapData(X, y).Luz = eB_Light.Horizontal
                MapData(X, y).light_value(0) = DAMELONGLUZ(0)
                MapData(X, y).light_value(1) = DAMELONGLUZ(0)
                
                MapData(X, y).light_value(2) = DAMELONGLUZ(Luz)
                MapData(X, y).light_value(3) = DAMELONGLUZ(Luz)
                
                
                MapData(X, y).LV(1) = 0
                MapData(X, y).LV(0) = 0
                
                MapData(X, y).LV(3) = Luz
                MapData(X, y).LV(2) = Luz
    
    
    End If
    ElseIf frmMain.cVertical Then
    
    If MapData(X, y).LV(1) = 0 Then
    MapData(X, y).Luz = eB_Light.Vertical
    
                MapData(X, y).light_value(0) = DAMELONGLUZ(Luz)
                MapData(X, y).light_value(2) = DAMELONGLUZ(Luz)
                
                MapData(X, y).light_value(1) = DAMELONGLUZ(0)
                MapData(X, y).light_value(3) = DAMELONGLUZ(0)
                
                
                MapData(X, y).LV(2) = Luz
                MapData(X, y).LV(0) = Luz
                
                MapData(X, y).LV(3) = 0
                MapData(X, y).LV(1) = 0
    Else
    
        MapData(X, y).Luz = eB_Light.Vertical
                MapData(X, y).light_value(0) = DAMELONGLUZ(0)
                MapData(X, y).light_value(2) = DAMELONGLUZ(0)
                
                MapData(X, y).light_value(1) = DAMELONGLUZ(Luz)
                MapData(X, y).light_value(3) = DAMELONGLUZ(Luz)
                
                
                MapData(X, y).LV(2) = 0
                MapData(X, y).LV(0) = 0
                
                MapData(X, y).LV(3) = Luz
                MapData(X, y).LV(1) = Luz
    
    
    End If
    ElseIf frmMain.cCROSSUR Then
    MapData(X, y).Luz = eB_Light.DIAGONALUR
    
                MapData(X, y).light_value(0) = DAMELONGLUZ(0)
                MapData(X, y).light_value(3) = DAMELONGLUZ(0)
                
                MapData(X, y).light_value(1) = DAMELONGLUZ(Luz)
                MapData(X, y).light_value(2) = DAMELONGLUZ(Luz)
                
            
                MapData(X, y).LV(1) = Luz
                MapData(X, y).LV(2) = Luz
                
                MapData(X, y).LV(3) = 0
                MapData(X, y).LV(0) = 0
    
    
    End If


Else

For lx = nX To Xx
    For ly = nY To xY
        If Not (Borde = 1 And (lx = nX Or lx = Xx Or ly = nY Or ly = xY)) Then
            MapData(lx, ly).Luz = Luz
            MapData(lx, ly).LV(0) = Luz
            MapData(lx, ly).LV(1) = Luz
            MapData(lx, ly).LV(2) = Luz
            MapData(lx, ly).LV(3) = Luz
            
            If MapData(lx, ly).Luz = 0 Then
                MapData(lx, ly).light_value(0) = 0
                MapData(lx, ly).light_value(1) = 0
                MapData(lx, ly).light_value(2) = 0
                MapData(lx, ly).light_value(3) = 0
            Else
                MapData(lx, ly).light_value(0) = ambient_light((HoraLuz * 9) + Luz)
                MapData(lx, ly).light_value(1) = ambient_light((HoraLuz * 9) + Luz)
                MapData(lx, ly).light_value(2) = ambient_light((HoraLuz * 9) + Luz)
                MapData(lx, ly).light_value(3) = ambient_light((HoraLuz * 9) + Luz)
                
            End If
        End If
    Next ly
Next lx
If Borde = 1 Then

For lx = nX To Xx
For ly = nY To xY
            If ly = xY Or ly = nY Or lx = nX Or lx = Xx Then
                If ly = xY And lx = Xx Then
                    AplicarBordeManual lx, ly, 2
                ElseIf ly = nY And lx = Xx Then
                    AplicarBordeManual lx, ly, 4
                ElseIf ly = xY And lx = nX Then
                    AplicarBordeManual lx, ly, 3
                ElseIf ly = nY And lx = nX Then
                    
                    AplicarBordeManual lx, ly, 5
                ElseIf lx = nX Or lx = Xx Then
                    AplicarBordeManual lx, ly, 1
                ElseIf ly = nY Or ly = xY Then
                    AplicarBordeManual lx, ly, 0
                End If

            End If


Next ly
Next lx




End If
End If


End Sub
Public Function DAMELONGLUZ(ByVal Luz As Byte) As Long
    
    If Luz = 0 Then
        DAMELONGLUZ = base_light
    ElseIf Luz <= 8 Then
        DAMELONGLUZ = ambient_light((HoraLuz * 9) + Luz + 1)
    ElseIf Luz < 200 Then
        DAMELONGLUZ = Lucez(Luz - 8)
    End If
    

End Function
Public Sub RenderNewMap(TileX As Integer, TileY As Integer, PixelOffsetX As Integer, PixelOffsetY As Integer)
'*************************************************
'Author: Unkwown
'Last modified: 31/05/06 by GS
'Last modified: 21/11/07 By Loopzer
'Last modifier: 24/11/08 by GS
'*************************************************

On Error GoTo errs
                            Dim Polygon_Ignore_Right As Byte
                            Dim Polygon_Ignore_Left  As Byte
                            Dim Polygon_Ignore_Top  As Byte
                            Dim Polygon_Ignore_lower As Byte
                            Dim Corner As Byte
Dim y       As Integer              'Keeps track of where on map we are
Dim X       As Integer
Dim MinY    As Integer              'Start Y pos on current map
Dim MaxY    As Integer              'End Y pos on current map
Dim MinX    As Integer              'Start X pos on current map
Dim MaxX    As Integer              'End X pos on current map
Dim ScreenX As Integer              'Keeps track of where to place tile on screen
Dim ScreenY As Integer
Dim Sobre   As Integer
Dim iPPx    As Integer              'Usado en el Layer de Chars
Dim iPPy    As Integer              'Usado en el Layer de Chars
Dim Grh     As Grh                  'Temp Grh for show tile and blocked
Dim bCapa    As Byte                 'cCapas ' 31/05/2006 - GS, control de Capas
Dim iGrhIndex           As Integer  'Usado en el Layer 1
Dim PixelOffsetXTemp    As Integer  'For centering grhs
Dim PixelOffsetYTemp    As Integer
Dim TempChar            As Char
Dim tiempo As Byte
Dim colorlist(3) As Long
Dim nGrh As tnGrh

    Dim VertexArray(0 To 3) As TLVERTEX
    Dim Tex As Direct3DTexture8
    Dim SrcWidth As Integer
    Dim Width As Integer
    Dim SrcHeight As Integer
    Dim Height As Integer
    Dim SrcBitmapWidth As Long
    Dim SrcBitmapHeight As Long
    Dim xb As Integer
    Dim yb As Integer
    'Dim iGrhIndex As Integer
    Dim srdesc As D3DSURFACE_DESC

tiempo = 255
colorlist(0) = D3DColorXRGB(255, 200, 0)
colorlist(1) = D3DColorXRGB(255, 200, 0)
colorlist(2) = D3DColorXRGB(255, 200, 0)
colorlist(3) = D3DColorXRGB(255, 200, 0)

Map_LightsRender
If Not guardobmp Then
MinY = (TileY - (WindowTileHeight \ 2)) - TileBufferSize
MaxY = (TileY + (WindowTileHeight \ 2)) + TileBufferSize
MinX = (TileX - (WindowTileWidth \ 2)) - TileBufferSize
MaxX = (TileX + (WindowTileWidth \ 2)) + TileBufferSize

Else
MinY = TileY - 8
MaxY = TileY + 16
MinX = TileX - 8
MaxX = TileX + 16

End If
' 31/05/2006 - GS, control de Capas
If Val(cCapaSel) >= 1 And (cCapaSel) <= 5 Then
    bCapa = cCapaSel
ElseIf cCapaSel = 9 Then
    bCapa = 2
Else
    bCapa = 1
End If
GenerarVista 'Loopzer
ScreenY = -8
tiempo = 254


For y = (MinY) To (MaxY)
    ScreenX = -8
    For X = (MinX) To (MaxX)

        If InMapBounds(X, y) Then
                                xb = (ScreenX - 1) * 32 + PixelOffsetX
                        yb = (ScreenY - 1) * 32 + PixelOffsetY
            'If X > 100 Or Y < 1 Then Exit For ' 30/05/2006

            'Layer 1 **********************************
            If VerCapa1 Then
                With MapData(X, y)
                    If MapData(X, y).Graphic(1).index > 0 Then


   

tiempo = 1
    
                        Set Tex = DXPool.GetTexture(MapData(X, y).Graphic(1).index)

                        Tex.GetLevelDesc 0, srdesc
    
  
    VertexArray(0).rhw = 1
    VertexArray(1).rhw = 1
    VertexArray(2).rhw = 1
    VertexArray(3).rhw = 1
        

        If MapData(X, y).Luz <= 201 Or MapData(X, y).Luz >= 218 Then
        
        
        'Find the left side of the rectangle
        VertexArray(0).X = xb
        VertexArray(0).tu = (Indice_X(.IndexB(1)) / srdesc.Width)
 
        'Find the top side of the rectangle
        VertexArray(0).y = yb
        VertexArray(0).tv = (Indice_Y(.IndexB(1)) / srdesc.Height)
   
        'Find the right side of the rectangle
        VertexArray(1).X = xb + TilePixelWidth
        VertexArray(1).tu = (Indice_X(.IndexB(1)) + TilePixelWidth) / srdesc.Width
 
        'These values will only equal each other when not a shadow
        VertexArray(2).X = VertexArray(0).X
        VertexArray(3).X = VertexArray(1).X
 

       VertexArray(2).y = yb + TilePixelWidth
       VertexArray(2).tv = (Indice_Y(.IndexB(1)) + TilePixelWidth) / srdesc.Height
    
       VertexArray(1).y = VertexArray(0).y
       VertexArray(1).tv = VertexArray(0).tv
       VertexArray(2).tu = VertexArray(0).tu
       VertexArray(3).y = VertexArray(2).y
       VertexArray(3).tu = VertexArray(1).tu
       VertexArray(3).tv = VertexArray(2).tv
   
    If MapData(X, y).light_value(0) <> 0 Then
        VertexArray(0).Color = MapData(X, y).light_value(0)
    Else
        VertexArray(0).Color = base_light
    End If
      If MapData(X, y).light_value(1) <> 0 Then
        VertexArray(1).Color = MapData(X, y).light_value(1)
    Else
        VertexArray(1).Color = base_light
    End If
    If MapData(X, y).light_value(2) <> 0 Then
        VertexArray(2).Color = MapData(X, y).light_value(2)
    Else
        VertexArray(2).Color = base_light
    End If
    If MapData(X, y).light_value(3) <> 0 Then
        VertexArray(3).Color = MapData(X, y).light_value(3)
    Else
        VertexArray(3).Color = base_light
    End If
   
   
   Else
   
        'Find the left side of the rectangle
        VertexArray(1).X = xb
        VertexArray(1).tu = (Indice_X(.IndexB(1)) / srdesc.Width)
 
        'Find the top side of the rectangle
        VertexArray(1).y = yb
        VertexArray(1).tv = (Indice_Y(.IndexB(1)) / srdesc.Height)
   
        'Find the right side of the rectangle
        VertexArray(3).X = xb + TilePixelWidth
        VertexArray(3).tu = (Indice_X(.IndexB(1)) + TilePixelWidth) / srdesc.Width
 
        'These values will only equal each other when not a shadow
        VertexArray(0).X = VertexArray(1).X
        VertexArray(2).X = VertexArray(3).X
 
    'Find the bottom of the rectangle
    VertexArray(0).y = yb + TilePixelWidth
    VertexArray(0).tv = (Indice_Y(.IndexB(1)) + TilePixelWidth) / srdesc.Height
 
    'Because this is a perfect rectangle, all of the values below will equal one of the values we already got
    VertexArray(3).y = VertexArray(1).y
    VertexArray(3).tv = VertexArray(1).tv
    VertexArray(0).tu = VertexArray(1).tu
    VertexArray(2).y = VertexArray(0).y
    VertexArray(2).tu = VertexArray(3).tu
    VertexArray(2).tv = VertexArray(0).tv
   
    
    If MapData(X, y).light_value(0) <> 0 Then
        VertexArray(0).Color = MapData(X, y).light_value(0)
    Else
        VertexArray(0).Color = base_light
    End If
      If MapData(X, y).light_value(1) <> 0 Then
        VertexArray(1).Color = MapData(X, y).light_value(1)
    Else
        VertexArray(1).Color = base_light
    End If
    If MapData(X, y).light_value(2) <> 0 Then
        VertexArray(2).Color = MapData(X, y).light_value(2)
    Else
        VertexArray(2).Color = base_light
    End If
    If MapData(X, y).light_value(3) <> 0 Then
        VertexArray(3).Color = MapData(X, y).light_value(3)
    Else
        VertexArray(3).Color = base_light
    End If
   
   End If


    

    

    ddevice.SetTexture 0, Tex
    
    
   


    ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexArray(0), 28
    

    If frmMain.cVerIndices.value And frmMain.LayerC.ListIndex = 0 Then DrawText xb, yb, CStr(.IndexB(1)), D3DWHITE

    
End If
    End With
End If
End If
        ScreenX = ScreenX + 1
    Next X
    ScreenY = ScreenY + 1
    If y > 100 Then Exit For
Next y
ScreenY = -8
            
For y = (MinY) To (MaxY)
    ScreenX = -8
    For X = (MinX) To (MaxX)

        If InMapBounds(X, y) Then
            'Layer 2 **********************************
          tiempo = 2
If MapData(X, y).Graphic(2).index > 0 And VerCapa2 Then
                                xb = (ScreenX - 1) * 32 + PixelOffsetX
                        yb = (ScreenY - 1) * 32 + PixelOffsetY
                                 Set Tex = DXPool.GetTexture(MapData(X, y).Graphic(2).index)
                        Tex.GetLevelDesc 0, srdesc
    With MapData(X, y)
  
    VertexArray(0).rhw = 1
    VertexArray(1).rhw = 1
    VertexArray(2).rhw = 1
    VertexArray(3).rhw = 1
        

        If MapData(X, y).Luz <= 201 Or MapData(X, y).Luz >= 218 Then
        
        
        'Find the left side of the rectangle
        VertexArray(0).X = xb
        VertexArray(0).tu = (Indice_X(.IndexB(2)) / srdesc.Width)
 
        'Find the top side of the rectangle
        VertexArray(0).y = yb
        VertexArray(0).tv = (Indice_Y(.IndexB(2)) / srdesc.Height)
   
        'Find the right side of the rectangle
        VertexArray(1).X = xb + TilePixelWidth
        VertexArray(1).tu = (Indice_X(.IndexB(2)) + TilePixelWidth) / srdesc.Width
 
        'These values will only equal each other when not a shadow
        VertexArray(2).X = VertexArray(0).X
        VertexArray(3).X = VertexArray(1).X
 

       VertexArray(2).y = yb + TilePixelWidth
       VertexArray(2).tv = (Indice_Y(.IndexB(2)) + TilePixelWidth) / srdesc.Height
    
       VertexArray(1).y = VertexArray(0).y
       VertexArray(1).tv = VertexArray(0).tv
       VertexArray(2).tu = VertexArray(0).tu
       VertexArray(3).y = VertexArray(2).y
       VertexArray(3).tu = VertexArray(1).tu
       VertexArray(3).tv = VertexArray(2).tv
   
    If MapData(X, y).light_value(0) <> 0 Then
        VertexArray(0).Color = MapData(X, y).light_value(0)
    Else
        VertexArray(0).Color = base_light
    End If
      If MapData(X, y).light_value(1) <> 0 Then
        VertexArray(1).Color = MapData(X, y).light_value(1)
    Else
        VertexArray(1).Color = base_light
    End If
    If MapData(X, y).light_value(2) <> 0 Then
        VertexArray(2).Color = MapData(X, y).light_value(2)
    Else
        VertexArray(2).Color = base_light
    End If
    If MapData(X, y).light_value(3) <> 0 Then
        VertexArray(3).Color = MapData(X, y).light_value(3)
    Else
        VertexArray(3).Color = base_light
    End If
                               If ((MapData(X, y).TipoTerreno And eTipoTerreno.Agua) Or (MapData(X, y).TipoTerreno And eTipoTerreno.Lava)) Then

       
                            Polygon_Ignore_Right = 0
                            Polygon_Ignore_Left = 0
                            Polygon_Ignore_Top = 0
                            Polygon_Ignore_lower = 0
                            Corner = 0
                            
                            If y <> 1 Then
                              If Not MapData(X, y - 1).TipoTerreno And eTipoTerreno.Agua And Not (MapData(X, y - 1).TipoTerreno And eTipoTerreno.Lava) Then Polygon_Ignore_Top = 1
                            End If
                            
                            If y <> 100 Then
                              If Not MapData(X, y + 1).TipoTerreno And eTipoTerreno.Agua And Not (MapData(X, y + 1).TipoTerreno And eTipoTerreno.Lava) Then Polygon_Ignore_lower = 1
                            End If
                            
                            If X <> 100 Then
                              If Not MapData(X + 1, y).TipoTerreno And eTipoTerreno.Agua And Not (MapData(X + 1, y).TipoTerreno And eTipoTerreno.Lava) Then Polygon_Ignore_Right = 1
                            End If
                            
                            If X <> 1 Then
                              If Not MapData(X - 1, y).TipoTerreno And eTipoTerreno.Agua And Not (MapData(X - 1, y).TipoTerreno And eTipoTerreno.Lava) Then Polygon_Ignore_Left = 1
                            End If
                            
                          If Polygon_Ignore_Left = 0 Then
                                If X > 1 And y > 1 Then
                                If MapData(X, y - 1).TipoTerreno And eTipoTerreno.Agua And (Not MapData(X - 1, y - 1).TipoTerreno And eTipoTerreno.Agua) Then
                                    Corner = 2
                                End If
                                End If
                                If X > 1 And y < 100 Then
                                If MapData(X, y + 1).TipoTerreno And eTipoTerreno.Agua And Not (MapData(X - 1, y + 1).TipoTerreno And eTipoTerreno.Agua) Then
                                    Corner = 1
                                End If
                                End If
                            End If
                            If Polygon_Ignore_Right = 0 Then
                                If X < 100 And y > 1 Then
                                If MapData(X, y - 1).TipoTerreno And eTipoTerreno.Agua And Not (MapData(X + 1, y - 1).TipoTerreno And eTipoTerreno.Agua) Then
                                    Corner = 4
                                End If
                                End If
                                If X < 100 And y < 100 Then
                                If MapData(X, y + 1).TipoTerreno And eTipoTerreno.Agua And Not (MapData(X + 1, y + 1).TipoTerreno And eTipoTerreno.Agua) Then
                                    Corner = 3
                                End If
                                End If
                            End If
                            


              
                            
                            If Polygon_Ignore_Right <> 1 Then
                            VertexArray(1).X = VertexArray(1).X + PolygonX
                            VertexArray(3).X = VertexArray(3).X + PolygonX
                            End If
                            If Polygon_Ignore_Left <> 1 Then
       VertexArray(0).X = VertexArray(0).X + PolygonX
                            VertexArray(2).X = VertexArray(2).X + PolygonX
                            End If

                            If Polygon_Ignore_Top <> 1 Then
                                VertexArray(0).y = VertexArray(0).y + polygonCount(1)
                                VertexArray(1).y = VertexArray(1).y - polygonCount(1)
                            End If

                            If Polygon_Ignore_lower <> 1 Then
                                VertexArray(2).y = VertexArray(2).y + polygonCount(1)
                                VertexArray(3).y = VertexArray(3).y - polygonCount(1)
                            End If
                            
                End If
                         

   
   Else
   
        'Find the left side of the rectangle
        VertexArray(1).X = xb
        VertexArray(1).tu = (Indice_X(.IndexB(2)) / srdesc.Width)
 
        'Find the top side of the rectangle
        VertexArray(1).y = yb
        VertexArray(1).tv = (Indice_Y(.IndexB(2)) / srdesc.Height)
   
        'Find the right side of the rectangle
        VertexArray(3).X = xb + TilePixelWidth
        VertexArray(3).tu = (Indice_X(.IndexB(2)) + TilePixelWidth) / srdesc.Width
 
        'These values will only equal each other when not a shadow
        VertexArray(0).X = VertexArray(1).X
        VertexArray(2).X = VertexArray(3).X
 
    'Find the bottom of the rectangle
    VertexArray(0).y = yb + TilePixelWidth
    VertexArray(0).tv = (Indice_Y(.IndexB(2)) + TilePixelWidth) / srdesc.Height
 
    'Because this is a perfect rectangle, all of the values below will equal one of the values we already got
    VertexArray(3).y = VertexArray(1).y
    VertexArray(3).tv = VertexArray(1).tv
    VertexArray(0).tu = VertexArray(1).tu
    VertexArray(2).y = VertexArray(0).y
    VertexArray(2).tu = VertexArray(3).tu
    VertexArray(2).tv = VertexArray(0).tv
   
    
    If MapData(X, y).light_value(0) <> 0 Then
        VertexArray(0).Color = MapData(X, y).light_value(0)
    Else
        VertexArray(0).Color = base_light
    End If
      If MapData(X, y).light_value(1) <> 0 Then
        VertexArray(1).Color = MapData(X, y).light_value(1)
    Else
        VertexArray(1).Color = base_light
    End If
    If MapData(X, y).light_value(2) <> 0 Then
        VertexArray(2).Color = MapData(X, y).light_value(2)
    Else
        VertexArray(2).Color = base_light
    End If
    If MapData(X, y).light_value(3) <> 0 Then
        VertexArray(3).Color = MapData(X, y).light_value(3)
    Else
        VertexArray(3).Color = base_light
    End If
    If ((MapData(X, y).TipoTerreno And eTipoTerreno.Agua) Or (MapData(X, y).TipoTerreno And eTipoTerreno.Lava)) Then

       
                            Polygon_Ignore_Right = 0
                            Polygon_Ignore_Left = 0
                            Polygon_Ignore_Top = 0
                            Polygon_Ignore_lower = 0
                            Corner = 0
                            
                            If y <> 1 Then
                              If Not MapData(X, y - 1).TipoTerreno And eTipoTerreno.Agua And Not (MapData(X, y - 1).TipoTerreno And eTipoTerreno.Lava) Then Polygon_Ignore_Top = 1
                            End If
                            
                            If y <> 100 Then
                              If Not MapData(X, y + 1).TipoTerreno And eTipoTerreno.Agua And Not (MapData(X, y + 1).TipoTerreno And eTipoTerreno.Lava) Then Polygon_Ignore_lower = 1
                            End If
                            
                            If X <> 100 Then
                              If Not MapData(X + 1, y).TipoTerreno And eTipoTerreno.Agua And Not (MapData(X + 1, y).TipoTerreno And eTipoTerreno.Lava) Then Polygon_Ignore_Right = 1
                            End If
2
                            If X <> 1 Then
                              If Not MapData(X - 1, y).TipoTerreno And eTipoTerreno.Agua And Not (MapData(X - 1, y).TipoTerreno And eTipoTerreno.Lava) Then Polygon_Ignore_Left = 1
                            End If
                            
                          If Polygon_Ignore_Left = 0 Then
                                If X > 1 And y > 1 Then
                                If MapData(X, y - 1).TipoTerreno And eTipoTerreno.Agua And (Not MapData(X - 1, y - 1).TipoTerreno And eTipoTerreno.Agua) Then
                                    Corner = 2
                                End If
                                End If
                                If X > 1 And y < 100 Then
                                If MapData(X, y + 1).TipoTerreno And eTipoTerreno.Agua And Not (MapData(X - 1, y + 1).TipoTerreno And eTipoTerreno.Agua) Then
                                    Corner = 1
                                End If
                                End If
                            End If
                            If Polygon_Ignore_Right = 0 Then
                                If X < 100 And y > 1 Then
                                If MapData(X, y - 1).TipoTerreno And eTipoTerreno.Agua And Not (MapData(X + 1, y - 1).TipoTerreno And eTipoTerreno.Agua) Then
                                    Corner = 4
                                End If
                                End If
                                If X < 100 And y < 100 Then
                                If MapData(X, y + 1).TipoTerreno And eTipoTerreno.Agua And Not (MapData(X + 1, y + 1).TipoTerreno And eTipoTerreno.Agua) Then
                                    Corner = 3
                                End If
                                End If
                            End If
                            


              
                            
                            
                            VertexArray(3).X = VertexArray(3).X + PolygonX
                            VertexArray(2).X = VertexArray(2).X + PolygonX


                            If Polygon_Ignore_Top <> 1 Then
                                VertexArray(1).y = VertexArray(1).y + polygonCount(1)
                                VertexArray(3).y = VertexArray(3).y - polygonCount(1)
                            End If

                            If Polygon_Ignore_lower <> 1 Then
                                VertexArray(0).y = VertexArray(0).y + polygonCount(1)
                                VertexArray(2).y = VertexArray(2).y - polygonCount(1)
                            End If
                            
            End If
                         

   End If


    


     ddevice.SetTexture 0, Tex
    
    
   


    ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexArray(0), 28
        If frmMain.cVerIndices.value And frmMain.LayerC.ListIndex = 1 Then DrawText xb, yb, CStr(.IndexB(2)), D3DWHITE


    End With
    End If
    
    If MapData(X, y).Graphic(5).index <> 0 And VerCapa5 Then
                                    xb = (ScreenX - 1) * 32 + PixelOffsetX
                        yb = (ScreenY - 1) * 32 + PixelOffsetY
        Set Tex = DXPool.GetTexture(MapData(X, y).Graphic(5).index)
        Tex.GetLevelDesc 0, srdesc
    With MapData(X, y)
  
    VertexArray(0).rhw = 1
    VertexArray(1).rhw = 1
    VertexArray(2).rhw = 1
    VertexArray(3).rhw = 1
        

        If MapData(X, y).Luz <= 201 Or MapData(X, y).Luz >= 218 Then
        
        
        'Find the left side of the rectangle
        VertexArray(0).X = xb
        VertexArray(0).tu = (Indice_X(.IndexB(5)) / srdesc.Width)
 
        'Find the top side of the rectangle
        VertexArray(0).y = yb
        VertexArray(0).tv = (Indice_Y(.IndexB(5)) / srdesc.Height)
   
        'Find the right side of the rectangle
        VertexArray(1).X = xb + TilePixelWidth
        VertexArray(1).tu = (Indice_X(.IndexB(5)) + TilePixelWidth) / srdesc.Width
 
        'These values will only equal each other when not a shadow
        VertexArray(2).X = VertexArray(0).X
        VertexArray(3).X = VertexArray(1).X
 

       VertexArray(2).y = yb + TilePixelWidth
       VertexArray(2).tv = (Indice_Y(.IndexB(5)) + TilePixelWidth) / srdesc.Height
    
       VertexArray(1).y = VertexArray(0).y
       VertexArray(1).tv = VertexArray(0).tv
       VertexArray(2).tu = VertexArray(0).tu
       VertexArray(3).y = VertexArray(2).y
       VertexArray(3).tu = VertexArray(1).tu
       VertexArray(3).tv = VertexArray(2).tv
   
    If MapData(X, y).light_value(0) <> 0 Then
        VertexArray(0).Color = MapData(X, y).light_value(0)
    Else
        VertexArray(0).Color = base_light
    End If
      If MapData(X, y).light_value(1) <> 0 Then
        VertexArray(1).Color = MapData(X, y).light_value(1)
    Else
        VertexArray(1).Color = base_light
    End If
    If MapData(X, y).light_value(2) <> 0 Then
        VertexArray(2).Color = MapData(X, y).light_value(2)
    Else
        VertexArray(2).Color = base_light
    End If
    If MapData(X, y).light_value(3) <> 0 Then
        VertexArray(3).Color = MapData(X, y).light_value(3)
    Else
        VertexArray(3).Color = base_light
    End If
   
   
   Else
   
        'Find the left side of the rectangle
        VertexArray(1).X = xb
        VertexArray(1).tu = (Indice_X(.IndexB(5)) / srdesc.Width)
 
        'Find the top side of the rectangle
        VertexArray(1).y = yb
        VertexArray(1).tv = (Indice_Y(.IndexB(5)) / srdesc.Height)
   
        'Find the right side of the rectangle
        VertexArray(3).X = xb + TilePixelWidth
        VertexArray(3).tu = (Indice_X(.IndexB(5)) + TilePixelWidth) / srdesc.Width
 
        'These values will only equal each other when not a shadow
        VertexArray(0).X = VertexArray(1).X
        VertexArray(2).X = VertexArray(3).X
 
    'Find the bottom of the rectangle
    VertexArray(0).y = yb + TilePixelWidth
    VertexArray(0).tv = (Indice_Y(.IndexB(5)) + TilePixelWidth) / srdesc.Height
 
    'Because this is a perfect rectangle, all of the values below will equal one of the values we already got
    VertexArray(3).y = VertexArray(1).y
    VertexArray(3).tv = VertexArray(1).tv
    VertexArray(0).tu = VertexArray(1).tu
    VertexArray(2).y = VertexArray(0).y
    VertexArray(2).tu = VertexArray(3).tu
    VertexArray(2).tv = VertexArray(0).tv
   
    
    If MapData(X, y).light_value(0) <> 0 Then
        VertexArray(0).Color = MapData(X, y).light_value(0)
    Else
        VertexArray(0).Color = base_light
    End If
      If MapData(X, y).light_value(1) <> 0 Then
        VertexArray(1).Color = MapData(X, y).light_value(1)
    Else
        VertexArray(1).Color = base_light
    End If
    If MapData(X, y).light_value(2) <> 0 Then
        VertexArray(2).Color = MapData(X, y).light_value(2)
    Else
        VertexArray(2).Color = base_light
    End If
    If MapData(X, y).light_value(3) <> 0 Then
        VertexArray(3).Color = MapData(X, y).light_value(3)
    Else
        VertexArray(3).Color = base_light
    End If
   
   End If


    
    VertexArray(0).y = VertexArray(0).y - MapData(X, y).AlturaPoligonos(0)
    VertexArray(1).y = VertexArray(1).y - MapData(X, y).AlturaPoligonos(1)
    VertexArray(2).y = VertexArray(2).y - MapData(X, y).AlturaPoligonos(2)
    VertexArray(3).y = VertexArray(3).y - MapData(X, y).AlturaPoligonos(3)
    

     ddevice.SetTexture 0, Tex
    
    
   


    ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexArray(0), 28
        If frmMain.cVerIndices.value And frmMain.LayerC.ListIndex = 4 Then DrawText xb, yb, CStr(.IndexB(5)), D3DWHITE


    End With
    End If
    
End If
            

        
        ScreenX = ScreenX + 1
    Next X
    ScreenY = ScreenY + 1
    If y > 100 Then Exit For
Next y
ScreenY = -8


tiempo = 3
For y = (MinY) To (MaxY)   '- 8+ 8
    ScreenX = -8
    For X = (MinX) To (MaxX)   '- 8 + 8
        If InMapBounds(X, y) Then
            If X > 100 Or X < -3 Then Exit For ' 30/05/2006

            iPPx = ((32 * ScreenX) - 32) + PixelOffsetX
            iPPy = ((32 * ScreenY) - 32) + PixelOffsetY
             'Object Layer **********************************
             If MapData(X, y).OBJInfo.objindex <> 0 And VerObjetos Then
                If MapData(X, y).Luz >= 202 And MapData(X, y).Luz <= 217 Then
                    modGrh.Grh_iRenderN MapData(X, y).ObjGrh, iPPx, iPPy, MapData(X, y).light_value, True
                Else
                    modGrh.Grh_RenderN MapData(X, y).ObjGrh, iPPx, iPPy, MapData(X, y).light_value, True
                End If
             End If
             If MapData(X, y).DecorI > 0 And MapData(X, y).DecorGrh.index > 0 And VerDecors Then
                If TipoSeleccionado = 1 Then
                    If ObjetoSeleccionado.X = X And ObjetoSeleccionado.y = y Then
                        If MapData(X, y).Luz >= 202 And MapData(X, y).Luz <= 217 Then

                        modGrh.Grh_iRenderN SeleccionnGrh, iPPx, iPPy + (EstaticData(NewIndexData(SeleccionIndex).Estatic).H * 0.5), SeleccionadoArrayColor, True
                        Else
                        modGrh.Grh_RenderN SeleccionnGrh, iPPx, iPPy + (EstaticData(NewIndexData(SeleccionIndex).Estatic).H * 0.5), SeleccionadoArrayColor, True
                        
                        End If
                    End If
                End If
                If MapData(X, y).Luz >= 202 And MapData(X, y).Luz <= 217 Then
                    modGrh.Grh_iRenderN MapData(X, y).DecorGrh, iPPx, iPPy, MapData(X, y).light_value, True
                    
                Else
                    modGrh.Grh_RenderN MapData(X, y).DecorGrh, iPPx, iPPy, MapData(X, y).light_value, True
                End If

             End If
             tiempo = 4
                  'Char layer **********************************

                 If MapData(X, y).CHarIndex <> 0 And VerNpcs Then
                 
                     TempChar = CharList(MapData(X, y).CHarIndex)

                     PixelOffsetXTemp = PixelOffsetX
                     PixelOffsetYTemp = PixelOffsetY
                    
                    If TipoSeleccionado = 2 Then
                    If ObjetoSeleccionado.X = X And ObjetoSeleccionado.y = y Then
                        If MapData(X, y).Luz >= 202 And MapData(X, y).Luz <= 217 Then

                        modGrh.Grh_iRenderN SeleccionnGrh, iPPx, iPPy + (EstaticData(NewIndexData(SeleccionIndex).Estatic).H * 0.5), SeleccionadoArrayColor, True
                        Else
                        modGrh.Grh_RenderN SeleccionnGrh, iPPx, iPPy + (EstaticData(NewIndexData(SeleccionIndex).Estatic).H * 0.5), SeleccionadoArrayColor, True
                        
                        End If
                    End If
                End If
                    
                    
                   'Dibuja solamente players
                   If TempChar.Head(TempChar.Heading).index <> 0 Then
                If MapData(X, y).Luz >= 202 And MapData(X, y).Luz <= 217 Then
                     modGrh.Anim_iRender TempChar.Body(TempChar.Heading), iPPx, iPPy, MapData(X, y).light_value, True, False
                     'Draw Head
                     modGrh.Grh_iRenderN TempChar.Head(TempChar.Heading), iPPx, iPPy + BodyData(TempChar.iBody).OffsetY + HeadData(TempChar.iHead).OffsetDibujoY, MapData(X, y).light_value, True
                   
                Else
                      modGrh.Anim_Render TempChar.Body(TempChar.Heading), iPPx, iPPy, MapData(X, y).light_value, True, False, BodyData(TempChar.iBody).OverWriteGrafico
                     'Draw Head
                     modGrh.Grh_RenderN TempChar.Head(TempChar.Heading), iPPx, iPPy + BodyData(TempChar.iBody).OffsetY + HeadData(TempChar.iHead).OffsetDibujoY, MapData(X, y).light_value, True
                                  
                End If
                   Else
                   
                If MapData(X, y).Luz >= 202 And MapData(X, y).Luz <= 217 Then
                     modGrh.Anim_iRender TempChar.Body(TempChar.Heading), iPPx, iPPy, MapData(X, y).light_value, True, False, BodyData(TempChar.iBody).OverWriteGrafico
                Else
                      modGrh.Anim_Render TempChar.Body(TempChar.Heading), iPPx, iPPy, MapData(X, y).light_value, True, False, BodyData(TempChar.iBody).OverWriteGrafico
                End If
            End If
            
                 End If


                 tiempo = 5


           If MapData(X, y).Graphic(3).index <> 0 And VerCapa3 Then
            Set Tex = DXPool.GetTexture(MapData(X, y).Graphic(3).index)
                Tex.GetLevelDesc 0, srdesc
    With MapData(X, y)
  
    VertexArray(0).rhw = 1
    VertexArray(1).rhw = 1
    VertexArray(2).rhw = 1
    VertexArray(3).rhw = 1
        

        If MapData(X, y).Luz <= 201 Or MapData(X, y).Luz >= 218 Then
        
        
        'Find the left side of the rectangle
        VertexArray(0).X = iPPx
        VertexArray(0).tu = (Indice_X(.IndexB(3)) / srdesc.Width)
 
        'Find the top side of the rectangle
        VertexArray(0).y = iPPy
        VertexArray(0).tv = (Indice_Y(.IndexB(3)) / srdesc.Height)
   
        'Find the right side of the rectangle
        VertexArray(1).X = iPPx + TilePixelWidth
        VertexArray(1).tu = (Indice_X(.IndexB(3)) + TilePixelWidth) / srdesc.Width
 
        'These values will only equal each other when not a shadow
        VertexArray(2).X = VertexArray(0).X
        VertexArray(3).X = VertexArray(1).X
 

       VertexArray(2).y = iPPy + TilePixelWidth
       VertexArray(2).tv = (Indice_Y(.IndexB(3)) + TilePixelWidth) / srdesc.Height
    
       VertexArray(1).y = VertexArray(0).y
       VertexArray(1).tv = VertexArray(0).tv
       VertexArray(2).tu = VertexArray(0).tu
       VertexArray(3).y = VertexArray(2).y
       VertexArray(3).tu = VertexArray(1).tu
       VertexArray(3).tv = VertexArray(2).tv
   
    If MapData(X, y).light_value(0) <> 0 Then
        VertexArray(0).Color = MapData(X, y).light_value(0)
    Else
        VertexArray(0).Color = base_light
    End If
      If MapData(X, y).light_value(1) <> 0 Then
        VertexArray(1).Color = MapData(X, y).light_value(1)
    Else
        VertexArray(1).Color = base_light
    End If
    If MapData(X, y).light_value(2) <> 0 Then
        VertexArray(2).Color = MapData(X, y).light_value(2)
    Else
        VertexArray(2).Color = base_light
    End If
    If MapData(X, y).light_value(3) <> 0 Then
        VertexArray(3).Color = MapData(X, y).light_value(3)
    Else
        VertexArray(3).Color = base_light
    End If
   
   
   Else
   
        'Find the left side of the rectangle
        VertexArray(1).X = iPPx
        VertexArray(1).tu = (Indice_X(.IndexB(3)) / srdesc.Width)
 
        'Find the top side of the rectangle
        VertexArray(1).y = iPPy
        VertexArray(1).tv = (Indice_Y(.IndexB(3)) / srdesc.Height)
   
        'Find the right side of the rectangle
        VertexArray(3).X = iPPx + TilePixelWidth
        VertexArray(3).tu = (Indice_X(.IndexB(3)) + TilePixelWidth) / srdesc.Width
 
        'These values will only equal each other when not a shadow
        VertexArray(0).X = VertexArray(1).X
        VertexArray(2).X = VertexArray(3).X
 
    'Find the bottom of the rectangle
    VertexArray(0).y = iPPy + TilePixelWidth
    VertexArray(0).tv = (Indice_Y(.IndexB(3)) + TilePixelWidth) / srdesc.Height
 
    'Because this is a perfect rectangle, all of the values below will equal one of the values we already got
    VertexArray(3).y = VertexArray(1).y
    VertexArray(3).tv = VertexArray(1).tv
    VertexArray(0).tu = VertexArray(1).tu
    VertexArray(2).y = VertexArray(0).y
    VertexArray(2).tu = VertexArray(3).tu
    VertexArray(2).tv = VertexArray(0).tv
   
    
    If MapData(X, y).light_value(0) <> 0 Then
        VertexArray(0).Color = MapData(X, y).light_value(0)
    Else
        VertexArray(0).Color = base_light
    End If
      If MapData(X, y).light_value(1) <> 0 Then
        VertexArray(1).Color = MapData(X, y).light_value(1)
    Else
        VertexArray(1).Color = base_light
    End If
    If MapData(X, y).light_value(2) <> 0 Then
        VertexArray(2).Color = MapData(X, y).light_value(2)
    Else
        VertexArray(2).Color = base_light
    End If
    If MapData(X, y).light_value(3) <> 0 Then
        VertexArray(3).Color = MapData(X, y).light_value(3)
    Else
        VertexArray(3).Color = base_light
    End If
   
   End If


    
    VertexArray(0).y = VertexArray(0).y - MapData(X, y).AlturaPoligonos(0)
    VertexArray(1).y = VertexArray(1).y - MapData(X, y).AlturaPoligonos(1)
    VertexArray(2).y = VertexArray(2).y - MapData(X, y).AlturaPoligonos(2)
    VertexArray(3).y = VertexArray(3).y - MapData(X, y).AlturaPoligonos(3)
    

     ddevice.SetTexture 0, Tex
    
    
   


    ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexArray(0), 28
        If frmMain.cVerIndices.value And frmMain.LayerC.ListIndex = 2 Then DrawText iPPx, iPPy, CStr(.IndexB(3)), D3DWHITE


    End With
    End If
             
             
             
             If MapData(X, y).SPOTLIGHT.index > 0 Then
                SPOT_LIGHTS(MapData(X, y).SPOTLIGHT.index).X = ((32 * ScreenX) - 32) + PixelOffsetX
                SPOT_LIGHTS(MapData(X, y).SPOTLIGHT.index).y = ((32 * ScreenY) - 32) + PixelOffsetY
                SPOT_LIGHTS(MapData(X, y).SPOTLIGHT.index).Mustbe_Render = True
                If frmMain.MarcarsPOT.value Then
                    nGrh.index = 247

                    modGrh.Grh_RenderN nGrh, ((32 * ScreenX) - 32) + PixelOffsetX, ((32 * ScreenY) - 32) + PixelOffsetY, MapData(X, y).light_value, True
                End If
            End If
             
             tiempo = 6

            tiempo = 7
        End If
        

        ScreenX = ScreenX + 1
    Next X
    ScreenY = ScreenY + 1
Next y




'Tiles blokeadas, techos, triggers , seleccion
ScreenY = -8
For y = (MinY) To (MaxY)
    ScreenX = -8
    For X = (MinX) To (MaxX)
        If X < 101 And X > 0 And y < 101 And y > 0 Then ' 30/05/2006
            iPPx = ((32 * ScreenX) - 32) + PixelOffsetX
            iPPy = ((32 * ScreenY) - 32) + PixelOffsetY
            
            
                         If MapData(X, y).particle_group Then
                  modDXEngine.Particle_Group_Render MapData(X, y).particle_group, iPPx, iPPy

             End If
            If frmMain.cVerLuces.value And MapData(X, y).Luz > 0 Then
                'modDXEngine.DXEngine_TextRender 1, MapData(x, Y).Luz, iPPx, iPPy, D3DColorXRGB(255, 0, 0), DT_CENTER, 32, 32
                modDXEngine.DrawText iPPx, iPPy, MapData(X, y).Luz, D3DRED
            ElseIf frmMain.chkParticle.value And MapData(X, y).particle_group Then
                DrawText iPPx, iPPy, "P:" & CStr(MapData(X, y).parti_index), D3DWHITE
            ElseIf frmMain.ChkInterior.value And MapData(X, y).InteriorVal > 0 Then
                DrawText iPPx, iPPy, CStr(MapData(X, y).InteriorVal), D3DWHITE
            ElseIf frmMain.cTipoTerreno.value Then
                If MapData(X, y).TipoTerreno > 0 Then DrawText iPPx, iPPy, CStr(MapData(X, y).TipoTerreno), D3DRED
            End If
            
            
           If MapData(X, y).Graphic(4).index <> 0 And VerCapa4 Then
            Set Tex = DXPool.GetTexture(MapData(X, y).Graphic(4).index)
                Tex.GetLevelDesc 0, srdesc
    With MapData(X, y)
  
    VertexArray(0).rhw = 1
    VertexArray(1).rhw = 1
    VertexArray(2).rhw = 1
    VertexArray(3).rhw = 1
        

        If MapData(X, y).Luz <= 201 Or MapData(X, y).Luz >= 218 Then
        
        
        'Find the left side of the rectangle
        VertexArray(0).X = iPPx
        VertexArray(0).tu = (Indice_X(.IndexB(4)) / srdesc.Width)
 
        'Find the top side of the rectangle
        VertexArray(0).y = iPPy
        VertexArray(0).tv = (Indice_Y(.IndexB(4)) / srdesc.Height)
   
        'Find the right side of the rectangle
        VertexArray(1).X = iPPx + TilePixelWidth
        VertexArray(1).tu = (Indice_X(.IndexB(4)) + TilePixelWidth) / srdesc.Width
 
        'These values will only equal each other when not a shadow
        VertexArray(2).X = VertexArray(0).X
        VertexArray(3).X = VertexArray(1).X
 

       VertexArray(2).y = iPPy + TilePixelWidth
       VertexArray(2).tv = (Indice_Y(.IndexB(4)) + TilePixelWidth) / srdesc.Height
    
       VertexArray(1).y = VertexArray(0).y
       VertexArray(1).tv = VertexArray(0).tv
       VertexArray(2).tu = VertexArray(0).tu
       VertexArray(3).y = VertexArray(2).y
       VertexArray(3).tu = VertexArray(1).tu
       VertexArray(3).tv = VertexArray(2).tv
   

        VertexArray(0).Color = base_light


        VertexArray(1).Color = base_light


        VertexArray(2).Color = base_light


        VertexArray(3).Color = base_light

   
   
   Else
   
        'Find the left side of the rectangle
        VertexArray(1).X = iPPx
        VertexArray(1).tu = (Indice_X(.IndexB(4)) / srdesc.Width)
 
        'Find the top side of the rectangle
        VertexArray(1).y = iPPy
        VertexArray(1).tv = (Indice_Y(.IndexB(4)) / srdesc.Height)
   
        'Find the right side of the rectangle
        VertexArray(3).X = iPPx + TilePixelWidth
        VertexArray(3).tu = (Indice_X(.IndexB(4)) + TilePixelWidth) / srdesc.Width
 
        'These values will only equal each other when not a shadow
        VertexArray(0).X = VertexArray(1).X
        VertexArray(2).X = VertexArray(3).X
 
    'Find the bottom of the rectangle
    VertexArray(0).y = iPPy + TilePixelWidth
    VertexArray(0).tv = (Indice_Y(.IndexB(4)) + TilePixelWidth) / srdesc.Height
 
    'Because this is a perfect rectangle, all of the values below will equal one of the values we already got
    VertexArray(3).y = VertexArray(1).y
    VertexArray(3).tv = VertexArray(1).tv
    VertexArray(0).tu = VertexArray(1).tu
    VertexArray(2).y = VertexArray(0).y
    VertexArray(2).tu = VertexArray(3).tu
    VertexArray(2).tv = VertexArray(0).tv
   
    

        VertexArray(0).Color = base_light

        VertexArray(1).Color = base_light

        VertexArray(2).Color = base_light

        VertexArray(3).Color = base_light

   
   End If


    
    VertexArray(0).y = VertexArray(0).y - MapData(X, y).AlturaPoligonos(0)
    VertexArray(1).y = VertexArray(1).y - MapData(X, y).AlturaPoligonos(1)
    VertexArray(2).y = VertexArray(2).y - MapData(X, y).AlturaPoligonos(2)
    VertexArray(3).y = VertexArray(3).y - MapData(X, y).AlturaPoligonos(3)
    

     ddevice.SetTexture 0, Tex
    
    
   


    ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexArray(0), 28
        If frmMain.cVerIndices.value And frmMain.LayerC.ListIndex = 2 Then DrawText iPPx, iPPy, CStr(.IndexB(4)), D3DWHITE


    End With
    End If
            
            
            
            If MapData(X, y).TileExit.Map <> 0 And VerTranslados Then
                nGrh.index = 245
                modGrh.Grh_RenderN nGrh, iPPx, iPPy, MapData(X, y).light_value, True
            End If
            
            If MapData(X, y).light_index Then
                nGrh.index = 4
                modGrh.Grh_RenderN nGrh, iPPx, iPPy, colorlist, True
            End If
            
            'Show blocked tiles
            If VerBlockeados And MapData(X, y).Blocked = 1 Then
                nGrh.index = 247
                modGrh.Grh_RenderN nGrh, iPPx, iPPy, MapData(X, y).light_value, True
            End If
            If VerGrilla Then
                'Grilla 24/11/2008 by GS
                modDXEngine.DXEngine_DrawBox ((32 * ScreenX) - 32) + PixelOffsetX, ((32 * ScreenY) - 32) + PixelOffsetY, 1, 32, RGB(255, 255, 255)
                modDXEngine.DXEngine_DrawBox ((32 * ScreenX) - 32) + PixelOffsetX, ((32 * ScreenY) - 32) + PixelOffsetY, 32, 1, RGB(255, 255, 255)
            End If
            If VerTriggers Then
                'Call DrawText(PixelPos(ScreenX), PixelPos(ScreenY), Str(MapData(X, Y).Trigger), vbRed)
                If frmMain.lListado(8).Visible Then
                    If MapData(X, y).TipoTerreno <> 0 Then
                    modDXEngine.DrawText ((32 * ScreenX) - 32) + PixelOffsetX, ((32 * ScreenY) - 32) + PixelOffsetY, "T:" & CStr(MapData(X, y).TipoTerreno), D3DWHITE
                    End If
                Else
                    If MapData(X, y).Trigger <> 0 Then
                    modDXEngine.DrawText ((32 * ScreenX) - 32) + PixelOffsetX, ((32 * ScreenY) - 32) + PixelOffsetY, "G:" & CStr(MapData(X, y).Trigger), D3DWHITE
                    End If
                End If
            End If
            If frmMain.MpNw.Visible Then
                If frmMain.cBorrarSobrante.value Then
                    'Mostramos un cursor
                    If frmMain.ccursor.value Then dibujarCursor Val(frmMain.SizeC.List(frmMain.SizeC.ListIndex))
                ElseIf frmMain.cInsertarSurface.value Then
                    If frmMain.ccursor.value Then dibujarCursor 32
                ElseIf frmMain.cEditarIndice.value Then
                    If frmMain.ccursor.value Then dibujarCursor 32
                ElseIf frmMain.cAplicarTerreno.value Then
                    dibujarCursor 32
                End If
            End If
            
            If Seleccionando Then
                'If ScreenX >= SeleccionIX And ScreenX <= SeleccionFX And ScreenY >= SeleccionIY And ScreenY <= SeleccionFY Then
                    If X >= SeleccionIX And y >= SeleccionIY Then
                        If X <= SeleccionFX And y <= SeleccionFY Then
                            modDXEngine.DXEngine_DrawBox ((32 * ScreenX) - 32) + PixelOffsetX, ((32 * ScreenY) - 32) + PixelOffsetY, 32, 32, RGB(100, 255, 255)
                        End If
                    End If
            End If

        End If
        ScreenX = ScreenX + 1
    Next X
    ScreenY = ScreenY + 1
Next y

Exit Sub

errs:
Debug.Print Err.Description & "_" & X & "_" & y & "_" & tiempo

End Sub
Public Sub Iniciar_IndicesNewMap()


Dim P As Long

For P = 0 To 255


    Indice_Y(P + 1) = (Int((P) / 16)) * 32
    Indice_X(P + 1) = (P Mod 16) * 32
    

Next P
End Sub
Public Sub dibujarCursor(ByVal Tamaño As Integer)

Dim t(3) As TLVERTEX
Dim value As Byte
Dim Tex As D3D8Textures
value = 55
Dim X As Single
Dim y As Single
Set Tex.Texture = DXPool.GetTexture(0)

ConvertTPtoCP 0, 0, X, y, Mx, My
t(0).X = X
t(0).y = y

t(1).X = X + Tamaño
t(1).y = y

t(2).X = X
t(2).y = y + Tamaño

t(3).X = t(1).X
t(3).y = t(2).y

t(0).Color = base_light
t(1).Color = base_light
t(2).Color = -1
t(3).Color = -1

t(0).rhw = 1
t(1).rhw = 1
t(2).rhw = 1
t(3).rhw = 1

t(0).tu = 0
t(0).tv = 0

t(1).tu = 1
t(1).tv = 0

t(2).tu = 0
t(2).tv = 1

t(3).tu = 1
t(3).tv = 1



ddevice.SetTexture 0, Tex.Texture
ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, t(0), 28



End Sub



'//This is just a simple wrapper function that makes filling the structures much much easier...
Private Function CreateTLVertex(ByVal X As Single, ByVal y As Single, z As Single, rhw As Single, Color As Long, _
                                               Specular As Long, tu As Single, tv As Single) As TLVERTEX

CreateTLVertex.X = X
CreateTLVertex.y = y
CreateTLVertex.z = z
CreateTLVertex.rhw = rhw
CreateTLVertex.Color = Color
CreateTLVertex.tu = tu
CreateTLVertex.tv = tv
End Function
Public Sub DibujarGEnPic(ByVal PIC As PictureBox, ByVal GrhIndex As Integer, ByVal X As Integer, ByVal y As Integer, Optional ByVal size As Byte, Optional destX As Integer, Optional destY As Integer, Optional Texto As String, Optional TextPos As Byte, Optional TextoColor As Long = D3DWHITE, Optional TextoAlpha As Byte = 255)

Dim DestRect As RECT
Dim tX As Byte
Dim tY As Byte

   DestRect.top = 0
   DestRect.left = 0
   DestRect.Bottom = frmMain.SizeC.List(frmMain.SizeC.ListIndex) '* size
   DestRect.Right = frmMain.SizeC.List(frmMain.SizeC.ListIndex)
   
   'If destRect.bottom > Pic.Height Then destRect.bottom = Pic.Height
   'If destRect.bottom <= 0 Then destRect.bottom = Pic.Height
   'If destRect.right > Pic.Width Then destRect.right = Pic.Width
   'If destRect.right <= 0 Then destRect.right = Pic.Width
   ddevice.Clear 1, DestRect, D3DCLEAR_TARGET, &H0, ByVal 0, 0
   ddevice.BeginScene
   Draw_RAWGraph GrhIndex, 0, 0
   If LenB(Texto) > 0 Then
    Select Case TextPos
        Case 0 'UpperLeft
            tX = 5
            tY = 2
        Case 1 'UpperRight
            tY = 2
            tX = 24
        Case 2 'BottomLeft
            tY = 20
            tX = 5
        Case 3 'BottomRight
            tY = 20
            tX = 24
    End Select
    DrawText CInt(tX), CInt(tY), Texto, TextoColor
   End If
   ddevice.EndScene
   
   
   ddevice.Present DestRect, ByVal 0, PIC.hWnd, ByVal 0


End Sub


Public Sub Draw_RAWGraph(ByVal FileNum As Integer, ByVal X As Long, ByVal y As Long, Optional Shadow As Boolean, Optional W As Integer, Optional H As Integer)
    
Dim dx3dTextures As D3D8Textures
Dim verts(3) As TLVERTEX
Dim light_value(0 To 3) As Long
    
    Set dx3dTextures.Texture = DXPool.GetTexture(FileNum)
    Dim srdesc As D3DSURFACE_DESC
    dx3dTextures.Texture.GetLevelDesc 0, srdesc
    
    ddevice.SetTexture 0, dx3dTextures.Texture
    
    
    
    If H > 0 Then
        srdesc.Height = H
    End If
    If W > 0 Then
        srdesc.Width = W
    End If
    

'      If ((srdesc.width - 1)) And ((srdesc.height - 1)) Then
    
        With verts(2)
            .X = X
            .y = y + srdesc.Height
            .tu = 0
            .tv = 1
            .rhw = 1
            .Color = -1
        End With
        With verts(0)
            .X = X
            .y = y
            .tu = 0
            .tv = 0
            .rhw = 1
            .Color = -1

        End With
        
        With verts(3)
            .X = X + srdesc.Width
            .y = y + srdesc.Height
            .tu = 1
            .tv = 1
            .rhw = 1
            .Color = -1

        End With
        
        With verts(1)
            .X = X + srdesc.Width
            .y = y
            .tu = 1
            .tv = 0
            .rhw = 1
            .Color = -1

        End With
            ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, verts(0), Len(verts(0))


End Sub

