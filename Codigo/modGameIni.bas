Attribute VB_Name = "modGameIni"
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
' modGameIni
'
' @remarks Operaciones de Cabezera y inicio.con
' @author unkwown
' @version 0.0.01
' @date 20060520

Option Explicit
Public TempFromReal(0 To 10000) As Integer
Public TempFromReale(0 To 10000) As Integer

Public Type tAura
    GrhIndex As Integer
    R As Byte
    G As Byte
    B As Byte
    A As Byte
    OffsetX As Integer
    OffsetY As Integer
    Giratoria As Byte
    Velocidad As Single
    Tipo As Byte
End Type

Public Type tAuraGrh
    fC As Single
    Count As Single
End Type
Public AuraDATA() As tAura
Private Type tGrafHandler
    lStart As Long
    lSize As Long
    data() As Byte
    Used As Boolean
    file As Byte
End Type
Public NumRealIndex As Integer
Public NumRealEstatic As Integer

Public Graphic_Handler() As tGrafHandler

Private Type tIXAR
    Le As Long
    St As Long
End Type
Public IXAR(1 To 11) As tIXAR
Public Type tCabecera 'Cabecera de los con
    Desc As String * 255
    CRC As Long
    MagicWord As Long
End Type

Public Type tGameIni
    Puerto As Long
    Musica As Byte
    fX As Byte
    tip As Byte
    Password As String
    Name As String
    DirGraficos As String
    DirSonidos As String
    DirMusica As String
    DirMapas As String
    NumeroDeBMPs As Long
    NumeroMapas As Integer
End Type

Public MiCabecera As tCabecera
Public Config_Inicio As tGameIni

Public Sub IniciarCabecera(ByRef Cabecera As tCabecera)
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************
Cabecera.Desc = "Argentum Online by Noland Studios. Copyright Noland-Studios 2001, pablomarquez@noland-studios.com.ar"
Cabecera.CRC = Rnd * 100
Cabecera.MagicWord = Rnd * 10
End Sub

Public Function LeerGameIni() As tGameIni
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************
Dim n As Integer
Dim GameIni As tGameIni
n = FreeFile
Open DirIndex & "Inicio.con" For Binary As #n
Get #n, , MiCabecera

Get #n, , GameIni

Close #n
LeerGameIni = GameIni
End Function

Public Sub EscribirGameIni(ByRef GameIniConfiguration As tGameIni)
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************
Dim n As Integer
n = FreeFile
Open DirIndex & "Inicio.con" For Binary As #n
Put #n, , MiCabecera
GameIniConfiguration.Password = "DAMMLAMERS!"
Put #n, , GameIniConfiguration
Close #n
End Sub
Public Sub Load_DecorData()
Dim dFile As String
Dim k As Long
Dim j As Long
dFile = App.PATH & "\Resources\Dats\Decor.dat"

numDecor = Val(GetVar(dFile, "INIT", "NumDecors"))

If numDecor > 0 Then
    ReDim DecorData(1 To numDecor)
    For k = 1 To numDecor
        With DecorData(k)
            .DecorType = Val(GetVar(dFile, "DECOR" & k, "DecorType"))
            .MaxHP = Val(GetVar(dFile, "DECOR" & k, "MaxHP"))
            .Atacable = Val(GetVar(dFile, "DECOR" & k, "Atacable"))
            .Clave = Val(GetVar(dFile, "DECOR" & k, "Clave"))
            .Respawn = Val(GetVar(dFile, "DECOR" & k, "Respawn"))
            .Objeto = Val(GetVar(dFile, "DECOR" & k, "Objeto"))
            .value = Val(GetVar(dFile, "DECOR" & k, "Value"))
            .TileH = Val(GetVar(dFile, "DECOR" & k, "TileH"))
            .TileW = Val(GetVar(dFile, "DECOR" & k, "Tilew"))
            For j = 1 To CANT_GRAF_DECOR
                .DecorGrh(j) = Val(ReadField(j, GetVar(dFile, "DECOR" & k, "DecorGrh"), Asc("-")))
            Next j
            .Name = GetVar(dFile, "DECOR" & k, "Name")
            frmMain.lListado(6).AddItem .Name & " # " & k
        End With
    Next k
End If


End Sub
Public Sub Load_DecoKeys()
Dim dFile As String
Dim k As Long

dFile = App.PATH & "\Resources\Dats\DecorKeys.dat"
NumDecorKeys = Val(GetVar(dFile, "INIT", "NumDecorKeyS"))
UltimaDecorKey_Comun = Val(GetVar(dFile, "INIT", "Ultima_Comun"))

If NumDecorKeys > 0 Then
    ReDim DecorKeys(1 To NumDecorKeys)
    For k = 1 To NumDecorKeys
        With DecorKeys(k)
            .Tipo_Objeto = Val(GetVar(dFile, "dKEY" & k, "Tipo_Objeto"))
            .Tipo_Clave = Val(GetVar(dFile, "dKEY" & k, "Tipo_Clave"))
            .Contenedor = Val(GetVar(dFile, "dKEY" & k, "Contenedor"))
            .Clave = Val(GetVar(dFile, "dKEY" & k, "Clave"))
            .X = Val(GetVar(dFile, "dKEY" & k, "X"))
            .Y = Val(GetVar(dFile, "dKEY" & k, "Y"))
        End With
    Next k
End If
End Sub
Public Sub Save_DecoKeys(ByVal n As Integer)
Dim dFile As String
Dim j As Long
dFile = App.PATH & "\Resources\Dats\DecorKeys.dat"

WriteVar dFile, "INIT", "ultima_Comun", CStr(UltimaDecorKey_Comun)
WriteVar dFile, "INIT", "NumDecorKeys", CStr(NumDecorKeys)


If n > 0 Then
    If n <= NumDecorKeys Then
        
        With DecorKeys(n)
            WriteVar dFile, "dKEY" & n, "Tipo_Objeto", CStr(.Tipo_Objeto)
            WriteVar dFile, "dKEY" & n, "Tipo_Clave", CStr(.Tipo_Clave)
            WriteVar dFile, "dKEY" & n, "Contenedor", CStr(.Contenedor)
            WriteVar dFile, "dKEY" & n, "Clave", CStr(.Clave)
            WriteVar dFile, "dKEY" & n, "X", CStr(.X)
            WriteVar dFile, "dKEY" & n, "X", CStr(.Y)

        End With
    End If
    
Else
    If NumDecorKeys > 0 Then

    For j = 1 To NumDecorKeys
        With DecorKeys(j)
            WriteVar dFile, "dKEY" & j, "Tipo_Objeto", CStr(.Tipo_Objeto)
            WriteVar dFile, "dKEY" & j, "Tipo_Clave", CStr(.Tipo_Clave)
            WriteVar dFile, "dKEY" & j, "Contenedor", CStr(.Contenedor)
            WriteVar dFile, "dKEY" & j, "Clave", CStr(.Clave)
            WriteVar dFile, "dKEY" & j, "X", CStr(.X)
            WriteVar dFile, "dKEY" & j, "X", CStr(.Y)

        End With

    Next j
    
    End If
End If

End Sub

Public Sub nLoad_NewBodys(ByVal k As Integer, ByVal S As Long)

Dim i As Long
Dim z As Long
Dim X As Integer
Get k, S + 2, NumNewBodys

If NumNewBodys > 0 Then
ReDim BodyData(1 To NumNewBodys)
For i = 1 To NumNewBodys

    With BodyData(i)


            Get k, , .bContinuo
            Get k, , .bReposo
            Get k, , .bAtaque
            Get k, , .bAtacado
            Get k, , .bReposo
            Get k, , .bDeath
            Get k, , .OverWriteGrafico
            Get k, , .OffsetY
            Get k, , .Capa
            For z = 1 To 4
                
                Get k, , X '.mMovement(z)
                If X > 0 Then .mMovement(z) = X 'NewAnimationData(X)

            Next z
            If .bReposo Then
                For z = 1 To 4
                    Get k, , X '.Reposo(z)
                    .Reposo(z) = X 'NewAnimationData(X)
                Next z
            End If
            If .bAtacado Then
                For z = 1 To 4
                    Get k, , X '.Attacked(z)
                    .Attacked(z) = X ' NewAnimationData(X)
                Next z
            End If
            If .bAtaque Then
                For z = 1 To 4
                    Get k, , X '.Attack(z)
                    .Attack(z) = X ' NewAnimationData(X)
                Next z
            End If
            If .bDeath Then
                For z = 1 To 4
                    Get k, , X '.Death(z)
                    .Death(z) = X 'NewAnimationData(X)
                Next z
            End If

    End With
Next i
End If

End Sub
Public Sub LoadIndexData()

   Dim k As Integer
   Dim o As Integer
    
   k = FreeFile

   Open App.PATH & "\RESOURCES\INIT\INDEX.BIN" For Binary Access Read Lock Read As k
       For o = 1 To 11
           Get k, , IXAR(o).St
           Get k, , IXAR(o).Le
       Next o
       

       'nCargarFxs8 k, IXAR(1).St
       nLoad_NewAnimation k, IXAR(2).St
       nLoad_NewBodys k, IXAR(3).St
       'Mod_Indexacion.nLoad_NewShields k, IXAR(4).St
       'Mod_Indexacion.nLoad_NewWeapons k, IXAR(5).St
       nLoad_NewIndex k, IXAR(6).St
       nLoad_NewEstatics k, IXAR(7).St
       'Mod_Indexacion.nLoad_NewHelmets k, IXAR(8).St
       'Mod_Indexacion.nLoad_NewMuniciones k, IXAR(9).St
       'Mod_Indexacion.nLoad_NewCapas k, IXAR(10).St
       nLoad_NewHeads k, IXAR(11).St
       modGameIni.LoadTempIndex
       
       
   Close k

End Sub
Public Sub nLoad_NewEstatics(ByVal k As Integer, ByVal S As Long)

Dim i As Long
Dim z As Long

Get k, S + 2, numNewEstatic

If numNewEstatic > 0 Then
ReDim EstaticData(1 To numNewEstatic)
For i = 1 To numNewEstatic

    With EstaticData(i)
            Get k, , .L
            Get k, , .t
            Get k, , .W
            Get k, , .H
            Get k, , .tw
            Get k, , .th
    End With
Next i
End If

End Sub
Public Sub nLoad_NewIndex(ByVal k As Integer, ByVal S As Long)

Dim i As Long
Dim z As Long

Get k, S + 2, numNewIndex
If numNewIndex > 0 Then
ReDim NewIndexData(1 To numNewIndex)
For i = 1 To numNewIndex
    TempFromReal(i) = i
    With NewIndexData(i)
            Get k, , .Dinamica
            Get k, , .Estatic
            Get k, , .OverWriteGrafico
            frmMain.lListado(5).AddItem i & " - [" & .OverWriteGrafico & "]"
    End With
Next i
End If

End Sub
Public Sub nLoad_NewHeads(ByVal k As Integer, ByVal S As Long)

Dim i As Long
Dim z As Long
Dim X As Integer
Get k, S + 2, Num_Heads

If Num_Heads > 0 Then
ReDim HeadData(1 To Num_Heads)
For i = 1 To Num_Heads

    With HeadData(i)
                Get k, , .OffsetDibujoY
                Get k, , .OffsetOjos
                Get k, , .Raza
                Get k, , .Genero
            For z = 1 To 4
                Get k, , .Frame(z)
            Next z
    End With
Next i
End If


End Sub
Public Sub nLoad_NewAnimation(ByVal k As Integer, ByVal S As Long)

Dim i As Long
Dim p As Long

Dim GrafCounter As Integer

Get k, S + 2, Num_NwAnim


If Num_NwAnim < 1 Then Exit Sub

ReDim NewAnimationData(1 To Num_NwAnim)

For i = 1 To Num_NwAnim
With NewAnimationData(i)
            .Numero = i
            
            Get k, , .Grafico
            Get k, , .Filas
            Get k, , .Columnas
            Get k, , .Height
            Get k, , .Width
            Get k, , .NumFrames
            Get k, , .Velocidad
            Get k, , .TileWidth
            Get k, , .TileHeight
            Get k, , .Romboidal
            Get k, , .OffsetX
            Get k, , .OffsetY
            Get k, , .TipoAnimacion
            ReDim .Indice(1 To .NumFrames)
            If .TipoAnimacion = 0 Then
                For p = 1 To .NumFrames
                    Get k, , .Indice(p).Grafico
                    Get k, , .Indice(p).X
                    Get k, , .Indice(p).Y
                Next p
            ElseIf .TipoAnimacion = 1 Then
                .NumFrames = .NumFrames + 2
                ReDim .Indice(1 To .NumFrames)
                For p = 1 To .NumFrames
                    If p <> 4 And p <> 8 Then
                        Get k, , .Indice(p).Grafico
                        Get k, , .Indice(p).X
                        Get k, , .Indice(p).Y
                    ElseIf p = 4 Then
                        .Indice(p).Grafico = .Indice(2).Grafico
                        .Indice(p).X = .Indice(2).X
                        .Indice(p).Y = .Indice(2).Y
                    ElseIf p = 8 Then
                        .Indice(p).Grafico = .Indice(6).Grafico
                        .Indice(p).X = .Indice(6).X
                        .Indice(p).Y = .Indice(6).Y
                    End If
                Next p
            End If
            
End With
Next i



End Sub
Public Sub Load_Graphic_Header()

Dim FF As Integer
Dim f2 As Integer
Dim data() As Byte
Dim Data2() As Byte
Dim index As Integer
FF = FreeFile


Open App.PATH & "\Resources\MPG_1.bin" For Binary Access Read Lock Read As #FF
        
        ReDim Graphic_Handler(0 To 5000) As tGrafHandler
        
        Do Until EOF(FF)
            
            Get FF, , index
            
            Graphic_Handler(index).Used = True
            
            Get FF, , Graphic_Handler(index).lStart
            Graphic_Handler(index).lStart = Graphic_Handler(index).lStart + 27
            
            Get FF, , Graphic_Handler(index).lSize
            
            Graphic_Handler(index).lSize = Graphic_Handler(index).lSize + 88
            
            Graphic_Handler(index).lStart = Graphic_Handler(index).lStart + 1
            
        Loop
    
    
        
Close #FF
    

    
End Sub
Public Function ExtractGraphic(ByVal Indice As Integer, ByRef data() As Byte) As Boolean
On Error GoTo Err
Dim FF As Integer
Dim Datax() As Byte
FF = FreeFile

    ReDim data(0 To Graphic_Handler(Indice).lSize + 24) As Byte
    ReDim Datax(0 To Graphic_Handler(Indice).lSize - 1) As Byte
    Open App.PATH & "\Resources\MPG_2.bin" For Binary Access Read Lock Read As #FF

        Get FF, Graphic_Handler(Indice).lStart, Datax
        CopyMemory data(0), PngInit(0), 17
        CopyMemory data(17), Datax(0), Graphic_Handler(Indice).lSize
        CopyMemory data(Graphic_Handler(Indice).lSize + 17), PngEnd(0), 7

    Close #FF

    ExtractGraphic = True
    Exit Function
Err:
    
    
End Function
Public Sub Load_Png_Array()

   PngInit(0) = 137
   PngInit(1) = 80
   PngInit(2) = 78
   PngInit(3) = 71
   PngInit(4) = 13
   PngInit(5) = 10
   PngInit(6) = 26
   PngInit(7) = 10
   PngInit(8) = 0
   PngInit(9) = 0
   PngInit(10) = 0
   PngInit(11) = 13
   PngInit(12) = 73
   PngInit(13) = 72
   PngInit(14) = 68
   PngInit(15) = 82
   PngInit(16) = 0
   
   PngEnd(0) = 69
   PngEnd(1) = 78
   PngEnd(2) = 68
   PngEnd(3) = 174
   PngEnd(4) = 66
   PngEnd(5) = 96
   PngEnd(6) = 130

End Sub

Public Sub CargarEfectos()
Dim FF As Integer

FF = FreeFile
Open App.PATH & "\RESOURCES\INIT\Efectos.bin" For Binary Access Read Lock Read As #FF
    CargarAurasBin FF
    CargarParticulasBin FF
    'CargarBuffdataBin FF
    'SPOTLIGHTS_LOADDATA FF
Close #FF


End Sub
Sub CargarAurasBin(ByVal FF As Integer)
Dim n As Integer

    Get FF, , n
    ReDim AuraDATA(0 To n) As tAura
    Get FF, , AuraDATA
End Sub
Sub CargarParticulasBin(ByVal FF As Integer)
Dim StreamFile As String
Dim LoopC As Long
Dim i As Long
Dim GrhListing As String
Dim TempSet As String
Dim ColorSet As Long

Get FF, , TotalStreams

'resize StreamData array
ReDim StreamData(1 To TotalStreams) As Stream
 
    'fill StreamData array with info from Particles.ini
    For LoopC = 1 To TotalStreams
        Get FF, , StreamData(LoopC).NumOfParticles
        StreamData(LoopC).NumTrueParticles = StreamData(LoopC).NumOfParticles
        
                
        Get FF, , StreamData(LoopC).x1
        Get FF, , StreamData(LoopC).y1 '= GetVar(StreamFile, Val(loopc), "Y1")
        Get FF, , StreamData(LoopC).x2 '= GetVar(StreamFile, Val(loopc), "X2")
        Get FF, , StreamData(LoopC).y2 '= GetVar(StreamFile, Val(loopc), "Y2")
        Get FF, , StreamData(LoopC).Angle '= GetVar(StreamFile, Val(loopc), "Angle")
        Get FF, , StreamData(LoopC).vecx1 '= GetVar(StreamFile, Val(loopc), "VecX1")
        Get FF, , StreamData(LoopC).vecx2 '= GetVar(StreamFile, Val(loopc), "VecX2")
        Get FF, , StreamData(LoopC).vecy1 '= GetVar(StreamFile, Val(loopc), "VecY1")
        Get FF, , StreamData(LoopC).vecy2 '= GetVar(StreamFile, Val(loopc), "VecY2")
        Get FF, , StreamData(LoopC).life1 '= GetVar(StreamFile, Val(loopc), "Life1")
        Get FF, , StreamData(LoopC).life2 '= GetVar(StreamFile, Val(loopc), "Life2")
        Get FF, , StreamData(LoopC).friction '= GetVar(StreamFile, Val(loopc), "Friction")
        Get FF, , StreamData(LoopC).Spin '= GetVar(StreamFile, Val(loopc), "Spin")
        Get FF, , StreamData(LoopC).spin_speedL '= GetVar(StreamFile, Val(loopc), "Spin_SpeedL")
        Get FF, , StreamData(LoopC).spin_speedH '= GetVar(StreamFile, Val(loopc), "Spin_SpeedH")
        Get FF, , StreamData(LoopC).AlphaBlend '= GetVar(StreamFile, Val(loopc), "AlphaBlend")
        Get FF, , StreamData(LoopC).gravity '= GetVar(StreamFile, Val(loopc), "Gravity")
        Get FF, , StreamData(LoopC).grav_strength '= GetVar(StreamFile, Val(loopc), "Grav_Strength")
        Get FF, , StreamData(LoopC).bounce_strength '= GetVar(StreamFile, Val(loopc), "Bounce_Strength")
        Get FF, , StreamData(LoopC).XMove '= GetVar(StreamFile, Val(loopc), "XMove")
        Get FF, , StreamData(LoopC).YMove '= GetVar(StreamFile, Val(loopc), "YMove")
        Get FF, , StreamData(LoopC).move_x1 '= GetVar(StreamFile, Val(loopc), "move_x1")
        Get FF, , StreamData(LoopC).move_x2 '= GetVar(StreamFile, Val(loopc), "move_x2")
        Get FF, , StreamData(LoopC).move_y1 '= GetVar(StreamFile, Val(loopc), "move_y1")
        Get FF, , StreamData(LoopC).move_y2 '= GetVar(StreamFile, Val(loopc), "move_y2")
        Get FF, , StreamData(LoopC).life_counter '= GetVar(StreamFile, Val(loopc), "life_counter")
        Get FF, , StreamData(LoopC).Speed '= Val(GetVar(StreamFile, Val(loopc), "Speed"))
        Get FF, , StreamData(LoopC).grh_resize '= Val(GetVar(StreamFile, Val(loopc), "resize"))
        Get FF, , StreamData(LoopC).grh_resizex '= Val(GetVar(StreamFile, Val(loopc), "rx"))
        Get FF, , StreamData(LoopC).grh_resizey '= Val(GetVar(StreamFile, Val(loopc), "ry"))
        Get FF, , StreamData(LoopC).NumGrhs '= GetVar(StreamFile, Val(loopc), "NumGrhs")
       
       ReDim StreamData(LoopC).grh_list(1 To StreamData(LoopC).NumGrhs)
       
      
       For i = 1 To StreamData(LoopC).NumGrhs
           Get FF, , StreamData(LoopC).grh_list(i)
       Next i
       
        For ColorSet = 1 To 4
            Get FF, , StreamData(LoopC).colortint(ColorSet - 1).R
            Get FF, , StreamData(LoopC).colortint(ColorSet - 1).G
            Get FF, , StreamData(LoopC).colortint(ColorSet - 1).B
        Next ColorSet
    Next LoopC
 
End Sub
Public Sub CargarTipoTerrenos()
With frmMain.lListado(8)
    .AddItem "Nada #0"
    .AddItem "Agua #1"
    .AddItem "Lava #2"
End With
End Sub
Public Sub LoadTempIndex()
Dim j As Long
Dim p As String
p = App.PATH & "\Resources\InitTemp\TempIndex.Dat"

ntIndex = Val(GetVar(p, "INIT", "NUM"))
If ntIndex > 0 Then
ReDim TempIndex(1 To ntIndex)
For j = 1 To ntIndex
    With TempIndex(j)
            .OverWriteGrafico = Val(GetVar(p, j, "OverWriteGrafico"))
            .temp = Val(GetVar(p, j, "Temp"))
            .Estatic = Val(GetVar(p, j, "Estatica"))
            .Dinamica = Val(GetVar(p, j, "Dinamica"))
            .Replace = Val(GetVar(p, j, "Reemplazo"))
    End With
Next j
End If
p = App.PATH & "\Resources\InitTemp\TempEstatics.Dat"

ntEstatic = Val(GetVar(p, "INIT", "NUM"))
If ntEstatic > 0 Then
ReDim TempEstatic(1 To ntEstatic)

For j = 1 To ntEstatic
    With TempEstatic(j)
        .L = Val(GetVar(p, j, "Left"))
        .t = Val(GetVar(p, j, "Top"))
        .W = Val(GetVar(p, j, "Width"))
        .H = Val(GetVar(p, j, "Height"))
        .tw = .W / 32
        .th = .H / 32
        .Replace = Val(GetVar(p, j, "Reemplazo"))
    End With
Next j
End If
NumRealEstatic = numNewEstatic
If ntEstatic > 0 Then
ReDim Preserve EstaticData(1 To numNewEstatic + ntEstatic)
For j = numNewEstatic + 1 To numNewEstatic + ntEstatic
    With EstaticData(j)
        .L = TempEstatic(j - numNewEstatic).L
        .t = TempEstatic(j - numNewEstatic).t
        .W = TempEstatic(j - numNewEstatic).W
        .H = TempEstatic(j - numNewEstatic).H
        .tw = TempEstatic(j - numNewEstatic).tw
        .th = TempEstatic(j - numNewEstatic).th
        TempFromReale(j) = j - numNewEstatic
    End With
Next j
numNewEstatic = numNewEstatic + ntEstatic
End If



NumRealIndex = numNewIndex

If ntIndex > 0 Then
ReDim Preserve NewIndexData(1 To numNewIndex + ntIndex)
For j = numNewIndex + 1 To numNewIndex + ntIndex
    With NewIndexData(j)
        If TempIndex(j - NumRealIndex).temp = 1 Then
            .Estatic = TempIndex(j - numNewIndex).Estatic + NumRealEstatic
        Else
            .Estatic = TempIndex(j - numNewIndex).Estatic
        End If
        .Dinamica = TempIndex(j - numNewIndex).Dinamica
        .OverWriteGrafico = TempIndex(j - numNewIndex).OverWriteGrafico
        TempFromReal(j) = j - numNewIndex
        frmMain.lListado(5).AddItem j & " - [" & .OverWriteGrafico & "]"
    End With
Next j

numNewIndex = numNewIndex + ntIndex
End If

End Sub
