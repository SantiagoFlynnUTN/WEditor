Attribute VB_Name = "modMapIO"
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
' modMapIO
'
' @remarks Funciones Especificas al trabajo con Archivos de Mapas
' @author gshaxor@gmail.com
' @version 0.1.15
' @date 20060602

Option Explicit
Public Const INITIAL_INVERT_MAPLIGHT As Byte = 125
Public MapaTemporal As Boolean

Public Type tnwMapBlock
    Layer(1 To 5) As Integer
    IndexB(1 To 5) As Byte

    TipoTerreno As Byte
    Trigger As Integer
    InteriorVal As Byte
    
    InteriorNum As Integer
    UltimoInteriorX As Integer
    Luces(0 To 3) As Byte
End Type



Private Type MAPLAYER5BUFFER
    posi As tcardinal
    Graf As Integer
    index As Byte
End Type

Private Type MAPSPOTBUFFER
    posi As tcardinal
    SPOT As tSPOT_LIGHTSmap
End Type
Private Type MAPPARTBUFFER
    posi As tcardinal
    PARTI As Byte
End Type
Private Type MAPDECSBUFFER
    posi As tcardinal
    DecorI As Integer
    EstadoDefault As Byte
End Type

Private MapTitulo As String     ' GS > Almacena el titulo del mapa para el .dat

''
' Obtener el tamaño de un archivo
'
' @param FileName Especifica el path del archivo
' @return   Nos devuelve el tamaño

Public Function FileSize(ByVal FileName As String) As Long
    '*************************************************
    'Author: Unkwown
    'Last modified: 20/05/06
    '*************************************************

    On Error GoTo FalloFile
    Dim nFileNum As Integer
    Dim lFileSize As Long
    
    nFileNum = FreeFile
    Open FileName For Input As nFileNum
    lFileSize = LOF(nFileNum)
    Close nFileNum
    FileSize = lFileSize
    
    Exit Function
FalloFile:
    FileSize = -1
End Function

''
' Nos dice si existe el archivo/directorio
'
' @param file Especifica el path
' @param FileType Especifica el tipo de archivo/directorio
' @return   Nos devuelve verdadero o falso

Public Function General_File_Exist(ByVal file As String, ByVal FileType As VbFileAttribute) As Boolean
    '*************************************************
    'Author: Unkwown
    'Last modified: 26/05/06
    '*************************************************
    If LenB(Dir(file, FileType)) = 0 Then
        General_File_Exist = False
    Else
        General_File_Exist = True
    End If

End Function

''
' Abre un Mapa
'
' @param Path Especifica el path del mapa



''
' Guarda el Mapa
'
' @param Path Especifica el path del mapa

Public Sub GuardarMapa(Optional PATH As String)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 01/11/08
    '*************************************************

    frmMain.Dialog.CancelError = True
    On Error GoTo ErrHandler

    If LenB(PATH) = 0 Then
        frmMain.ObtenerNombreArchivo True
        PATH = frmMain.Dialog.FileName
        If LenB(PATH) = 0 Then Exit Sub
    End If
    'If Tipo = 1 Then
    'Call MapaV2_Guardar(PATH)
    'ElseIf Tipo = 2 Then

    'Dim p As Long
    'For p = 1 To 160 '
    '   ReDim MapData(1 To 100, 1 To 100)
    '    Erase CharList
    '    modMapIO.AbrirMapaComun "C:\Maps\Nuevos Mapas\" & p & ".map"
    '    modMapIO.Guardar_Nuevo_Mapa "C:\MAPS\NUEVO_FORMATO2\" & p & ".map"
    'Next p

    modMapIO.Guardar_Nuevo_Mapa PATH

    'End If
    Exit Sub
ErrHandler:
    Debug.Print Err.Description

End Sub
Public Sub GuardarMapaI(Optional PATH As String)

    frmMain.Dialog.CancelError = True
    On Error GoTo ErrHandler

    If LenB(PATH) = 0 Then
        frmMain.ObtenerNombreArchivo True
        PATH = frmMain.Dialog.FileName
        If LenB(PATH) = 0 Then Exit Sub
    End If

    MapaInterface_Guardar PATH

ErrHandler:
End Sub
Public Sub GuardarMapaIn(Optional PATH As String)

    frmMain.Dialog.CancelError = True
    On Error GoTo ErrHandler

    If LenB(PATH) = 0 Then
        frmMain.ObtenerNombreArchivo True
        PATH = frmMain.Dialog.FileName
        If LenB(PATH) = 0 Then Exit Sub
    End If

    MapaInterfacen_Guardar PATH

ErrHandler:
End Sub



''
' Nos pregunta donde guardar el mapa en caso de modificarlo
'
' @param Path Especifica si existiera un path donde guardar el mapa

Public Sub DeseaGuardarMapa(Optional PATH As String)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************

    If MapInfo.Changed = 1 Then
        If MsgBox(MSGMod, vbExclamation + vbYesNo) = vbYes Then
            Guardar_Nuevo_Mapa PATH
        End If
    End If
End Sub

Public Sub NuevoMapa()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 21/05/06
    '*************************************************

    On Error GoTo erh

    Dim LoopC As Integer

    bAutoGuardarMapaCount = 0

    'frmMain.mnuUtirialNuevoFormato.Checked = True

    frmMain.TimAutoGuardarMapa.Enabled = False


    MapaCargado = False

    For LoopC = 0 To frmMain.MapPest.Count - 1
        frmMain.MapPest(LoopC).Enabled = False
    Next

    frmMain.MousePointer = 11

    ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock

    For LoopC = 1 To LastChar
        If CharList(LoopC).Active = 1 Then Call EraseChar(LoopC)
    Next LoopC
    Dim X As Long
    Dim Y As Long
    Dim nc As Integer
    Dim nf As Integer
    Dim z As Single
    Dim tw As Integer
    Dim th As Integer
    Dim k As Long
    Dim TempX As Integer
    Dim TempY As Integer

    If NuevoTex > 0 Then
        With TexWE(NuevoTex)
            z = (.Largo / 32)
            th = (.Largo \ 32)
            If z > th Then th = th + 1
            z = (.Ancho / 32)
            tw = (.Ancho \ 32)
            If z > tw Then tw = tw + 1
            z = 100 / th
            nf = 100 \ th
            If z > nf Then nf = nf + 1
            z = 100 \ tw
            nc = 100 \ tw
            If z > nc Then nc = nc + 1
       
            For Y = 1 To nf
                TempY = (Y - 1) * th
                For X = 1 To nc
                    TempX = (X - 1) * tw
                    For k = 1 To .NumIndex
                        If (.index(k).X + (TempX * 32)) < 3200 And (.index(k).Y + (TempY * 32)) < 3200 Then
                            If (.index(k).X + (TempX * 32) + EstaticData(NewIndexData(.index(k).Num).Estatic).W) <= 3200 And (.index(k).Y + (TempY * 32) + EstaticData(NewIndexData(.index(k).Num).Estatic).H) <= 3200 Then
                                MapData((TempX + (.index(k).X \ 32) + 1), (TempY + (.index(k).Y \ 32)) + 1).Graphic(1).index = .index(k).Num
                            
                            
                            End If
                        End If
                    Next k



                Next X
            Next Y
        End With
    End If
    MapaTemporal = False

    MapInfo.MapVersion = 0
    MapInfo.Name = "Nuevo Mapa"
    MapInfo.Music = 0
    MapInfo.PK = True
    MapInfo.MagiaSinEfecto = 0
    MapInfo.Terreno = "BOSQUE"
    MapInfo.Zona = "CAMPO"
    MapInfo.Restringir = "NO"
    MapInfo.NoEncriptarMP = 0

    Call MapInfo_Actualizar
    modDXEngine.Particle_Group_Remove_All
    bRefreshRadar = True ' Radar

    'Set changed flag
    MapInfo.Changed = 0
    frmMain.MousePointer = 0

    ' Vacio deshacer
    modEdicion.Deshacer_Clear

    MapaCargado = True

    frmMain.SetFocus


    Exit Sub

erh:     MsgBox Err.Description & ":" & Y & ":" & X & ":" & k
End Sub


Public Sub MapaV2_Guardar(ByVal SaveAs As String)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************

    On Error GoTo ErrorSave
    Dim FreeFileMap As Long
    Dim freefileinf As Long
    Dim LoopC As Long
    Dim TempInt As Integer
    Dim Y As Long
    Dim X As Long
    Dim ByFlags As Byte
    Dim Interiores_Number As Integer
    Dim R As Byte
    Dim G As Byte
    Dim B As Byte
    Dim k As Byte
    k = 254

    If General_File_Exist(SaveAs, vbNormal) = True Then
        If MsgBox("¿Desea sobrescribir " & SaveAs & "?", vbCritical + vbYesNo) = vbNo Then
            Exit Sub
        Else
            Kill SaveAs
        End If
    End If

    frmMain.MousePointer = 11

    'Precalculamos interiores.
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            If MapData(X, Y).InteriorVal > 0 Then
                Interiores_Number = Interiores_Number + 1
            End If
        Next X
    Next Y

    ' y borramos el .inf tambien
    If General_File_Exist(left$(SaveAs, Len(SaveAs) - 4) & ".inf", vbNormal) = True Then
        Kill left$(SaveAs, Len(SaveAs) - 4) & ".inf"
    End If

    'Open .map file
    FreeFileMap = FreeFile
    Open SaveAs For Binary As FreeFileMap
    Seek FreeFileMap, 1

    SaveAs = left$(SaveAs, Len(SaveAs) - 4)
    SaveAs = SaveAs & ".inf"

    'Open .inf file
    freefileinf = FreeFile
    Open SaveAs For Binary As freefileinf


    'map Header
    

    Put FreeFileMap, , TempInt
    Put FreeFileMap, , MiCabecera
    Put FreeFileMap, , MapHead.Version
    Put FreeFileMap, , MapHead.SombrasAmbientales
    Put FreeFileMap, , MapHead.GraficoMapa
    Put FreeFileMap, , Interiores_Number
    
    'inf Header
    Seek freefileinf, 1
    Put freefileinf, , TempInt
    Put freefileinf, , TempInt
    Put freefileinf, , TempInt
    Put freefileinf, , TempInt
    Put freefileinf, , TempInt
    

            
    
    'Write .map file
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            
            ByFlags = 0
                
            If MapData(X, Y).Blocked = 1 Then ByFlags = ByFlags Or 1
            If MapData(X, Y).Graphic(2).index Then ByFlags = ByFlags Or 2
            If MapData(X, Y).Graphic(3).index Then ByFlags = ByFlags Or 4
            If MapData(X, Y).Graphic(4).index Then ByFlags = ByFlags Or 8
            If MapData(X, Y).Trigger Then ByFlags = ByFlags Or 16
            If MapData(X, Y).Luz > 0 Then ByFlags = ByFlags Or 32
            If MapData(X, Y).particle_group Or MapData(X, Y).SPOTLIGHT.index > 0 Then ByFlags = ByFlags Or 64
            If MapData(X, Y).InteriorVal > 0 Then ByFlags = ByFlags Or 128
                
            'If MapData(X, Y).light_index Then ByFlags = ByFlags Or 64
            'If MapData(X, Y).AlturaPoligonos(0) Or MapData(X, Y).AlturaPoligonos(1) _
            '    Or MapData(X, Y).AlturaPoligonos(2) Or MapData(X, Y).AlturaPoligonos(3) Then ByFlags = ByFlags Or 128
            Put FreeFileMap, , ByFlags
                
            Put FreeFileMap, , MapData(X, Y).Graphic(1).index
                
            For LoopC = 2 To 4
                If MapData(X, Y).Graphic(LoopC).index Then _
                    Put FreeFileMap, , MapData(X, Y).Graphic(LoopC).index
            Next LoopC
                
            If MapData(X, Y).Trigger Then _
                Put FreeFileMap, , MapData(X, Y).Trigger
                
                
            If MapData(X, Y).Luz > 100 Then
                Put FreeFileMap, , MapData(X, Y).Luz
                Put FreeFileMap, , MapData(X, Y).LV
            ElseIf MapData(X, Y).Luz > 0 Then
                Put FreeFileMap, , MapData(X, Y).Luz
            End If
                    
            If MapData(X, Y).SPOTLIGHT.index > 0 Then
                If MapData(X, Y).particle_group Then
                    Put FreeFileMap, , CByte(k + 1)
                    Put FreeFileMap, , MapData(X, Y).parti_index
                        
                Else
                    Put FreeFileMap, , k
                End If
                Put FreeFileMap, , MapData(X, Y).SPOTLIGHT.SPOT_TIPO
                If MapData(X, Y).SPOTLIGHT.SPOT_TIPO = 0 Then
                    Put FreeFileMap, , MapData(X, Y).SPOTLIGHT.Grafico
                End If
                Put FreeFileMap, , MapData(X, Y).SPOTLIGHT.SPOT_COLOR_BASE
                If MapData(X, Y).SPOTLIGHT.SPOT_COLOR_BASE = 99 Then
                    Put FreeFileMap, , MapData(X, Y).SPOTLIGHT.Color
                End If
                Put FreeFileMap, , MapData(X, Y).SPOTLIGHT.INTENSITY
                Put FreeFileMap, , MapData(X, Y).SPOTLIGHT.OffsetX
                Put FreeFileMap, , MapData(X, Y).SPOTLIGHT.OffsetY
                Put FreeFileMap, , MapData(X, Y).SPOTLIGHT.SPOT_COLOR_EXTRA
                If MapData(X, Y).SPOTLIGHT.SPOT_COLOR_EXTRA = 99 Then
                    Put FreeFileMap, , MapData(X, Y).SPOTLIGHT.COLOR_EXTRA
                End If
                Put FreeFileMap, , MapData(X, Y).SPOTLIGHT.EXTRA_GRAFICO
            Else
                If MapData(X, Y).particle_group Then _
                    Put FreeFileMap, , MapData(X, Y).parti_index
            End If
                
            If MapData(X, Y).InteriorVal > 0 Then Put FreeFileMap, , MapData(X, Y).InteriorVal
                
            'If MapData(X, Y).light_index Then
            '    Put FreeFileMap, , Lights(MapData(X, Y).light_index).Range
            '    R = Lights(MapData(X, Y).light_index).RGBCOLOR.R
            '    G = Lights(MapData(X, Y).light_index).RGBCOLOR.G
            '    B = Lights(MapData(X, Y).light_index).RGBCOLOR.B
            '    Put FreeFileMap, , R
            '    Put FreeFileMap, , G
            '    Put FreeFileMap, , B
            'End If
                
            'If MapData(X, Y).AlturaPoligonos(0) Or MapData(X, Y).AlturaPoligonos(1) _
            '    Or MapData(X, Y).AlturaPoligonos(2) Or MapData(X, Y).AlturaPoligonos(3) Then
            '    Put FreeFileMap, , MapData(X, Y).AlturaPoligonos(0)
            '    Put FreeFileMap, , MapData(X, Y).AlturaPoligonos(1)
            '    Put FreeFileMap, , MapData(X, Y).AlturaPoligonos(2)
            '   Put FreeFileMap, , MapData(X, Y).AlturaPoligonos(3)
            '
            ''   If MapData(X, Y).AlturaPoligonos(0) Then _
            '       Put FreeFileMap, , MapData(X, Y).AlturaPoligonos(0)
            '   If MapData(X, Y).AlturaPoligonos(1) Then _
            ''        Put FreeFileMap, , MapData(X, Y).AlturaPoligonos(1)
            '    If MapData(X, Y).AlturaPoligonos(2) Then _
            '         Put FreeFileMap, , MapData(X, Y).AlturaPoligonos(2)
            '     If MapData(X, Y).AlturaPoligonos(3) Then _
            '         Put FreeFileMap, , MapData(X, Y).AlturaPoligonos(3)
            ' End If
                
            '.inf file
                
            ByFlags = 0
                
            If MapData(X, Y).TileExit.Map Then ByFlags = ByFlags Or 1
            If MapData(X, Y).NPCIndex Then ByFlags = ByFlags Or 2
            If MapData(X, Y).OBJInfo.objindex Then ByFlags = ByFlags Or 4
                
            Put freefileinf, , ByFlags
                
            If MapData(X, Y).TileExit.Map Then
                Put freefileinf, , MapData(X, Y).TileExit.Map
                Put freefileinf, , MapData(X, Y).TileExit.X
                Put freefileinf, , MapData(X, Y).TileExit.Y
            End If
                
            If MapData(X, Y).NPCIndex Then
                
                Put freefileinf, , CInt(MapData(X, Y).NPCIndex)
            End If
                
            If MapData(X, Y).OBJInfo.objindex Then
                Put freefileinf, , MapData(X, Y).OBJInfo.objindex
                Put freefileinf, , MapData(X, Y).OBJInfo.Amount
            End If
            
            
            
        Next X
    Next Y
    
    'Close .map file
    Close FreeFileMap
    
    'Close .inf file
    Close freefileinf


    Call Pestañas(SaveAs)

    'write .dat file
    SaveAs = left$(SaveAs, Len(SaveAs) - 4) & ".dat"
    MapInfo_Guardar SaveAs

    'Change mouse icon
    frmMain.MousePointer = 0
    MapInfo.Changed = 0

    Exit Sub

ErrorSave:
    MsgBox "Error en GuardarV2, nro. " & Err.Number & " - " & Err.Description
End Sub
Public Sub MapaInterface_Guardar(ByVal SaveAs As String)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************

    On Error GoTo ErrorSave
    Dim FreeFileMap As Long
    Dim freefileinf As Long
    Dim LoopC As Long
    Dim TempInt As Integer
    Dim Y As Long
    Dim X As Long
    Dim ByFlags As Byte

    Dim R As Byte
    Dim G As Byte
    Dim B As Byte

    If General_File_Exist(SaveAs, vbNormal) = True Then
        If MsgBox("¿Desea sobrescribir " & SaveAs & "?", vbCritical + vbYesNo) = vbNo Then
            Exit Sub
        Else
            Kill SaveAs
        End If
    End If

    frmMain.MousePointer = 11

    ' y borramos el .inf tambien
    If General_File_Exist(left$(SaveAs, Len(SaveAs) - 4) & ".inf", vbNormal) = True Then
        Kill left$(SaveAs, Len(SaveAs) - 4) & ".inf"
    End If

    'Open .map file
    FreeFileMap = FreeFile
    Open SaveAs For Binary As FreeFileMap
    Seek FreeFileMap, 1


    'map Header
    

       
    'Write .map file
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize

            ByFlags = 0
                
            If MapData(X, Y).Blocked = 1 Then ByFlags = ByFlags Or 1
            If MapData(X, Y).Graphic(2).index Then ByFlags = ByFlags Or 2
            If MapData(X, Y).Graphic(3).index Then ByFlags = ByFlags Or 4
            If MapData(X, Y).Graphic(4).index Then ByFlags = ByFlags Or 8
            If MapData(X, Y).OBJInfo.objindex Then ByFlags = ByFlags Or 16
            If MapData(X, Y).Luz > 0 Then ByFlags = ByFlags Or 32
            Put FreeFileMap, , ByFlags
                
            Put FreeFileMap, , MapData(X, Y).Graphic(1).index
                
            For LoopC = 2 To 4
                If MapData(X, Y).Graphic(LoopC).index Then _
                    Put FreeFileMap, , MapData(X, Y).Graphic(LoopC).index
            Next LoopC
                
            If MapData(X, Y).OBJInfo.objindex Then
                Put FreeFileMap, , MapData(X, Y).OBJInfo.objindex
            End If
                
            If MapData(X, Y).Luz Then _
                Put FreeFileMap, , MapData(X, Y).Luz
            If MapData(X, Y).Luz > 100 Then
                    
                    
            End If
                    
            If MapData(X, Y).Trigger = 5 Then GoTo terminar
        Next X
    Next Y
terminar:
    'Close .map file
    Close FreeFileMap

    Call Pestañas(SaveAs)

    'write .dat file
    SaveAs = left$(SaveAs, Len(SaveAs) - 4) & ".dat"

    'Change mouse icon
    frmMain.MousePointer = 0
    MapInfo.Changed = 0

    Exit Sub

ErrorSave:
    MsgBox "Error en GuardarV2, nro. " & Err.Number & " - " & Err.Description
End Sub
Public Sub MapaInterfacen_Guardar(ByVal SaveAs As String)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************

    On Error GoTo ErrorSave
    Dim FreeFileMap As Long
    Dim freefileinf As Long
    Dim LoopC As Long
    Dim TempInt As Integer
    Dim Y As Long
    Dim X As Long
    Dim ByFlags As Byte

    Dim R As Byte
    Dim G As Byte
    Dim B As Byte
    Dim k As Byte
    k = 254
    If General_File_Exist(SaveAs, vbNormal) = True Then
        If MsgBox("¿Desea sobrescribir " & SaveAs & "?", vbCritical + vbYesNo) = vbNo Then
            Exit Sub
        Else
            Kill SaveAs
        End If
    End If

    frmMain.MousePointer = 11

    ' y borramos el .inf tambien
    If General_File_Exist(left$(SaveAs, Len(SaveAs) - 4) & ".inf", vbNormal) = True Then
        Kill left$(SaveAs, Len(SaveAs) - 4) & ".inf"
    End If

    'Open .map file
    FreeFileMap = FreeFile
    Open SaveAs For Binary As FreeFileMap
    Seek FreeFileMap, 1



       
    'Write .map file
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize

            ByFlags = 0
                
            If MapData(X, Y).Blocked = 1 Then ByFlags = ByFlags Or 1
            If MapData(X, Y).Graphic(2).index Then ByFlags = ByFlags Or 2
            If MapData(X, Y).Graphic(3).index Then ByFlags = ByFlags Or 4
            If MapData(X, Y).Graphic(4).index Then ByFlags = ByFlags Or 8
            If MapData(X, Y).OBJInfo.objindex Then ByFlags = ByFlags Or 16
            If MapData(X, Y).Luz > 0 Then ByFlags = ByFlags Or 32
            If MapData(X, Y).particle_group Or MapData(X, Y).SPOTLIGHT.index > 0 Then ByFlags = ByFlags Or 64
            Put FreeFileMap, , ByFlags
                
            Put FreeFileMap, , MapData(X, Y).Graphic(1).index
                
            For LoopC = 2 To 4
                If MapData(X, Y).Graphic(LoopC).index Then _
                    Put FreeFileMap, , MapData(X, Y).Graphic(LoopC).index
            Next LoopC
                
            If MapData(X, Y).OBJInfo.objindex Then
                Put FreeFileMap, , MapData(X, Y).OBJInfo.objindex
            End If

            If MapData(X, Y).Luz > 100 Then
                Put FreeFileMap, , MapData(X, Y).Luz
                Put FreeFileMap, , MapData(X, Y).LV
            ElseIf MapData(X, Y).Luz > 0 Then
                Put FreeFileMap, , MapData(X, Y).Luz
            End If
                    
            If MapData(X, Y).SPOTLIGHT.index > 0 Then
                If MapData(X, Y).particle_group Then
                    Put FreeFileMap, , CByte(k + 1)
                    Put FreeFileMap, , MapData(X, Y).parti_index
                        
                Else
                    Put FreeFileMap, , k
                End If
                Put FreeFileMap, , MapData(X, Y).SPOTLIGHT.SPOT_TIPO
                If MapData(X, Y).SPOTLIGHT.SPOT_TIPO = 0 Then
                    Put FreeFileMap, , MapData(X, Y).SPOTLIGHT.Grafico
                End If
                Put FreeFileMap, , MapData(X, Y).SPOTLIGHT.SPOT_COLOR_BASE
                If MapData(X, Y).SPOTLIGHT.SPOT_COLOR_BASE = 99 Then
                    Put FreeFileMap, , MapData(X, Y).SPOTLIGHT.Color
                End If
                Put FreeFileMap, , MapData(X, Y).SPOTLIGHT.INTENSITY
                Put FreeFileMap, , MapData(X, Y).SPOTLIGHT.OffsetX
                Put FreeFileMap, , MapData(X, Y).SPOTLIGHT.OffsetY
                Put FreeFileMap, , MapData(X, Y).SPOTLIGHT.SPOT_COLOR_EXTRA
                If MapData(X, Y).SPOTLIGHT.SPOT_COLOR_EXTRA = 99 Then
                    Put FreeFileMap, , MapData(X, Y).SPOTLIGHT.COLOR_EXTRA
                End If
                Put FreeFileMap, , MapData(X, Y).SPOTLIGHT.EXTRA_GRAFICO
            Else
                If MapData(X, Y).particle_group Then _
                    Put FreeFileMap, , MapData(X, Y).parti_index
            End If
                    
            If MapData(X, Y).Trigger And 16 Then GoTo terminar
        Next X
    Next Y
terminar:
    'Close .map file
    Close FreeFileMap

    Call Pestañas(SaveAs)

    'write .dat file
    SaveAs = left$(SaveAs, Len(SaveAs) - 4) & ".dat"

    'Change mouse icon
    frmMain.MousePointer = 0
    MapInfo.Changed = 0

    Exit Sub

ErrorSave:
    MsgBox "Error en GuardarV2, nro. " & Err.Number & " - " & Err.Description
End Sub

''
' Guardar Mapa con el formato V1
'
' @param SaveAs Especifica donde guardar el mapa

Public Sub MapaV1_Guardar(SaveAs As String)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************

    On Error GoTo ErrorSave
    Dim FreeFileMap As Long
    Dim freefileinf As Long
    Dim LoopC As Long
    Dim TempInt As Integer
    Dim t As String
    Dim Y As Long
    Dim X As Long
    
    If General_File_Exist(SaveAs, vbNormal) = True Then
        If MsgBox("¿Desea sobrescribir " & SaveAs & "?", vbCritical + vbYesNo) = vbNo Then
            Exit Sub
        Else
            Kill SaveAs
        End If
    End If
    
    'Change mouse icon
    frmMain.MousePointer = 11
    t = SaveAs
    If General_File_Exist(left(SaveAs, Len(SaveAs) - 4) & ".inf", vbNormal) = True Then
        Kill left(SaveAs, Len(SaveAs) - 4) & ".inf"
    End If
    
    'Open .map file
    FreeFileMap = FreeFile
    Open SaveAs For Binary As FreeFileMap
    Seek FreeFileMap, 1
    
    
    SaveAs = left(SaveAs, Len(SaveAs) - 4)
    SaveAs = SaveAs & ".inf"
    'Open .inf file
    freefileinf = FreeFile
    Open SaveAs For Binary As freefileinf
    Seek freefileinf, 1

    Put FreeFileMap, , TempInt
    Put FreeFileMap, , MiCabecera
    
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    Put FreeFileMap, , TempInt
    
    'inf Header
    Put freefileinf, , TempInt
    Put freefileinf, , TempInt
    Put freefileinf, , TempInt
    Put freefileinf, , TempInt
    Put freefileinf, , TempInt
    
    'Write .map file
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            
            '.map file
            
            ' Bloqueos
            Put FreeFileMap, , MapData(X, Y).Blocked
            
            ' Capas
            For LoopC = 1 To 4
                'If LoopC = 2 Then Call FixCoasts(MapData(X, Y).Graphic(LoopC).grh_index, X, Y)
                Put FreeFileMap, , MapData(X, Y).Graphic(LoopC).index
            Next LoopC
            
            ' Triggers
            Put FreeFileMap, , MapData(X, Y).Trigger
            Put FreeFileMap, , TempInt
            
            '.inf file
            'Tile exit
            Put freefileinf, , MapData(X, Y).TileExit.Map
            Put freefileinf, , MapData(X, Y).TileExit.X
            Put freefileinf, , MapData(X, Y).TileExit.Y
            
            'NPC
            Put freefileinf, , MapData(X, Y).NPCIndex
            
            'Object
            Put freefileinf, , MapData(X, Y).OBJInfo.objindex
            Put freefileinf, , MapData(X, Y).OBJInfo.Amount
            
            'Empty place holders for future expansion
            Put freefileinf, , TempInt
            Put freefileinf, , TempInt
            
        Next X
    Next Y
    
    'Close .map file
    Close FreeFileMap
    'Close .inf file
    Close freefileinf
    FreeFileMap = FreeFile
    Open t & "2" For Binary Access Write As FreeFileMap
    Put FreeFileMap, , MapData
    Close FreeFileMap
    Call Pestañas(SaveAs)
    
    'write .dat file
    SaveAs = left(SaveAs, Len(SaveAs) - 4) & ".dat"
    MapInfo_Guardar SaveAs
    
    'Change mouse icon
    frmMain.MousePointer = 0
    MapInfo.Changed = 0
    
    Exit Sub
ErrorSave:
    MsgBox "Error " & Err.Number & " - " & Err.Description
End Sub
Public Sub AbrirMapaComun(ByVal file As String)

    Dim f As Integer
    Dim P As Long
    Dim size As Long
    Dim Size1 As Long
    Dim Size2 As Long
    Dim j As Long
    Dim data(279999) As Byte
    Dim nSpot As Integer
    Dim nPart As Integer
    Dim Data2() As Byte
    Dim SPOT() As MAPSPOTBUFFER
    Dim PART() As MAPPARTBUFFER


    Dim MapB() As tnwMapBlock
    Dim TempInt As Integer
    Dim FI As Integer
    Dim FileInfo As Boolean
    Dim Size3 As Long
    Dim DECS() As MAPDECSBUFFER
    Dim nDecs As Integer
    ReDim MapB(1 To 100, 1 To 100)
    Dim PATH As String
    On Error GoTo errs

    LightDestroyAll
    Particle_Group_Remove_All
    Map_ResetMontañita
    

    If UCase$(Right$(file, 4)) = "TEMP" Then
        MapaTemporal = True
        file = left$(file, Len(file) - 4)
    Else
        MapaTemporal = False
    End If
    
    If FileExist(left$(file, Len(file) - 3) & "inf", vbNormal) Then
        FileInfo = True
        FI = FreeFile
        Open left$(file, Len(file) - 3) & "inf" For Binary Access Read Lock Read As #FI
        Seek #FI, 1
        Get FI, , TempInt
        Get FI, , TempInt
        Get FI, , TempInt
        Get FI, , TempInt
        Get FI, , TempInt
    
    End If
    f = FreeFile
    If MapaTemporal Then

        Open file & "temp" For Binary Access Read Lock Read As #f
    Else
        Open file For Binary Access Read Lock Read As #f
    End If
    Get #f, , data
    Get #f, , nSpot
    Get #f, , nPart
    Get #f, , nDecs

    If nSpot > 0 Then
        ReDim SPOT(1 To nSpot)
        Size1 = 28 * nSpot
    End If
    If nPart > 0 Then
        ReDim PART(1 To nPart)
        Size2 = (3 * nPart)
    End If
    If nDecs > 0 Then
        ReDim DECS(1 To nDecs)
        Size3 = (LenB(DECS(1)) * nDecs)
    End If

    size = Size1 + Size2 + Size3
    If size > 0 Then
        ReDim Data2(0 To size - 1)
        Get #f, , Data2
        If Size1 > 0 Then
            CopyMemory SPOT(1), Data2(0), Size1
        End If
        If Size2 > 0 Then
            If Size1 > 0 Then
                CopyMemory PART(1), Data2(Size1), Size2
            Else
                CopyMemory PART(1), Data2(0), Size2
            End If
        End If
        If Size3 > 0 Then
            CopyMemory DECS(1), Data2(Size1 + Size2), Size3
        End If

    End If
    Close #f

    CopyMemory MapB(1, 1), data(0), 280000

    If nSpot > 0 Then
        For P = 1 To nSpot
            SPOTLIGHTS_CREAR SPOT(P).SPOT.SPOT_TIPO, SPOT(P).SPOT.SPOT_COLOR_BASE, SPOT(P).SPOT.SPOT_COLOR_EXTRA _
                , SPOT(P).SPOT.INTENSITY, 1, SPOT(P).SPOT.Grafico, SPOT(P).posi.X, SPOT(P).posi.Y, 0, SPOT(P).SPOT.Color, SPOT(P).SPOT.COLOR_EXTRA, SPOT(P).SPOT.EXTRA_GRAFICO
        Next P
    End If
    If nPart > 0 Then
        For P = 1 To nPart
            If PART(P).PARTI > 0 Then General_Particle_Create PART(P).PARTI, PART(P).posi.X, PART(P).posi.Y
            MapData(PART(P).posi.X, PART(P).posi.Y).parti_index = PART(P).PARTI
        Next P
    End If
    If nDecs > 0 Then
        For P = 1 To nDecs
            MapData(DECS(P).posi.X, DECS(P).posi.Y).DecorI = DECS(P).DecorI
            MapData(DECS(P).posi.X, DECS(P).posi.Y).DecorGrh.index = DecorData(DECS(P).DecorI).DecorGrh(1)
            MapData(DECS(P).posi.X, DECS(P).posi.Y).DecorInfo.EstadoDefault = DECS(P).EstadoDefault
        Next P
    End If

    Dim ByFlags As Byte

    For j = 1 To 100
        For P = 1 To 100

        
            MapData(P, j).Graphic(1).index = MapB(P, j).Layer(1)
            MapData(P, j).Graphic(2).index = MapB(P, j).Layer(2)
            MapData(P, j).Graphic(3).index = MapB(P, j).Layer(3)
            MapData(P, j).Graphic(4).index = MapB(P, j).Layer(4)
            MapData(P, j).Graphic(5).index = MapB(P, j).Layer(5)
        
            MapData(P, j).InteriorVal = MapB(P, j).InteriorVal
            MapData(P, j).IndexB(1) = MapB(P, j).IndexB(1) + 1
            MapData(P, j).IndexB(2) = MapB(P, j).IndexB(2) + 1
            MapData(P, j).IndexB(3) = MapB(P, j).IndexB(3) + 1
            MapData(P, j).IndexB(4) = MapB(P, j).IndexB(4) + 1
            MapData(P, j).IndexB(5) = MapB(P, j).IndexB(5) + 1
            MapData(P, j).TipoTerreno = MapB(P, j).TipoTerreno
            If MapData(P, j).IndexB(1) = 2 Then
                MapData(P, j).Graphic(1).index = MapData(P, j).Graphic(1).index + NumRealIndex
            End If
            If MapData(P, j).IndexB(2) = 2 Then
                MapData(P, j).Graphic(2).index = MapData(P, j).Graphic(2).index + NumRealIndex
            End If
            If MapData(P, j).IndexB(3) = 2 Then
                MapData(P, j).Graphic(3).index = MapData(P, j).Graphic(3).index + NumRealIndex
            End If
            If MapData(P, j).IndexB(4) = 2 Then
                MapData(P, j).Graphic(4).index = MapData(P, j).Graphic(4).index + NumRealIndex
            End If
            If MapData(P, j).IndexB(5) = 2 Then
                MapData(P, j).Graphic(5).index = MapData(P, j).Graphic(5).index + NumRealIndex
            End If
            If MapB(P, j).Trigger And 256 Then
                MapData(P, j).Blocked = 1
                MapData(P, j).Trigger = MapB(P, j).Trigger Xor 256
            Else
                MapData(P, j).Trigger = MapB(P, j).Trigger
            End If
        
            If FileInfo Then
                LeerFileInfo FI, P, j
            End If
        
        


        Next P
    Next j
    RecalcularLuces MapB

    If FileInfo Then
        Close FI
        

        
        bRefreshRadar = True ' Radar
                
        'Set changed flag
        MapInfo.Changed = 0
        
        ' Vacia el Deshacer
        modEdicion.Deshacer_Clear

    End If

    Exit Sub
errs:
    MsgBox "ERROR ABRIR MAPA COMUN : " & Err.Description




End Sub
Public Sub LeerFileInfo(ByVal FI As Integer, ByVal P As Integer, ByVal j As Integer)
    Dim ByFlags As Byte

    Get FI, , ByFlags
                
    If ByFlags And 1 Then
        Get FI, , MapData(P, j).TileExit.Map
        Get FI, , MapData(P, j).TileExit.X
        Get FI, , MapData(P, j).TileExit.Y
    End If
        
    If ByFlags And 2 Then
        'Get and make NPC
        Get FI, , MapData(P, j).NPCIndex
        Get FI, , MapData(P, j).NpcInfo.Heading
        Get FI, , MapData(P, j).NpcInfo.Nivel
        Get FI, , MapData(P, j).NpcInfo.Respawn
        If MapData(P, j).NpcInfo.Respawn = 1 Then
            Get FI, , MapData(P, j).NpcInfo.RespawnSamePos
            Get FI, , MapData(P, j).NpcInfo.RespawnTime
        End If
                    
        If MapData(P, j).NPCIndex < 0 Then
            MapData(P, j).NPCIndex = 0
            MapData(P, j).NpcInfo.Nivel = 0
            MapData(P, j).NpcInfo.Heading = 0
            MapData(P, j).NpcInfo.Respawn = 0
            MapData(P, j).NpcInfo.RespawnSamePos = 0
            MapData(P, j).NpcInfo.RespawnTime = 0
                        
                        
        Else
            Call MakeChar(NextOpenChar(), NpcData(MapData(P, j).NPCIndex).Body, NpcData(MapData(P, j).NPCIndex).Head, MapData(P, j).NpcInfo.Heading, CInt(P), CInt(j))
        End If
    End If
        
    If ByFlags And 4 Then
        'Get and make Object
        Get FI, , MapData(P, j).OBJInfo.objindex
        Get FI, , MapData(P, j).OBJInfo.Amount
        If MapData(P, j).OBJInfo.objindex > 0 Then
            MapData(P, j).ObjGrh.index = ObjData(MapData(P, j).OBJInfo.objindex).grh_index
        End If
    End If
            
            
        
    If ByFlags And 8 Then
        Get FI, , MapData(P, j).DecorInfo.Clave
        If MapData(P, j).DecorInfo.Clave > 0 Then MapData(P, j).DecorInfo.TipoClave = DecorKeys(MapData(P, j).DecorInfo.Clave).Tipo_Clave
    End If

End Sub





Public Sub MapaV3_Guardar(Mapa As String)
    '*************************************************
    'Author: Loopzer
    'Last modified: 22/11/07
    '*************************************************
    'copy&paste RLZ
    On Error GoTo ErrorSave
    Dim FreeFileMap As Long
    
    If General_File_Exist(Mapa, vbNormal) = True Then
        If MsgBox("¿Desea sobrescribir " & Mapa & "?", vbCritical + vbYesNo) = vbNo Then
            Exit Sub
        Else
            Kill Mapa
        End If
    End If
    
    frmMain.MousePointer = 11
    
    FreeFileMap = FreeFile
    Open Mapa For Binary Access Write As FreeFileMap
    Put FreeFileMap, , MapData
    Close FreeFileMap
    Call Pestañas(Mapa)
    
    
    Mapa = left(Mapa, Len(Mapa) - 4) & ".dat"
    MapInfo_Guardar Mapa
    
    'Change mouse icon
    frmMain.MousePointer = 0
    MapInfo.Changed = 0
    
    Exit Sub
ErrorSave:
    MsgBox "Error " & Err.Number & " - " & Err.Description
End Sub




' *****************************************************************************
' MAPINFO *********************************************************************
' *****************************************************************************

''
' Guardar Informacion del Mapa (.dat)
'
' @param Archivo Especifica el Path del archivo .DAT

Public Sub MapInfo_Guardar(ByVal Archivo As String)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 28/05/06
    '*************************************************

    If LenB(MapTitulo) = 0 Then
        MapTitulo = NameMap_Save
    End If

    Call WriteVar(Archivo, MapTitulo, "Name", MapInfo.Name)
    Call WriteVar(Archivo, MapTitulo, "MusicNum", MapInfo.Music)
    Call WriteVar(Archivo, MapTitulo, "MagiaSinefecto", Val(MapInfo.MagiaSinEfecto))
    Call WriteVar(Archivo, MapTitulo, "InviSinEfecto", Val(MapInfo.InviSinEfecto))
    Call WriteVar(Archivo, MapTitulo, "ResuSinEfecto", Val(MapInfo.ResuSinEfecto))
    Call WriteVar(Archivo, MapTitulo, "NoEncriptarMP", Val(MapInfo.NoEncriptarMP))

    Call WriteVar(Archivo, MapTitulo, "Terreno", MapInfo.Terreno)
    Call WriteVar(Archivo, MapTitulo, "Zona", MapInfo.Zona)
    Call WriteVar(Archivo, MapTitulo, "Restringir", MapInfo.Restringir)
    Call WriteVar(Archivo, MapTitulo, "BackUp", Str(MapInfo.BackUp))

    If MapInfo.PK Then
        Call WriteVar(Archivo, MapTitulo, "Pk", "0")
    Else
        Call WriteVar(Archivo, MapTitulo, "Pk", "1")
    End If
End Sub

''
' Abrir Informacion del Mapa (.dat)
'
' @param Archivo Especifica el Path del archivo .DAT

Public Sub MapInfo_Cargar(ByVal Archivo As String)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 02/06/06
    '*************************************************

    On Error Resume Next
    Dim Leer As New clsIniReader
    Dim LoopC As Integer
    Dim PATH As String
    MapTitulo = Empty
    Leer.Initialize Archivo

    For LoopC = Len(Archivo) To 1 Step -1
        If mid(Archivo, LoopC, 1) = "\" Then
            PATH = left(Archivo, LoopC)
            Exit For
        End If
    Next
    Archivo = Right(Archivo, Len(Archivo) - (Len(PATH)))
    MapTitulo = UCase(left(Archivo, Len(Archivo) - 4))

    MapInfo.Name = Leer.GetValue(MapTitulo, "Name")
    MapInfo.Music = Leer.GetValue(MapTitulo, "MusicNum")
    MapInfo.MagiaSinEfecto = Val(Leer.GetValue(MapTitulo, "MagiaSinEfecto"))
    MapInfo.InviSinEfecto = Val(Leer.GetValue(MapTitulo, "InviSinEfecto"))
    MapInfo.ResuSinEfecto = Val(Leer.GetValue(MapTitulo, "ResuSinEfecto"))
    MapInfo.NoEncriptarMP = Val(Leer.GetValue(MapTitulo, "NoEncriptarMP"))
    
    If Val(Leer.GetValue(MapTitulo, "Pk")) = 0 Then
        MapInfo.PK = True
    Else
        MapInfo.PK = False
    End If
    
    MapInfo.Terreno = Leer.GetValue(MapTitulo, "Terreno")
    MapInfo.Zona = Leer.GetValue(MapTitulo, "Zona")
    MapInfo.Restringir = Leer.GetValue(MapTitulo, "Restringir")
    MapInfo.BackUp = Val(Leer.GetValue(MapTitulo, "BACKUP"))
    
    Call MapInfo_Actualizar
    
End Sub

''
' Actualiza el formulario de MapInfo
'

Public Sub MapInfo_Actualizar()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 02/06/06
    '*************************************************

    On Error Resume Next
    ' Mostrar en Formularios
    frmMapInfo.txtMapNombre.Text = MapInfo.Name
    frmMapInfo.txtMapMusica.Text = MapInfo.Music
    frmMapInfo.txtMapTerreno.Text = MapInfo.Terreno
    frmMapInfo.txtMapZona.Text = MapInfo.Zona
    frmMapInfo.txtMapRestringir.Text = MapInfo.Restringir
    frmMapInfo.chkMapBackup.value = MapInfo.BackUp
    frmMapInfo.chkMapMagiaSinEfecto.value = MapInfo.MagiaSinEfecto
    frmMapInfo.chkMapInviSinEfecto.value = MapInfo.InviSinEfecto
    frmMapInfo.chkMapResuSinEfecto.value = MapInfo.ResuSinEfecto
    frmMapInfo.chkMapNoEncriptarMP.value = MapInfo.NoEncriptarMP
    frmMapInfo.chkMapPK.value = IIf(MapInfo.PK = True, 1, 0)
    frmMapInfo.txtMapVersion = MapInfo.MapVersion



End Sub

''
' Calcula la orden de Pestañas
'
' @param Map Especifica path del mapa

Public Sub Pestañas(ByVal Map As String)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 28/05/06
    '*************************************************
    On Error Resume Next
    Dim LoopC As Integer

    For LoopC = Len(Map) To 1 Step -1
        If mid(Map, LoopC, 1) = "\" Then
            PATH_Save = left(Map, LoopC)
            Exit For
        End If
    Next
    Map = Right(Map, Len(Map) - (Len(PATH_Save)))
    For LoopC = Len(left(Map, Len(Map) - 4)) To 1 Step -1
        If IsNumeric(mid(left(Map, Len(Map) - 4), LoopC, 1)) = False Then
            NumMap_Save = Right(left(Map, Len(Map) - 4), Len(left(Map, Len(Map) - 4)) - LoopC)
            NameMap_Save = left(Map, LoopC)
            Exit For
        End If
    Next
    For LoopC = (NumMap_Save - 4) To (NumMap_Save + 8)
        If General_File_Exist(PATH_Save & NameMap_Save & LoopC & ".map", vbArchive) = True Then
            frmMain.MapPest(LoopC - NumMap_Save + 4).Visible = True
            frmMain.MapPest(LoopC - NumMap_Save + 4).Enabled = True
            frmMain.MapPest(LoopC - NumMap_Save + 4).Caption = NameMap_Save & LoopC
        Else
            frmMain.MapPest(LoopC - NumMap_Save + 4).Visible = False
        End If
    Next
End Sub


Public Function dameNombreMapa(ByVal file As String) As String

    Dim i As Long
    Dim t As Long

    t = Len(file)

    For i = 0 To t - 1
    
        If mid$(file, t - i, 1) = "\" Then Exit For

    Next i

    file = mid$(file, t - i + 1)

    t = Len(file)

    For i = 0 To t - 1

        If mid$(file, t - i, 1) = "." Then Exit For

    Next i

    file = left$(file, t - (i + 1))

    dameNombreMapa = file
End Function
Public Sub Guardar_Nuevo_Mapa(ByVal PATH As String)
    Dim f As Integer
    Dim X As Long
    Dim Y As Long
    Dim NI As Integer
    Dim INn As Integer
    Dim IL() As tcardinal
    Dim lx As Integer
    Dim ly As Integer
    Dim SaltosFila(1 To 100) As Integer
    Dim SPL() As tcardinal
    Dim PRL() As tcardinal
    Dim DEC() As tcardinal
    Dim nSPL As Integer
    Dim nPRL As Integer
    Dim nDecs As Integer
    Dim TB As Byte
    Dim MapB() As tnwMapBlock
    Dim data(279999) As Byte
    Dim cSpots() As MAPSPOTBUFFER
    Dim cParts() As MAPPARTBUFFER
    Dim cDecs() As MAPDECSBUFFER

    Dim Usados(0 To 10000) As Byte

    ReDim MapB(1 To 100, 1 To 100)
    Dim TempInt As Integer

    On Error GoTo errx
    f = FreeFile


    If MapaTemporal Then
        If UCase$(Right$(PATH, 4)) = "TEMP" Then
            Open PATH For Binary Access Write Lock Write As #f
        Else
            Open PATH & "temp" For Binary Access Write Lock Write As #f
        End If
    Else
        Open PATH For Binary Access Write Lock Write As #f

    End If
    For Y = 1 To 100
        For X = 1 To 100

            If MapData(X, Y).SPOTLIGHT.index > 0 Then
                nSPL = nSPL + 1
                ReDim Preserve SPL(1 To nSPL)
                SPL(nSPL).X = X
                SPL(nSPL).Y = Y
            End If
            If MapData(X, Y).particle_group > 0 Then
                nPRL = nPRL + 1
                ReDim Preserve PRL(1 To nPRL)
                PRL(nPRL).X = X
                PRL(nPRL).Y = Y
            End If
            If MapData(X, Y).DecorI > 0 Then
                nDecs = nDecs + 1
                ReDim Preserve DEC(1 To nDecs)
                DEC(nDecs).X = X
                DEC(nDecs).Y = Y
            End If

            MapB(X, Y).Layer(1) = TempFromReal(MapData(X, Y).Graphic(1).index)
            MapB(X, Y).Layer(2) = TempFromReal(MapData(X, Y).Graphic(2).index)
            MapB(X, Y).Layer(3) = TempFromReal(MapData(X, Y).Graphic(3).index)
            MapB(X, Y).Layer(4) = TempFromReal(MapData(X, Y).Graphic(4).index)
            MapB(X, Y).Layer(5) = TempFromReal(MapData(X, Y).Graphic(5).index)
        
            If MapaTemporal Then
                Usados(MapData(X, Y).Graphic(1).index) = 1
                Usados(MapData(X, Y).Graphic(2).index) = 1
                Usados(MapData(X, Y).Graphic(3).index) = 1
                Usados(MapData(X, Y).Graphic(4).index) = 1
                Usados(MapData(X, Y).Graphic(5).index) = 1
            End If
        
            '/////En un mapa no grafico esto lo que hace es avisar si el indice es un indice temporal
            'o efectivamente esta indexado.
            'En un mapa grafico efectivamente es el indice dentro del grafico.
            If MapData(X, Y).IndexB(1) > 0 Then MapB(X, Y).IndexB(1) = MapData(X, Y).IndexB(1) - 1
            If MapData(X, Y).IndexB(2) > 1 Then
                MapB(X, Y).IndexB(2) = MapData(X, Y).IndexB(2) - 1
            End If
            If MapData(X, Y).IndexB(3) > 0 Then MapB(X, Y).IndexB(3) = MapData(X, Y).IndexB(3) - 1
            If MapData(X, Y).IndexB(4) > 0 Then MapB(X, Y).IndexB(4) = MapData(X, Y).IndexB(4) - 1
            If MapData(X, Y).IndexB(5) > 0 Then MapB(X, Y).IndexB(5) = MapData(X, Y).IndexB(5) - 1
            '///////////////////////////////////////////////////////////////////////////////

            MapB(X, Y).TipoTerreno = MapData(X, Y).TipoTerreno
        
            MapB(X, Y).Trigger = MergeTriggerBlock(MapData(X, Y).Trigger, MapData(X, Y).Blocked)
            MapB(X, Y).InteriorVal = MapData(X, Y).InteriorVal
            If MapData(X, Y).InteriorVal > 0 Then
                NI = NI + 1
                lx = NI
                INn = NI
                ReDim Preserve IL(1 To NI)
                IL(NI).X = X
                IL(NI).Y = Y
                If Y <> ly Then
                    If ly > 0 Then SaltosFila(ly) = NI
                    ly = Y
                End If
            Else
                If Y <> ly Then
                    lx = 0
                End If
                INn = 0
            End If
            MapB(X, Y).InteriorNum = INn
            MapB(X, Y).UltimoInteriorX = lx
        
        
            If MapData(X, Y).Luz >= 202 And MapData(X, Y).Luz <= 217 Then
                MapB(X, Y).Luces(0) = CByte(125 + MapData(X, Y).LV(0))
                MapB(X, Y).Luces(1) = CByte(125 + MapData(X, Y).LV(1))
                MapB(X, Y).Luces(2) = CByte(125 + MapData(X, Y).LV(2))
                MapB(X, Y).Luces(3) = CByte(125 + MapData(X, Y).LV(3))
            Else
                MapB(X, Y).Luces(0) = CByte(MapData(X, Y).LV(0))
                MapB(X, Y).Luces(1) = CByte(MapData(X, Y).LV(1))
                MapB(X, Y).Luces(2) = CByte(MapData(X, Y).LV(2))
                MapB(X, Y).Luces(3) = CByte(MapData(X, Y).LV(3))
            End If
            'LUCES
        Next X
    Next Y

    CopyMemory data(0), MapB(1, 1), 280000

    Dim P As Long
    Dim Size1 As Long
    Dim Size2 As Long
    Dim Size3 As Long
    Dim size As Long
    Dim size4 As Long

    Put #f, , data()
    Put #f, , nSPL
    Put #f, , nPRL
    Put #f, , nDecs


    If nSPL > 0 Then
        Dim Prev As Long
        ReDim cSpots(1 To nSPL)
        Size1 = LenB(cSpots(1)) * nSPL
        For P = 1 To nSPL
            cSpots(P).posi.X = SPL(P).X
            cSpots(P).posi.Y = SPL(P).Y
            cSpots(P).SPOT = MapData(SPL(P).X, SPL(P).Y).SPOTLIGHT
        Next P
    End If

    If nPRL > 0 Then

        ReDim cParts(1 To nPRL)
        Size2 = LenB(cParts(1)) * nPRL
        For P = 1 To nPRL
            cParts(P).posi.X = PRL(P).X
            cParts(P).posi.Y = PRL(P).Y
            cParts(P).PARTI = MapData(PRL(P).X, PRL(P).Y).parti_index
        Next P
    End If
    If nDecs > 0 Then
        ReDim cDecs(1 To nDecs)
        Size3 = LenB(cDecs(1)) * nDecs
        For P = 1 To nDecs
            cDecs(P).posi.X = DEC(P).X
            cDecs(P).posi.Y = DEC(P).Y
            cDecs(P).DecorI = MapData(DEC(P).X, DEC(P).Y).DecorI
            cDecs(P).EstadoDefault = MapData(DEC(P).X, DEC(P).Y).DecorInfo.EstadoDefault
        Next P
    End If
    'If nLy5 > 0 Then
    '    ReDim cLY5(1 To nLy5)
    '    size4 = LenB(cLY5(1)) * nLy5
    '    For P = 1 To nLy5
    '        cLY5(P).posi.X = LY5(P).X
    '        cLY5(P).posi.Y = LY5(P).Y
    '        cLY5(P).Graf = MapData(LY5(P).X, LY5(P).Y).Graphic(5).index
    '        cLY5(P).index = MapData(LY5(P).X, LY5(P).Y).IndexB(5)
    '    Next P
    'End If

    Dim BATA() As Byte
    size = Size1 + Size2 + Size3 ' + size4
    If size > 0 Then
        ReDim BATA(0 To size - 1) As Byte
    End If
    If nSPL > 0 Then
        CopyMemory BATA(0), cSpots(1), Size1
    End If
    If nPRL > 0 Then
        If nSPL > 0 Then
            CopyMemory BATA(Size1), cParts(1), Size2
        Else
            CopyMemory BATA(0), cParts(1), Size2
        End If
    End If
    If nDecs > 0 Then
        CopyMemory BATA(Size2 + Size1), cDecs(1), Size3
    End If
    'If nLy5 > 0 Then
    '    CopyMemory BATA(Size1 + Size2 + Size3), cLY5(1), size4
    'End If

    If size > 0 Then
        Put #f, , BATA
    End If
    Close #f
    f = FreeFile


    Open left$(PATH, Len(PATH) - 3) & "int" For Binary Access Write Lock Write As #f
    Put #f, , NI
    Put #f, , SaltosFila
    Put #f, , IL
    Close #f
    f = FreeFile


    Dim ByFlags As Byte

    If frmMain.chkGuardarInf.Checked Then

        Open left$(PATH, Len(PATH) - 3) & "inf" For Binary Access Write Lock Write As #f
        Seek #f, 1
        Put #f, , TempInt
        Put #f, , TempInt
        Put #f, , TempInt
        Put #f, , TempInt
        Put #f, , TempInt
        For Y = 1 To 100
            For X = 1 To 100
                ByFlags = 0
                
                If MapData(X, Y).TileExit.Map Then ByFlags = ByFlags Or 1
                If MapData(X, Y).NPCIndex Then ByFlags = ByFlags Or 2

                If MapData(X, Y).OBJInfo.objindex Then ByFlags = ByFlags Or 4
                If MapData(X, Y).DecorI > 0 Then ByFlags = ByFlags Or 8
                
                Put f, , ByFlags
                
                If MapData(X, Y).TileExit.Map Then
                    Put f, , MapData(X, Y).TileExit.Map
                    Put f, , MapData(X, Y).TileExit.X
                    Put f, , MapData(X, Y).TileExit.Y
                End If
                
                If MapData(X, Y).NPCIndex Then
                
                    Put f, , CInt(MapData(X, Y).NPCIndex)
                    Put f, , MapData(X, Y).NpcInfo.Heading
                    Put f, , MapData(X, Y).NpcInfo.Nivel
                    Put f, , MapData(X, Y).NpcInfo.Respawn
                    If MapData(X, Y).NpcInfo.Respawn = 1 Then
                        Put f, , MapData(X, Y).NpcInfo.RespawnSamePos
                        Put f, , MapData(X, Y).NpcInfo.RespawnTime
                    End If
                End If
                
                If MapData(X, Y).OBJInfo.objindex Then
                    Put f, , MapData(X, Y).OBJInfo.objindex
                    Put f, , MapData(X, Y).OBJInfo.Amount
                End If
                
                If MapData(X, Y).DecorI > 0 Then
                    Put f, , MapData(X, Y).DecorInfo.Clave
                End If

            Next X
        Next Y
        
        Close #f

    End If
    Dim EstaticUsada(0 To 10000) As Byte
    Dim UsoSt As Integer
    If MapaTemporal Then
        Dim n As Long
        Dim H As String
        Dim G As Integer
        H = left$(PATH, Len(PATH) - 7) & "TempIndex"
    
        For n = NumRealIndex To 10000
            If Usados(n) = 1 Then
                G = G + 1
                WriteVar H, CStr(G), "Index", CStr(TempFromReal(n))
                WriteVar H, CStr(G), "OverWriteGrafico", CStr(NewIndexData(n).OverWriteGrafico)
                If TempIndex(TempFromReal(n)).temp = 1 Then
                    WriteVar H, CStr(G), "Estatica", CStr(NewIndexData(n).Estatic - NumRealEstatic)
                Else
                    WriteVar H, CStr(G), "Estatica", CStr(NewIndexData(n).Estatic)
                End If
                WriteVar H, CStr(G), "Temp", CStr(TempIndex(TempFromReal(n)).temp)
                WriteVar H, CStr(G), "Replace", CStr(TempIndex(TempFromReal(n)).Replace)
                If TempIndex(TempFromReal(n)).temp = 1 Then
                    If EstaticUsada(TempIndex(TempFromReal(n)).Estatic) = 0 Then
                        EstaticUsada(TempIndex(TempFromReal(n)).Estatic) = 1
                        UsoSt = UsoSt + 1
                    End If
                End If
            End If
        Next n
        WriteVar H, "INIT", "NumTI", CStr(G)
        WriteVar H, "INIT", "NumTE", CStr(UsoSt)
        G = 0
        If UsoSt > 0 Then
            For n = 1 To 10000
                If EstaticUsada(n) = 1 Then
                    G = G + 1
                    WriteVar H, "e" & G, "Index", CStr(n)
                    WriteVar H, "e" & G, "Left", CStr(TempEstatic(n).L)
                    WriteVar H, "e" & G, "Top", CStr(TempEstatic(n).t)
                    WriteVar H, "e" & G, "Width", CStr(TempEstatic(n).W)
                    WriteVar H, "e" & G, "Height", CStr(TempEstatic(n).H)
                    WriteVar H, "e" & G, "Left", CStr(TempEstatic(n).Replace)
                End If
            Next n
        End If
    End If

    frmMain.Caption = "Guardado " & PATH
    MapInfo.Changed = 0
    Exit Sub
errx:
    Debug.Print X & "_" & Y & "_" & Err.Description

    Close #f
    MsgBox "Error al guardar el mapa en posicion " & X & "_" & Y

End Sub

Public Function MergeTriggerBlock(ByVal Trigger As Integer, ByVal Blocked As Byte) As Integer

    If Blocked = 1 Then
        MergeTriggerBlock = Trigger Or 256
    Else
        MergeTriggerBlock = Trigger
    End If
End Function

Function HayAguaGrh(ByVal X As Integer, ByVal Y As Integer) As Byte
    
      
End Function
Function HayLavaGrh(ByVal X As Integer, ByVal Y As Integer) As Byte
    If MapData(X, Y).Graphic(1).index > 0 Then
        If NewIndexData(MapData(X, Y).Graphic(1).index).OverWriteGrafico = 9510 Then
      
            HayLavaGrh = 1

        End If
    End If
End Function
Public Sub InsertarSurface(ByVal X As Integer, ByVal Y As Integer)
    Dim P As Long
    Dim t As Long
    If X >= 1 And Y >= 1 And X <= 99 And Y <= 99 Then
        For t = 0 To (Val(frmMain.SizeC.List(frmMain.SizeC.ListIndex)) / 32) - 1
            If SobreY + t > 100 Then Exit For
            For P = 0 To (Val(frmMain.SizeC.List(frmMain.SizeC.ListIndex)) / 32) - 1
                If SobreX + P > 100 Then Exit For
            
                MapData(SobreX + P, SobreY + t).IndexB(frmMain.LayerC.ListIndex + 1) = ((t) * 16) + (P + 1)
                MapData(SobreX + P, SobreY + t).SizeC = frmMain.SizeC.ListIndex
                MapData(SobreX + P, SobreY + t).Graphic(frmMain.LayerC.ListIndex + 1).index = frmMain.txtNumSurface
            Next P

        Next t
    End If

End Sub
Public Sub BorrarSurface(ByVal X As Integer, ByVal Y As Integer)

    Dim P As Long
    Dim t As Long
    Dim iX As Integer
    Dim iY As Integer
    Dim tY As Integer
    Dim tX As Integer
    Dim s As Integer
    If X >= 1 And Y >= 1 And X <= 99 And Y <= 99 Then

        If MapData(X, Y).IndexB(frmMain.LayerC.ListIndex + 1) = 1 Then
            iX = X
            iY = Y
        Else
            tY = Int((MapData(X, Y).IndexB(frmMain.LayerC.ListIndex + 1) - 1) / 16)
            iY = Y - tY
            tX = (MapData(X, Y).IndexB(frmMain.LayerC.ListIndex + 1) - 1) Mod 16
            iX = X - tX
        End If
        s = MapData(X, Y).SizeC
        For t = iY To iY + (Val(frmMain.SizeC.List(s)) / 32)
            If t > 100 Then Exit For
            For P = iX To iX + (Val(frmMain.SizeC.List(s)) / 32)
                If P > 100 Then Exit For
                MapData(P, t).IndexB(frmMain.LayerC.ListIndex + 1) = 0
                MapData(P, t).Graphic(frmMain.LayerC.ListIndex + 1).index = 0
                MapData(P, t).SizeC = 0
            Next P

        Next t
    End If

End Sub
Public Sub AbrirMapaGrafico(ByVal file As String)

    Dim f As Integer
    Dim P As Long
    Dim size As Long
    Dim Size1 As Long
    Dim Size2 As Long
    Dim j As Long
    Dim data(279999) As Byte
    Dim nSpot As Integer
    Dim nPart As Integer
    Dim Data2() As Byte
    Dim SPOT() As MAPSPOTBUFFER
    Dim PART() As MAPPARTBUFFER
    Dim MapB() As tnwMapBlock
    Dim TempInt As Integer
    Dim FI As Integer
    Dim FileInfo As Boolean
    Dim Size3 As Long
    Dim DECS() As MAPDECSBUFFER
    Dim nDecs As Integer
    ReDim MapB(1 To 100, 1 To 100)
    Dim nLy5 As Integer
    Dim LY5() As MAPLAYER5BUFFER
    Dim size4 As Long

    Debug.Print LenB(MapB(1, 1))

    MapaTemporal = False
 

    If FileExist(left$(file, Len(file) - 3) & "inf", vbNormal) Then
        FileInfo = True
        FI = FreeFile
        Open left$(file, Len(file) - 3) & "inf" For Binary Access Read Lock Read As #FI
        Seek #FI, 1
        Get FI, , TempInt
        Get FI, , TempInt
        Get FI, , TempInt
        Get FI, , TempInt
        Get FI, , TempInt
    
    End If
    f = FreeFile
    Open file For Binary Access Read Lock Read As #f

    Get #f, , data
    Get #f, , nSpot
    Get #f, , nPart
    Get #f, , nDecs

    If nSpot > 0 Then
        ReDim SPOT(1 To nSpot)
        Size1 = 28 * nSpot
    End If
    If nPart > 0 Then
        ReDim PART(1 To nPart)
        Size2 = (3 * nPart)
    End If
    If nDecs > 0 Then
        ReDim DECS(1 To nDecs)
        Size3 = (4 * nDecs)
    End If

    
    size = Size1 + Size2 + Size3
    If size > 0 Then
        ReDim Data2(0 To size - 1)
        Get #f, , Data2
        If Size1 > 0 Then
            CopyMemory SPOT(1), Data2(0), Size1
        End If
        If Size2 > 0 Then
            If Size1 > 0 Then
                CopyMemory PART(1), Data2(Size1), Size2
            Else
                CopyMemory PART(1), Data2(0), Size2
            End If
        End If
        If nDecs > 0 Then
            CopyMemory DECS(1), Data2(Size1 + Size2), Size3
        End If

    End If
    Close #f

    CopyMemory MapB(1, 1), data(0), 280000

    If nSpot > 0 Then
        For P = 1 To nSpot
            SPOTLIGHTS_CREAR SPOT(P).SPOT.SPOT_TIPO, SPOT(P).SPOT.SPOT_COLOR_BASE, SPOT(P).SPOT.SPOT_COLOR_EXTRA _
                , SPOT(P).SPOT.INTENSITY, 1, SPOT(P).SPOT.Grafico, SPOT(P).posi.X, SPOT(P).posi.Y, 0, SPOT(P).SPOT.Color, SPOT(P).SPOT.COLOR_EXTRA, SPOT(P).SPOT.EXTRA_GRAFICO
        Next P
    End If
    If nPart > 0 Then
        For P = 1 To nPart
            If PART(P).PARTI > 0 Then General_Particle_Create PART(P).PARTI, PART(P).posi.X, PART(P).posi.Y
        Next P
    End If
    If nDecs > 0 Then
        For P = 1 To nDecs
            MapData(DECS(P).posi.X, DECS(P).posi.Y).DecorI = DECS(P).DecorI
        Next P
    End If

    Dim ByFlags As Byte

    For j = 1 To 100
        For P = 1 To 100
            MapData(P, j).Graphic(1).index = MapB(P, j).Layer(1)
            MapData(P, j).Graphic(2).index = MapB(P, j).Layer(2)
            MapData(P, j).Graphic(3).index = MapB(P, j).Layer(3)
            MapData(P, j).Graphic(4).index = MapB(P, j).Layer(4)
            MapData(P, j).Graphic(5).index = MapB(P, j).Layer(5)
        
            MapData(P, j).IndexB(1) = MapB(P, j).IndexB(1) + 1
            MapData(P, j).IndexB(2) = MapB(P, j).IndexB(2) + 1
            MapData(P, j).IndexB(3) = MapB(P, j).IndexB(3) + 1
            MapData(P, j).IndexB(4) = MapB(P, j).IndexB(4) + 1
            MapData(P, j).IndexB(5) = MapB(P, j).IndexB(5) + 1
        
            MapData(P, j).InteriorVal = MapB(P, j).InteriorVal
            MapData(P, j).TipoTerreno = MapB(P, j).TipoTerreno
            If MapB(P, j).Trigger And 256 Then
                MapData(P, j).Blocked = 1
                MapData(P, j).Trigger = MapB(P, j).Trigger Xor 256
            Else
                MapData(P, j).Trigger = MapB(P, j).Trigger
            End If
        
            If FileInfo Then
                LeerFileInfo FI, P, j
            End If
            
        
        
        
        


        Next P
    Next j
    RecalcularLuces MapB

    If FileInfo Then
        Close FI
        

        
        bRefreshRadar = True ' Radar
                
        'Set changed flag
        MapInfo.Changed = 0
        
        ' Vacia el Deshacer
        modEdicion.Deshacer_Clear

    End If




End Sub
Sub RecalcularLuces(ByRef MapBu() As tnwMapBlock)
    On Error GoTo errl
    Dim X As Long
    Dim Y As Long

    Dim Luz As Byte
    Dim pcLuz As Byte

    Luz = 14
    pcLuz = Luz * 9

    For X = 1 To 100
        For Y = 1 To 100

            With MapBu(X, Y)
                If .Luces(0) >= INITIAL_INVERT_MAPLIGHT Then
                    MapData(X, Y).LV(0) = .Luces(0) - INITIAL_INVERT_MAPLIGHT
                Else
                    MapData(X, Y).LV(0) = .Luces(0)
                End If
                If .Luces(1) >= INITIAL_INVERT_MAPLIGHT Then
                    MapData(X, Y).LV(1) = .Luces(1) - INITIAL_INVERT_MAPLIGHT
                Else
                    MapData(X, Y).LV(1) = .Luces(1)
                End If
                If .Luces(2) >= INITIAL_INVERT_MAPLIGHT Then
                    MapData(X, Y).LV(2) = .Luces(2) - INITIAL_INVERT_MAPLIGHT
                Else
                    MapData(X, Y).LV(2) = .Luces(2)
                End If
                If .Luces(3) >= INITIAL_INVERT_MAPLIGHT Then
                    MapData(X, Y).LV(3) = .Luces(3) - INITIAL_INVERT_MAPLIGHT
                Else
                    MapData(X, Y).LV(3) = .Luces(3)
                End If
            
                If .Luces(0) >= INITIAL_INVERT_MAPLIGHT Then
                    MapData(X, Y).Luz = 203
                ElseIf .Luces(0) <> 0 Or .Luces(1) <> 0 Or .Luces(2) <> 0 Or .Luces(3) <> 0 Then
                    MapData(X, Y).Luz = 20
                End If
            
                If .Luces(0) > 0 Then
                    If .Luces(0) < 9 Then
                        MapData(X, Y).light_value(0) = ambient_light(pcLuz + .Luces(0) + 1)
                    ElseIf .Luces(0) < INITIAL_INVERT_MAPLIGHT Then
                        MapData(X, Y).light_value(0) = extra_light(.Luces(0) - 8)
                    ElseIf .Luces(0) > INITIAL_INVERT_MAPLIGHT Then
                        If .Luces(0) < INITIAL_INVERT_MAPLIGHT + 9 Then
                            MapData(X, Y).light_value(0) = ambient_light(pcLuz + (.Luces(0) - INITIAL_INVERT_MAPLIGHT + 1))
                        Else
                            MapData(X, Y).light_value(0) = extra_light(.Luces(0) - INITIAL_INVERT_MAPLIGHT - 8)
                        End If
                    End If
                Else
                    MapData(X, Y).light_value(0) = 0
                End If
                If .Luces(1) > 0 Then

                    If .Luces(1) < 9 Then
                        MapData(X, Y).light_value(1) = ambient_light(pcLuz + .Luces(1) + 1)
                    ElseIf .Luces(1) < INITIAL_INVERT_MAPLIGHT Then
                        MapData(X, Y).light_value(1) = extra_light(.Luces(1) - 8)
                    ElseIf .Luces(1) > INITIAL_INVERT_MAPLIGHT Then
                        If .Luces(1) < INITIAL_INVERT_MAPLIGHT + 9 Then
                            MapData(X, Y).light_value(1) = ambient_light(pcLuz + (.Luces(1) - INITIAL_INVERT_MAPLIGHT + 1))
                        Else
                            MapData(X, Y).light_value(1) = extra_light(.Luces(1) - INITIAL_INVERT_MAPLIGHT - 8)
                        End If
                    End If
                Else
                    MapData(X, Y).light_value(1) = 0
                End If
                If .Luces(2) > 0 Then
                    If .Luces(2) < 9 Then
                        MapData(X, Y).light_value(2) = ambient_light(pcLuz + .Luces(2) + 1)
                    ElseIf .Luces(2) < INITIAL_INVERT_MAPLIGHT Then
                        MapData(X, Y).light_value(2) = extra_light(.Luces(2) - 8)
                    ElseIf .Luces(2) > INITIAL_INVERT_MAPLIGHT Then
                        If .Luces(2) < INITIAL_INVERT_MAPLIGHT + 9 Then
                            MapData(X, Y).light_value(2) = ambient_light(pcLuz + (.Luces(2) - INITIAL_INVERT_MAPLIGHT + 1))
                        Else
                            MapData(X, Y).light_value(2) = extra_light(.Luces(2) - INITIAL_INVERT_MAPLIGHT - 8)
                        End If
                    End If
                Else
                    MapData(X, Y).light_value(2) = 0
                End If
                If .Luces(3) > 0 Then
                    If .Luces(3) < 9 Then
                        MapData(X, Y).light_value(3) = ambient_light(pcLuz + .Luces(3) + 1)
                    ElseIf .Luces(3) < INITIAL_INVERT_MAPLIGHT Then
                        MapData(X, Y).light_value(3) = extra_light(.Luces(3) - 8)
                    ElseIf .Luces(3) > INITIAL_INVERT_MAPLIGHT Then
                        If .Luces(3) < INITIAL_INVERT_MAPLIGHT + 9 Then
                            MapData(X, Y).light_value(3) = ambient_light(pcLuz + (.Luces(3) - INITIAL_INVERT_MAPLIGHT + 1))
                        Else
                            MapData(X, Y).light_value(3) = extra_light(.Luces(3) - INITIAL_INVERT_MAPLIGHT - 8)
                        End If
                    End If
                Else
                    MapData(X, Y).light_value(3) = 0
                End If
            End With
        Next Y
    Next X

    base_light = ambient_light((pcLuz) + 1)
    Exit Sub
errl:
    Debug.Print X & "_" & Y
    MsgBox "ERROR RECALCULARLUCES: " & Err.Description
End Sub
Public Sub BorrarSobrante(ByVal X As Integer, ByVal Y As Integer)
    Dim P As Long
    Dim j As Long

    For j = Y To Y + (Val(frmMain.SizeC.List(frmMain.SizeC.ListIndex)) / 32) - 1
        If j > 100 Then Exit For
        For P = X To X + (Val(frmMain.SizeC.List(frmMain.SizeC.ListIndex)) / 32) - 1
            If P > 100 Then Exit For
            MapData(P, j).Graphic(frmMain.LayerC.ListIndex + 1).index = 0
            MapData(P, j).IndexB(frmMain.LayerC.ListIndex + 1) = 0
            MapData(P, j).SizeC = 0
        Next P
    Next j

End Sub
Public Sub EditarIndice(ByVal X As Integer, ByVal Y As Integer)
    If Val(frmMain.txtNumIndice) > 0 And Val(frmMain.txtNumSurface) > 0 And Val(frmMain.txtNumIndice) <= 256 Then
        MapData(X, Y).IndexB(frmMain.LayerC.ListIndex + 1) = Val(frmMain.txtNumIndice)
        MapData(X, Y).Graphic(frmMain.LayerC.ListIndex + 1).index = Val(frmMain.txtNumSurface)
        MapData(X, Y).SizeC = 4

    End If
End Sub
Public Function SetTipoTerreno() As Byte
    Dim i As Long
    Dim TB As Byte
    For i = 0 To frmMain.Check1.UBound
        
        If frmMain.Check1(i).value Then
            TB = TB Or (2 ^ i)
        End If
    Next i
    
    SetTipoTerreno = TB
End Function
Public Sub GetTipoTerreno(ByVal TB As Byte)
    Dim i As Long
    For i = 0 To frmMain.Check1.UBound
        If TB And (2 ^ i) Then
            frmMain.Check1(i).value = 1
        Else
            frmMain.Check1(i).value = 0
        End If
    Next i
End Sub
