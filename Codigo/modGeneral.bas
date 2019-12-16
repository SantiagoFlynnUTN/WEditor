Attribute VB_Name = "modGeneral"
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
' modGeneral
'
' @remarks Funciones Generales
' @author unkwown
' @version 0.4.11
' @date 20061015

Option Explicit
Public timer_elapsed_time As Single
Public NuevoOk As Boolean
Public SobreIndex As Integer
Public NuevoTex As Integer
Public Type typDevMODE
    dmDeviceName       As String * 32
    dmSpecVersion      As Integer
    dmDriverVersion    As Integer
    dmSize             As Integer
    dmDriverExtra      As Integer
    dmFields           As Long
    dmOrientation      As Integer
    dmPaperSize        As Integer
    dmPaperLength      As Integer
    dmPaperWidth       As Integer
    dmScale            As Integer
    dmCopies           As Integer
    dmDefaultSource    As Integer
    dmPrintQuality     As Integer
    dmColor            As Integer
    dmDuplex           As Integer
    dmYResolution      As Integer
    dmTTOption         As Integer
    dmCollate          As Integer
    dmFormName         As String * 32
    dmUnusedPadding    As Integer
    dmBitsPerPel       As Integer
    dmPelsWidth        As Long
    dmPelsHeight       As Long
    dmDisplayFlags     As Long
    dmDisplayFrequency As Long
End Type
Public Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lptypDevMode As Any) As Boolean
Public Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lptypDevMode As Any, ByVal dwFlags As Long) As Long

Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

Public Const CCDEVICENAME = 32
Public Const CCFORMNAME = 32
Public Const DM_BITSPERPEL = &H40000
Public Const DM_PELSWIDTH = &H80000
Public Const DM_DISPLAYFREQUENCY = &H400000
Public Const DM_PELSHEIGHT = &H100000
Public Const CDS_UPDATEREGISTRY = &H1
Public Const CDS_TEST = &H4
Public Const DISP_CHANGE_SUCCESSFUL = 0
Public Const DISP_CHANGE_RESTART = 1

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function GetActiveWindow Lib "user32" () As Long

''
' Checks if this is the active (foreground) application or not.
'
' @return   True if any of the app's windows are the foreground window, false otherwise.

Public Function IsAppActive() As Boolean
    '***************************************************
    'Author: Juan Mart�n Sotuyo Dodero (maraxus)
    'Last Modify Date: 03/03/2007
    'Checks if this is the active application or not
    '***************************************************
    IsAppActive = (GetActiveWindow <> 0)
End Function

''
' Realiza acciones de desplasamiento segun las teclas que hallamos precionado
'

Public Sub CheckKeys()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 28/05/06
    '*************************************************
    If HotKeysAllow = False Then Exit Sub

    
    Static timer As Long
    If GetTickCount - timer > 30 Then
        timer = GetTickCount
    Else
        Exit Sub
    End If
    
    If GetKeyState(vbKeyUp) < 0 Then
        If UserPos.Y < 1 Then Exit Sub ' 10
        If LegalPos(UserPos.X, UserPos.Y - 1) And WalkMode = True Then
            If dLastWalk + 50 > GetTickCount Then Exit Sub
            UserPos.Y = UserPos.Y - 1
            MoveCharbyPos UserCharIndex, UserPos.X, UserPos.Y
            dLastWalk = GetTickCount
        ElseIf WalkMode = False Then
            UserPos.Y = UserPos.Y - 1
        End If
        bRefreshRadar = True ' Radar
        frmMain.SetFocus
        Exit Sub
    End If

    If GetKeyState(vbKeyRight) < 0 Then
        If UserPos.X > 100 Then Exit Sub ' 89
        If LegalPos(UserPos.X + 1, UserPos.Y) And WalkMode = True Then
            If dLastWalk + 50 > GetTickCount Then Exit Sub
            UserPos.X = UserPos.X + 1
            MoveCharbyPos UserCharIndex, UserPos.X, UserPos.Y
            dLastWalk = GetTickCount
        ElseIf WalkMode = False Then
            UserPos.X = UserPos.X + 1
        End If
        bRefreshRadar = True ' Radar
        frmMain.SetFocus
        Exit Sub
    End If

    If GetKeyState(vbKeyDown) < 0 Then
        If UserPos.Y > 100 Then Exit Sub ' 92
        If LegalPos(UserPos.X, UserPos.Y + 1) And WalkMode = True Then
            If dLastWalk + 50 > GetTickCount Then Exit Sub
            UserPos.Y = UserPos.Y + 1
            MoveCharbyPos UserCharIndex, UserPos.X, UserPos.Y
            dLastWalk = GetTickCount
        ElseIf WalkMode = False Then
            UserPos.Y = UserPos.Y + 1
        End If
        bRefreshRadar = True ' Radar
        frmMain.SetFocus
        Exit Sub
    End If

    If GetKeyState(vbKeyLeft) < 0 Then
        If UserPos.X < 1 Then Exit Sub ' 12
        If LegalPos(UserPos.X - 1, UserPos.Y) And WalkMode = True Then
            If dLastWalk + 50 > GetTickCount Then Exit Sub
            UserPos.X = UserPos.X - 1
            MoveCharbyPos UserCharIndex, UserPos.X, UserPos.Y
            dLastWalk = GetTickCount
        ElseIf WalkMode = False Then
            UserPos.X = UserPos.X - 1
        End If
        bRefreshRadar = True ' Radar
        frmMain.SetFocus
        Exit Sub
    End If
End Sub
Public Function general_field_read(ByVal Pos As Integer, Text As String, SepASCII As Integer) As String
    '*************************************************
    'Author: Unkwown
    'Last modified: 20/05/06
    '*************************************************
    Dim i As Integer
    Dim lastPos As Integer
    Dim CurChar As String * 1
    Dim FieldNum As Integer
    Dim Seperator As String

    Seperator = Chr(SepASCII)
    lastPos = 0
    FieldNum = 0

    For i = 1 To Len(Text)
        CurChar = mid(Text, i, 1)
        If CurChar = Seperator Then
            FieldNum = FieldNum + 1
            If FieldNum = Pos Then
                general_field_read = mid(Text, lastPos + 1, (InStr(lastPos + 1, Text, Seperator, vbTextCompare) - 1) - (lastPos))
                Exit Function
            End If
            lastPos = i
        End If
    Next i
    FieldNum = FieldNum + 1

    If FieldNum = Pos Then
        general_field_read = mid(Text, lastPos + 1)
    End If

End Function


''
' Completa y corrije un path
'
' @param Path Especifica el path con el que se trabajara
' @return   Nos devuelve el path completado

Private Function autoCompletaPath(ByVal PATH As String) As String
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 22/05/06
    '*************************************************
    PATH = Replace(PATH, "/", "\")
    If left(PATH, 1) = "\" Then
        ' agrego app.path & path
        PATH = App.PATH & PATH
    End If
    If Right(PATH, 1) <> "\" Then
        ' me aseguro que el final sea con "\"
        PATH = PATH & "\"
    End If
    autoCompletaPath = PATH
End Function

''
' Carga la configuracion del WorldEditor de WorldEditor.ini
'

Private Sub CargarMapIni()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 24/11/08
    '*************************************************
    On Error GoTo Fallo
    Dim tStr As String

    Dim Leer As New clsIniReader


    inipath = App.PATH & "\"
    DirGraficos = App.PATH & "\Resources\Graficos\"
    DirGraficosMap = App.PATH & "\Resources\GRH_MAP\"
    DirIndex = App.PATH & "\Resources\Init\"
    DirMidi = App.PATH & "\Resources\Midi\"
    frmMusica.fleMusicas.PATH = DirMidi
    DirDats = App.PATH & "\Resources\Dats\"
    dirwavs = App.PATH & "\Resources\WaV\"

    If General_File_Exist(inipath & "WorldEditor.ini", vbArchive) = False Then
        frmMain.mnuGuardarUltimaConfig.Checked = True

    
        MaxGrhs = 15000
        UserPos.X = 50
        UserPos.Y = 50
        PantallaX = 19
        PantallaY = 22
        MsgBox "Falta el archivo 'WorldEditor.ini' de configuraci�n.", vbInformation
        Exit Sub
    End If


    Call Leer.Initialize(inipath & "WorldEditor.ini")
    DirGraficos = App.PATH & "\Resources\Graficos\"
    DirGraficosMap = App.PATH & "\Resources\GRH_MAP\"
    DirIndex = App.PATH & "\Resources\Init\"
    DirMidi = App.PATH & "\Resources\Midi\"
    frmMusica.fleMusicas.PATH = DirMidi
    DirDats = App.PATH & "\Resources\Dats\"
    dirwavs = App.PATH & "\Resources\WaV\"
    ' Obj de Translado
    Cfg_TrOBJ = Val(Leer.GetValue("CONFIGURACION", "ObjTranslado"))

    frmMain.mnuAutoCapturarTranslados.Checked = IIf(Val(Leer.GetValue("CONFIGURACION", "AutoCapturarTrans")) = 1, True, False)
    frmMain.mnuAutoCapturarSuperficie.Checked = IIf(Val(Leer.GetValue("CONFIGURACION", "AutoCapturarSup")) = 1, True, False)

    frmMain.mnuUtilizarDeshacer.Checked = IIf(Val(Leer.GetValue("CONFIGURACION", "UtilizarDeshacer")) = 1, True, False)

    ' Guardar Ultima Configuracion
    frmMain.mnuGuardarUltimaConfig.Checked = IIf(Val(Leer.GetValue("CONFIGURACION", "GuardarConfig")) = 1, True, False)

    ' Index
    MaxGrhs = Val(general_var_get(inipath & "WorldEditor.ini", "INDEX", "MaxGrhs"))
    If MaxGrhs < 1 Then MaxGrhs = 15000

    'Reciente
    frmMain.Dialog.InitDir = Leer.GetValue("PATH", "UltimoMapa")

    If General_File_Exist(DirGraficos, vbDirectory) = False Then
        MsgBox "El directorio de Graficos es incorrecto", vbCritical + vbOKOnly
        End
    End If

    If General_File_Exist(DirMidi, vbDirectory) = False Then
        MsgBox "El directorio de MIDI es incorrecto", vbCritical + vbOKOnly
        End
    End If

    frmMusica.fleMusicas.PATH = DirMidi

    If General_File_Exist(DirIndex, vbDirectory) = False Then
        MsgBox "El directorio de Index es incorrecto", vbCritical + vbOKOnly
        End
    End If

    If General_File_Exist(DirDats, vbDirectory) = False Then
        MsgBox "El directorio de Dats es incorrecto", vbCritical + vbOKOnly
        End
    End If

    If General_File_Exist(dirwavs, vbDirectory) = False Then
        MsgBox "El directorio de WAV es incorrecto", vbCritical + vbOKOnly
        End
    End If

    tStr = Leer.GetValue("MOSTRAR", "LastPos") ' x-y
    UserPos.X = Val(general_field_read(1, tStr, Asc("-")))
    UserPos.Y = Val(general_field_read(2, tStr, Asc("-")))
    If UserPos.X < XMinMapSize Or UserPos.X > XMaxMapSize Then
        UserPos.X = 50
    End If
    If UserPos.Y < YMinMapSize Or UserPos.Y > YMaxMapSize Then
        UserPos.Y = 50
    End If

    ' Menu Mostrar
    frmMain.mnuVerAutomatico.Checked = Val(Leer.GetValue("MOSTRAR", "ControlAutomatico"))
    frmMain.mnuVerCapa2.Checked = IIf(Val(Leer.GetValue("MOSTRAR", "Capa2")) = 1, True, False)
    frmMain.mnuVerCapa3.Checked = IIf(Val(Leer.GetValue("MOSTRAR", "Capa3")) = 1, True, False)
    frmMain.mnuVerCapa4.Checked = IIf(Val(Leer.GetValue("MOSTRAR", "Capa4")) = 1, True, False)
    frmMain.mnuVerTranslados.Checked = IIf(Val(Leer.GetValue("MOSTRAR", "Translados")) = 1, True, False)
    frmMain.mnuVerObjetos.Checked = IIf(Val(Leer.GetValue("MOSTRAR", "Objetos")) = 1, True, False)
    frmMain.mnuVerNPCs.Checked = IIf(Val(Leer.GetValue("MOSTRAR", "NPCs")) = 1, True, False)
    frmMain.mnuVerTriggers.Checked = IIf(Val(Leer.GetValue("MOSTRAR", "Triggers")) = 1, True, False)
    frmMain.mnuVerGrilla.Checked = IIf(Val(Leer.GetValue("MOSTRAR", "Grilla")) = 1, True, False) ' Grilla
    frmMain.mVerDecors.Checked = IIf(Val(Leer.GetValue("MOSTRAR", "Decors")) = 1, True, False)
    VerGrilla = frmMain.mnuVerGrilla.Checked
    frmMain.mnuVerBloqueos.Checked = IIf(Val(Leer.GetValue("MOSTRAR", "Bloqueos")) = 1, True, False)
    frmMain.cVerTriggers.value = frmMain.mnuVerTriggers.Checked
    frmMain.cVerBloqueos.value = frmMain.mnuVerBloqueos.Checked

    ' Tama�o de visualizacion
    PantallaX = Val(Leer.GetValue("MOSTRAR", "PantallaX"))
    PantallaY = Val(Leer.GetValue("MOSTRAR", "PantallaY"))
    If PantallaX > 23 Or PantallaX <= 2 Then PantallaX = 23
    If PantallaY > 32 Or PantallaY <= 2 Then PantallaY = 32

    ' [GS] 02/10/06
    ' Tama�o de visualizacion en el cliente
    ClienteHeight = Val(Leer.GetValue("MOSTRAR", "ClienteHeight"))
    ClienteWidth = Val(Leer.GetValue("MOSTRAR", "ClienteWidth"))
    If ClienteHeight <= 0 Then ClienteHeight = 18
    If ClienteWidth <= 0 Then ClienteWidth = 24

    Exit Sub
Fallo:
    MsgBox "ERROR " & Err.Number & " en WorldEditor.ini" & vbCrLf & Err.Description, vbCritical
    Resume Next
End Sub

Public Function TomarBPP() As Integer
    Dim ModoDeVideo As typDevMODE
    Call EnumDisplaySettings(0, -1, ModoDeVideo)
    TomarBPP = CInt(ModoDeVideo.dmBitsPerPel)
End Function
Public Sub CambioDeVideo()
    '*************************************************
    'Author: Loopzer
    '*************************************************
    Exit Sub
    Dim ModoDeVideo As typDevMODE
    Dim R As Long
    Call EnumDisplaySettings(0, -1, ModoDeVideo)
    If ModoDeVideo.dmPelsWidth < 1024 Or ModoDeVideo.dmPelsHeight < 768 Then
        Select Case MsgBox("La aplicacion necesita una resolucion minima de 1024 X 768 ,�Acepta el Cambio de resolucion?", vbInformation + vbOKCancel, "World Editor")
            Case vbOK
                ModoDeVideo.dmPelsWidth = 1024
                ModoDeVideo.dmPelsHeight = 768
                ModoDeVideo.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
                R = ChangeDisplaySettings(ModoDeVideo, CDS_TEST)
                If R <> 0 Then
                    MsgBox "Error al cambiar la resolucion, La aplicacion se cerrara."
                    End
                End If
            Case vbCancel
                End
        End Select
    End If
End Sub

Public Sub Main()
    '*************************************************
    'Author: Unkwown
    'Last modified: 25/11/08 - GS
    '*************************************************
    'On Error Resume Next
    Dim OffsetCounterX As Integer
    Dim OffsetCounterY As Integer
    Dim Chkflag As Integer
    Static polygon_one_down As Single
    Dim waterHeight As Integer
    waterHeight = 4
    
    
    SeleccionnGrh.index = SeleccionIndex
    SeleccionadoArrayColor(0) = vbRed
    SeleccionadoArrayColor(1) = vbRed
    SeleccionadoArrayColor(2) = vbRed
    SeleccionadoArrayColor(3) = vbRed
    
    If App.PrevInstance = True Then End
    'CambioDeVideo
    modTileEngine.Iniciar_IndicesNewMap
    'If General_File_Exist(DirIndex & "AO.dat", vbArchive) Then

    'End If
    
    Set DXPool = New clsTextureManager

    
    Call CargarMapIni

    Call IniciarCabecera(MiCabecera)
    DoEvents

    setup_ambient


    
    frmCargando.verX = "v" & App.Major & "." & App.Minor & "." & App.Revision
    frmCargando.Show
    frmCargando.SetFocus
    frmCargando.X.Caption = "Iniciando DirectSound..."
    DoEvents
    Call Audio.Initialize(frmMain.hWnd, dirwavs, DirMidi)

    Audio.MusicActivated = True
    Audio.SoundActivated = True
    
    frmCargando.X.Caption = "Cargando Indice de Superficies..."
    'modIndices.CargarIndicesSuperficie

    modExtras.LoadTexWe
    modGameIni.Load_Graphic_Header
    Load_Png_Array
    frmCargando.X.Caption = "Indexando Cargado de Imagenes..."
    DoEvents
        
    'frmMain.MainViewPic.width = PantallaX ^ 3
    'frmMain.MainViewPic.height = PantallaY ^ 3
        
    modDXEngine.DXEngine_Initialize frmMain.hWnd, frmMain.MainViewPic.hWnd, True
    
    modGrh.Animations_Initialize 0.03, 32
        
    Meteo.Initialize
    Meteo.Set_Time 15, 0
        
    If InitTileEngine(frmMain.hWnd, frmMain.MainViewPic.top, frmMain.MainViewPic.left, 32, 32, ClienteHeight, ClienteWidth, 9) Then    ' 30/05/2006
        'Display form handle, View window offset from 0,0 of display form, Tile Size, Display size in tiles, Screen buffer
        
        
        frmCargando.P1.Visible = True
        frmCargando.L(0).Visible = True
        frmCargando.X.Caption = "Cargando indexaciones..."
        modGameIni.LoadIndexData
        modGameIni.CargarEfectos
        DoEvents
        

        
        frmCargando.P3.Visible = True
        frmCargando.L(2).Visible = True
        frmCargando.X.Caption = "Cargando NPC's..."
        modIndices.CargarIndicesNPC
        DoEvents
        
        frmCargando.P4.Visible = True
        frmCargando.L(3).Visible = True
        frmCargando.X.Caption = "Cargando Objetos..."
        modIndices.CargarIndicesOBJ
        DoEvents
        
        frmCargando.X.Caption = "Cargando Decors..."
        modGameIni.Load_DecorData
        DoEvents
        
        frmCargando.X.Caption = "Cargando Decorskeys..."
        modGameIni.Load_DecoKeys
        DoEvents
        
        
        
        frmCargando.P5.Visible = True
        frmCargando.L(4).Visible = True
        frmCargando.X.Caption = "Cargando Triggers..."
        modIndices.CargarIndicesTriggers
        modGameIni.CargarTipoTerrenos
        DoEvents
        
        frmCargando.P6.Visible = True
        frmCargando.L(5).Visible = True
    End If
    
    frmCargando.SetFocus
    frmCargando.X.Caption = "Iniciando Ventana de Edici�n..."
    
    frmCargando.Hide
    frmMain.Show
    modMapIO.NuevoMapa
    
    prgRun = True
    cFPS = 0
    Chkflag = 0
    dTiempoGT = GetTickCount
    
    Dim alturaAgua As Integer
    Dim ma1_bajando As Byte
    alturaAgua = 6

    Do While prgRun
        If ma1_bajando = 0 Then
            ma(0) = ma(0) + timer_ticks_per_frame / 2
            If ma(0) >= alturaAgua Then
                ma(0) = alturaAgua
                ma1_bajando = 1
            End If
        Else
            ma(0) = ma(0) - timer_ticks_per_frame / 2
            If ma(0) <= -alturaAgua Then
                ma(0) = -alturaAgua
                ma1_bajando = 0
            End If
        End If
    
        ma(1) = ma(0)
    
        If ma1_bajando = 0 Then
            ma(1) = ma(1) + (alturaAgua / 2)
            If ma(1) >= alturaAgua Then ma(1) = alturaAgua - (ma(1) - alturaAgua)
        Else
            ma(1) = ma(1) - (alturaAgua / 2)
            If ma(1) <= -alturaAgua Then ma(1) = -alturaAgua + Abs(ma(1) + alturaAgua)
        End If
    
        ma(1) = -ma(1)
    
        If AddtoUserPos.X <> 0 Then
            OffsetCounterX = (OffsetCounterX - (8 * Sgn(AddtoUserPos.X)))
            If Abs(OffsetCounterX) >= Abs(TilePixelWidth * AddtoUserPos.X) Then
                OffsetCounterX = 0
                AddtoUserPos.X = 0
            End If
        ElseIf AddtoUserPos.Y <> 0 Then
            OffsetCounterY = OffsetCounterY - (8 * Sgn(AddtoUserPos.Y))
            If Abs(OffsetCounterY) >= Abs(TilePixelHeight * AddtoUserPos.Y) Then
                OffsetCounterY = 0
                AddtoUserPos.Y = 0
            End If
        End If

        If (GetTickCount - dTiempoGT) >= 1000 Then
            CaptionWorldEditor frmMain.Dialog.FileName, (MapInfo.Changed = 1)
            frmMain.Caption = cFPS
            cFPS = 1
            dTiempoGT = GetTickCount

        Else
            cFPS = cFPS + 1
        End If
        

        If Chkflag = 3 Then
            If frmMain.WindowState <> 1 Then Call CheckKeys
            If polygon_one_down = 1 Then
                polygonCount(0) = polygonCount(0) + (3 * (MEE * 0.0015))
                If polygonCount(0) >= waterHeight Then
                    polygonCount(0) = waterHeight
                    polygon_one_down = 0
                End If
            Else
                polygonCount(0) = polygonCount(0) - (3 * (MEE * 0.0015))
                If polygonCount(0) <= -waterHeight Then
                    polygonCount(0) = -waterHeight
                    polygon_one_down = 1
                End If
            End If
               
            PolygonX = Abs(polygonCount(0))
            POlyGonXMin = PolygonX * 0.25
               
            polygonCount(1) = polygonCount(0)
              
            If polygon_one_down = 0 Then
                polygonCount(1) = polygonCount(1) + (waterHeight * 0.25)
                If polygonCount(1) >= waterHeight Then polygonCount(1) = waterHeight - (polygonCount(1) - waterHeight)
            Else
                polygonCount(1) = polygonCount(1) - (waterHeight * 0.25)
                If polygonCount(1) <= -waterHeight Then polygonCount(1) = -waterHeight + Abs(polygonCount(1) + waterHeight)
            End If
              
            polygonCount(1) = -polygonCount(1)

            MEE = PasoTiempo(Movement_Engine_Counter)
            Chkflag = 0

            DXEngine_BeginRender
            If TIPOMAPAX = 0 Then
                
                Call RenderScreen(UserPos.X - AddtoUserPos.X, UserPos.Y - AddtoUserPos.Y, OffsetCounterX, OffsetCounterY)


            Else
                Call RenderNewMap(UserPos.X - AddtoUserPos.X, UserPos.Y - AddtoUserPos.Y, OffsetCounterX, OffsetCounterY)
            End If
            
            modDXEngine.SPOTLIGHTS_RENDER
            
            DXEngine_EndRender
            
            timer_elapsed_time = General_Get_Elapsed_Time()
            modGrh.AnimSpeedCalculate timer_elapsed_time


        End If
        
        Chkflag = Chkflag + 1
        
       
        

        
        If bRefreshRadar Then Call RefreshAllChars
        
        DoEvents

        'Sleep 1
    Loop
    
    If MapInfo.Changed = 1 Then
        If MsgBox(MSGMod, vbExclamation + vbYesNo) = vbYes Then
            modMapIO.GuardarMapa frmMain.Dialog.FileName
        End If
    End If
    
    DXEngine_Deinitialize
    Dim f
    
    For Each f In Forms
        Unload f
    Next
    
    End
End Sub

Public Function general_var_get(file As String, ByVal Main As String, Var As String) As String
    '*************************************************
    'Author: Unkwown
    'Last modified: 20/05/06
    '*************************************************

    Dim sSpaces As String ' This will hold the input that the program will retrieve
    Dim szReturn As String ' This will be the defaul value if the string is not found
    szReturn = vbNullString
    sSpaces = Space(5000) ' This tells the computer how long the longest string can be. If you want, you can change the number 75 to any number you wish
    GetPrivateProfileString Main, Var, szReturn, sSpaces, Len(sSpaces), file
    general_var_get = RTrim(sSpaces)
    general_var_get = left(general_var_get, Len(general_var_get) - 1)
End Function

Public Sub WriteVar(file As String, Main As String, Var As String, value As String)
    '*************************************************
    'Author: Unkwown
    'Last modified: 20/05/06
    '*************************************************
    writeprivateprofilestring Main, Var, value, file
End Sub

Public Sub ToggleWalkMode()
    '*************************************************
    'Author: Unkwown
    'Last modified: 28/05/06 - GS
    '*************************************************
    On Error GoTo fin:
    If WalkMode = False Then
        WalkMode = True
    Else
        frmMain.mnuModoCaminata.Checked = False
        WalkMode = False
    End If

    If WalkMode = False Then
        'Erase character
        Call EraseChar(UserCharIndex)
        MapData(UserPos.X, UserPos.Y).CHarIndex = 0
    Else
        'MakeCharacter
        If LegalPos(UserPos.X, UserPos.Y) Then
            Call MakeChar(NextOpenChar(), 11, 5, SOUTH, UserPos.X, UserPos.Y)
            UserCharIndex = MapData(UserPos.X, UserPos.Y).CHarIndex
            frmMain.mnuModoCaminata.Checked = True
        Else
            MsgBox "ERROR: Ubicacion ilegal."
            WalkMode = False
        End If
    End If
fin:
End Sub

Public Sub FixCoasts(ByVal GrhIndex As Integer, ByVal X As Integer, ByVal Y As Integer)
    '*************************************************
    'Author: Unkwown
    'Last modified: 20/05/06
    '*************************************************

    'If GrhIndex = 7284 Or GrhIndex = 7290 Or GrhIndex = 7291 Or GrhIndex = 7297 Or _
     GrhIndex = 7300 Or GrhIndex = 7301 Or GrhIndex = 7302 Or GrhIndex = 7303 Or _
     GrhIndex = 7304 Or GrhIndex = 7306 Or GrhIndex = 7308 Or GrhIndex = 7310 Or _
     GrhIndex = 7311 Or GrhIndex = 7313 Or GrhIndex = 7314 Or GrhIndex = 7315 Or _
     GrhIndex = 7316 Or GrhIndex = 7317 Or GrhIndex = 7319 Or GrhIndex = 7321 Or _
     GrhIndex = 7325 Or GrhIndex = 7326 Or GrhIndex = 7327 Or GrhIndex = 7328 Or GrhIndex = 7332 Or _
     GrhIndex = 7338 Or GrhIndex = 7339 Or GrhIndex = 7345 Or GrhIndex = 7348 Or _
     GrhIndex = 7349 Or GrhIndex = 7350 Or GrhIndex = 7351 Or GrhIndex = 7352 Or _
     GrhIndex = 7349 Or GrhIndex = 7350 Or GrhIndex = 7351 Or _
     GrhIndex = 7354 Or GrhIndex = 7357 Or GrhIndex = 7358 Or GrhIndex = 7360 Or _
     GrhIndex = 7362 Or GrhIndex = 7363 Or GrhIndex = 7365 Or GrhIndex = 7366 Or _
     GrhIndex = 7367 Or GrhIndex = 7368 Or GrhIndex = 7369 Or GrhIndex = 7371 Or _
     GrhIndex = 7373 Or GrhIndex = 7375 Or GrhIndex = 7376 Then MapData(X, Y).Graphic(2).grh_index = 0

End Sub

Public Function General_Random_Number(ByVal LowerBound As Variant, ByVal UpperBound As Variant) As Single
    '*************************************************
    'Author: Unkwown
    'Last modified: 20/05/06
    '*************************************************
    Randomize timer
    General_Random_Number = (UpperBound - LowerBound) * Rnd + LowerBound
End Function


''
' Actualiza todos los Chars en el mapa
'

Public Sub RefreshAllChars()
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 28/05/06
    '*************************************************
    On Error Resume Next
    Dim LoopC As Integer

    frmMain.ApuntadorRadar.Move UserPos.X - 12, UserPos.Y - 10
    frmMain.picRadar.Cls

    For LoopC = 1 To LastChar
        If CharList(LoopC).Active = 1 Then
            MapData(CharList(LoopC).Pos.X, CharList(LoopC).Pos.Y).CHarIndex = LoopC
            If CharList(LoopC).Heading <> 0 Then
                frmMain.picRadar.ForeColor = vbGreen
                frmMain.picRadar.Line (0 + CharList(LoopC).Pos.X, 0 + CharList(LoopC).Pos.Y)-(2 + CharList(LoopC).Pos.X, 0 + CharList(LoopC).Pos.Y)
                frmMain.picRadar.Line (0 + CharList(LoopC).Pos.X, 1 + CharList(LoopC).Pos.Y)-(2 + CharList(LoopC).Pos.X, 1 + CharList(LoopC).Pos.Y)
            End If
        End If
    Next LoopC

    bRefreshRadar = False
End Sub


''
' Actualiza el Caption del menu principal
'
' @param Trabajando Indica el path del mapa con el que se esta trabajando
' @param Editado Indica si el mapa esta editado

Public Sub CaptionWorldEditor(ByVal Trabajando As String, ByVal Editado As Boolean)
    '*************************************************
    'Author: ^[GS]^
    'Last modified: 20/05/06
    '*************************************************
    If Trabajando = vbNullString Then
        Trabajando = "Nuevo Mapa"
    End If



    frmMain.Caption = "WorldEditor v" & App.Major & "." & App.Minor & " Build " & App.Revision & " - [" & Trabajando & "]"
    If Editado = True Then
        frmMain.Caption = frmMain.Caption & " (modificado)"
    End If

    If MapaTemporal Then
        frmMain.Caption = frmMain.Caption & " [ Mapa Temporal ]"
    End If

End Sub



Public Function General_Get_Elapsed_Time() As Single
    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    'Gets the time that past since the last call
    '**************************************************************
    Dim start_time As Currency
    Static end_time As Currency
    Static timer_freq As Currency

    'Get the timer frequency
    If timer_freq = 0 Then
        QueryPerformanceFrequency timer_freq
    End If

    'Get current time
    QueryPerformanceCounter start_time
    
    'Calculate elapsed time
    General_Get_Elapsed_Time = (start_time - end_time) / timer_freq * 1000
    
    'Get next end time
    QueryPerformanceCounter end_time
End Function
Function FileExist(ByVal file As String, ByVal FileType As VbFileAttribute) As Boolean
    FileExist = (Dir$(file, FileType) <> "")
End Function