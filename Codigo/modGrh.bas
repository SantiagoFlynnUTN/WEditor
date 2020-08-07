Attribute VB_Name = "modGrh"
Private Const LoopAdEternum As Integer = 999
Public Type tnGrh
    index As Integer
    fC As Single

End Type
'Holds data about where a bmp can be found,
'How big it is and animation info
Public Type Grh_Data
    Active As Boolean
    texture_index As Long
    Src_X As Integer
    Src_Y As Integer
    src_width As Integer
    src_height As Integer
    
    frame_count As Integer
    frame_list(1 To 25) As Long
    frame_speed As Single
    MiniMap_color As Long
End Type

'Points to a Grh_Data and keeps animation info
Public Type Grh
    grh_index As Integer
    alpha_blend As Boolean
    Angle As Single
    frame_speed As Single
    frame_counter As Single
    Started As Boolean
    LoopTimes As Integer
End Type

'Grh Data Array
Public grh_list() As Grh_Data
Public grh_count As Long

Dim AnimBaseSpeed As Single
Public timer_ticks_per_frame As Single

Public base_tile_size As Integer

Public Sub Grh_Initialize(ByRef Grh As Grh, ByVal grh_index As Long, Optional ByVal alpha_blend As Boolean, Optional ByVal Angle As Single, Optional ByVal Started As Byte = 2, Optional ByVal LoopTimes As Integer = LoopAdEternum)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Sets up a grh. MUST be done before rendering
'*****************************************************************
    If grh_index <= 0 Then Exit Sub

    'Copy of parameters
    Grh.grh_index = grh_index
    Grh.alpha_blend = alpha_blend
    Grh.Angle = Angle
    Grh.LoopTimes = LoopTimes
    
    'Start it if it's a animated grh
    If Started = 2 Then
        If grh_list(Grh.grh_index).frame_count > 1 Then
            Grh.Started = True
        Else
            Grh.Started = False
        End If
    Else
        Grh.Started = Started
    End If
    
    'Set frame counters
    Grh.frame_counter = 1
    
    Grh.frame_speed = grh_list(Grh.grh_index).frame_speed
End Sub


Public Sub Grh_iRender(ByRef Grh As Grh, ByVal screen_x As Long, ByVal screen_Y As Long, ByRef rgb_list() As Long, Optional ByVal center As Boolean)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/28/2003
'
'**************************************************************
    Dim tile_width As Single
    Dim tile_height As Single
    Dim grh_index As Long


    If Grh.grh_index = 0 Then Exit Sub

    'Animation
    If Grh.Started Then
        Grh.frame_counter = Grh.frame_counter + (timer_ticks_per_frame * Grh.frame_speed / 1000)
        If Grh.frame_counter > grh_list(Grh.grh_index).frame_count Then
            If Grh.LoopTimes < 2 Then
                Grh.frame_counter = 1
                Grh.Started = False
            Else
                Grh.frame_counter = 1
                If Grh.LoopTimes <> LoopAdEternum Then
                    Grh.LoopTimes = Grh.LoopTimes - 1
                End If
            End If
        End If
    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    If Grh.frame_counter <= 0 Then Grh.frame_counter = 1
    grh_index = grh_list(Grh.grh_index).frame_list(Grh.frame_counter)
    
    If grh_index = 0 Then Exit Sub 'This is an error condition
    
    'Center Grh over X,Y pos
    If center Then
        tile_width = grh_list(grh_index).src_width / base_tile_size
        tile_height = grh_list(grh_index).src_height / base_tile_size
        If tile_width <> 1 Then
            screen_x = screen_x - Int(tile_width * base_tile_size / 2) + base_tile_size / 2
        End If
        If tile_height <> 1 Then
            screen_Y = screen_Y - Int(tile_height * base_tile_size) + base_tile_size
        End If
    End If
    
    'Draw it to device
    DXEngine_iTextureRender grh_list(grh_index).texture_index, _
        screen_x, screen_Y, _
        grh_list(grh_index).src_width, grh_list(grh_index).src_height, _
        rgb_list, _
        grh_list(grh_index).Src_X, grh_list(grh_index).Src_Y, _
        grh_list(grh_index).src_width, grh_list(grh_index).src_height, _
        Grh.alpha_blend, _
        Grh.Angle
End Sub
Public Sub Grh_Render(ByRef Grh As Grh, ByVal screen_x As Long, ByVal screen_Y As Long, ByRef rgb_list() As Long, Optional ByVal center As Boolean)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/28/2003
'
'**************************************************************
    Dim tile_width As Single
    Dim tile_height As Single
    Dim grh_index As Long


    If Grh.grh_index = 0 Then Exit Sub

    'Animation
    If Grh.Started Then
        Grh.frame_counter = Grh.frame_counter + (timer_ticks_per_frame * Grh.frame_speed / 1000)
        If Grh.frame_counter > grh_list(Grh.grh_index).frame_count Then
            If Grh.LoopTimes < 2 Then
                Grh.frame_counter = 1
                Grh.Started = False
            Else
                Grh.frame_counter = 1
                If Grh.LoopTimes <> LoopAdEternum Then
                    Grh.LoopTimes = Grh.LoopTimes - 1
                End If
            End If
        End If
    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    If Grh.frame_counter <= 0 Then Grh.frame_counter = 1
    grh_index = grh_list(Grh.grh_index).frame_list(Grh.frame_counter)
    
    If grh_index = 0 Then Exit Sub 'This is an error condition
    
    'Center Grh over X,Y pos
    If center Then
        tile_width = grh_list(grh_index).src_width / base_tile_size
        tile_height = grh_list(grh_index).src_height / base_tile_size
        If tile_width <> 1 Then
            screen_x = screen_x - Int(tile_width * base_tile_size / 2) + base_tile_size / 2
        End If
        If tile_height <> 1 Then
            screen_Y = screen_Y - Int(tile_height * base_tile_size) + base_tile_size
        End If
    End If
    
    'Draw it to device
    DXEngine_TextureRender grh_list(grh_index).texture_index, _
        screen_x, screen_Y, _
        grh_list(grh_index).src_width, grh_list(grh_index).src_height, _
        rgb_list, _
        grh_list(grh_index).Src_X, grh_list(grh_index).Src_Y, _
        grh_list(grh_index).src_width, grh_list(grh_index).src_height, _
        Grh.alpha_blend, _
        Grh.Angle
End Sub
Public Sub Grh_iRenderN(ByRef Grh As tnGrh, ByVal screen_x As Long, ByVal screen_Y As Long, ByRef rgb_list() As Long, Optional ByVal center As Boolean, Optional ByVal HalfCenter As Boolean)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/28/2003
'
'**************************************************************
    Dim tile_width As Single
    Dim tile_height As Single
    Dim grh_index As Long
    Dim jx As Integer
    Dim jy As Integer
    Dim jw As Integer
    Dim jh As Integer
    Dim jtw As Single
    Dim jth As Single
    Dim jg As Integer
    If NewIndexData(Grh.index).Dinamica > 0 Then
        With NewAnimationData(NewIndexData(Grh.index).Dinamica)
            Grh.fC = Grh.fC + (timer_ticks_per_frame * .NumFrames / .Velocidad)
            If Grh.fC > .NumFrames Then Grh.fC = (Grh.fC Mod .NumFrames) + 1
            If Grh.fC <= 0 Then Grh.fC = 1
            jx = .Indice(Grh.fC).X
            jy = .Indice(Grh.fC).Y
            jw = .Width
            jh = .Height
            jtw = .TileWidth
            jth = .TileHeight
            jg = (.Indice(Grh.fC).Grafico - .Indice(1).Grafico) + NewIndexData(Grh.index).OverWriteGrafico
        End With
    Else
        With EstaticData(NewIndexData(Grh.index).Estatic)
            jx = .L
            jy = .t
            jh = .H
            jw = .W
            jg = NewIndexData(Grh.index).OverWriteGrafico
            jtw = .tw
            jth = .th
        End With
    End If
    
    
    
    
    If center Then
        If jtw <> 1 Then
            screen_x = screen_x - Int(jtw * base_tile_size / 2) + base_tile_size / 2
        End If
        If jth <> 1 Then
            If HalfCenter Then
            screen_Y = screen_Y - Int(jth * base_tile_size / 2) + base_tile_size / 2
            Else
            screen_Y = screen_Y - Int(jth * base_tile_size) + base_tile_size
            
            End If
        End If
    End If
    
    'Draw it to device
    DXEngine_iTextureRender jg, _
        screen_x, screen_Y, _
        jw, jh, _
        rgb_list, _
        jx, jy, _
         jw, jh, _
        0, _
        0
End Sub
Public Sub Grh_RenderN(ByRef Grh As tnGrh, ByVal screen_x As Long, ByVal screen_Y As Long, ByRef rgb_list() As Long, Optional ByVal center As Boolean, Optional ByVal HalfCenter As Boolean)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/28/2003
'
'**************************************************************
    Dim tile_width As Single
    Dim tile_height As Single
    Dim grh_index As Long
    Dim jx As Integer
    Dim jy As Integer
    Dim jw As Integer
    Dim jh As Integer
    Dim jtw As Single
    Dim jth As Single
    Dim jg As Integer
    If NewIndexData(Grh.index).Dinamica > 0 Then
        With NewAnimationData(NewIndexData(Grh.index).Dinamica)
            Grh.fC = Grh.fC + (timer_ticks_per_frame * .NumFrames / .Velocidad)
            If Grh.fC > .NumFrames Then Grh.fC = (Grh.fC Mod .NumFrames) + 1
            If Grh.fC <= 0 Then Grh.fC = 1
            jx = .Indice(Grh.fC).X
            jy = .Indice(Grh.fC).Y
            jw = .Width
            jh = .Height
            jtw = .TileWidth
            jth = .TileHeight
            jg = (.Indice(Grh.fC).Grafico - .Indice(1).Grafico) + NewIndexData(Grh.index).OverWriteGrafico
        End With
    Else
        If NewIndexData(Grh.index).Estatic <= 0 Then Exit Sub
        With EstaticData(NewIndexData(Grh.index).Estatic)
            jx = .L
            jy = .t
            jh = .H
            jw = .W
            jg = NewIndexData(Grh.index).OverWriteGrafico
            jtw = .tw
            jth = .th
        End With
    End If
    
    
    
    
    If center Then
        If jtw <> 1 Then
            screen_x = screen_x - Int(jtw * base_tile_size / 2) + base_tile_size / 2
        End If
        'If jth <> 1 Then
            If HalfCenter Then
            screen_Y = screen_Y - Int(jth * base_tile_size / 2) + base_tile_size / 2
            Else
             screen_Y = screen_Y - Int(jth * base_tile_size) + base_tile_size
           
            End If
        'End If
    End If
    
    'Draw it to device
    DXEngine_TextureRender jg, _
        screen_x, screen_Y, _
        jw, jh, _
        rgb_list, _
        jx, jy, _
         jw, jh, _
        0, _
        0
End Sub
Public Sub Anim_iRender(ByRef Grh As tnGrh, ByVal screen_x As Long, ByVal screen_Y As Long, ByRef rgb_list() As Long, Optional ByVal center As Boolean, Optional ByVal Anim As Boolean, Optional ByVal Grafico As Integer)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/28/2003
'
'**************************************************************
    Dim tile_width As Single
    Dim tile_height As Single
    Dim grh_index As Long
    Dim jx As Integer
    Dim jy As Integer
    Dim jw As Integer
    Dim jh As Integer
    Dim jtw As Single
    Dim jth As Single
    Dim jg As Integer
    With NewAnimationData(Grh.index)
        If Anim Then Grh.fC = Grh.fC + (timer_ticks_per_frame * .NumFrames / .Velocidad)
        
        If Grh.fC > .NumFrames Then Grh.fC = (Grh.fC Mod .NumFrames) + 1
        If Grh.fC <= 0 Then Grh.fC = 1
        jx = .Indice(Grh.fC).X
        jy = .Indice(Grh.fC).Y
        jw = .Width
        jh = .Height
        jtw = .TileWidth
        jth = .TileHeight
        If Grafico > 0 Then
            jg = (.Indice(Grh.fC).Grafico - .Indice(1).Grafico) + Grafico
        Else
            jg = .Indice(Grh.fC).Grafico
        End If
    End With
    
    
    If center Then
        If jtw <> 1 Then
            screen_x = screen_x - Int(jtw * base_tile_size / 2) + base_tile_size / 2
        End If
        If jth <> 1 Then
            screen_Y = screen_Y - Int(jth * base_tile_size) + base_tile_size
        End If
    End If
    
    'Draw it to device
    DXEngine_iTextureRender jg, _
        screen_x, screen_Y, _
        jw, jh, _
        rgb_list, _
        jx, jy, _
         jw, jh, _
        0, _
        0
End Sub
Public Sub Anim_Render(ByRef Grh As tnGrh, ByVal screen_x As Long, ByVal screen_Y As Long, ByRef rgb_list() As Long, Optional ByVal center As Boolean, Optional ByVal Anim As Boolean, Optional ByVal Grafico As Integer)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/28/2003
'
'**************************************************************
    Dim tile_width As Single
    Dim tile_height As Single
    Dim grh_index As Long
    Dim jx As Integer
    Dim jy As Integer
    Dim jw As Integer
    Dim jh As Integer
    Dim jtw As Single
    Dim jth As Single
    Dim jg As Integer
    With NewAnimationData(Grh.index)
        If Anim Then Grh.fC = Grh.fC + (timer_ticks_per_frame * .NumFrames / .Velocidad)
        If Grh.fC > .NumFrames Then Grh.fC = (Grh.fC Mod .NumFrames) + 1
        If Grh.fC <= 0 Then Grh.fC = 1
        jx = .Indice(Grh.fC).X
        jy = .Indice(Grh.fC).Y
        jw = .Width
        jh = .Height
        jtw = .TileWidth
        jth = .TileHeight
        If Grafico > 0 Then
            jg = (.Indice(Grh.fC).Grafico - .Indice(1).Grafico) + Grafico
        Else
            jg = .Indice(Grh.fC).Grafico
        End If
    End With
    
    
    If center Then
        If jtw <> 1 Then
            screen_x = screen_x - Int(jtw * base_tile_size / 2) + base_tile_size / 2
        End If
        If jth <> 1 Then
            screen_Y = screen_Y - Int(jth * base_tile_size) + base_tile_size
        End If
    End If
    
    'Draw it to device
    DXEngine_TextureRender jg, _
        screen_x, screen_Y, _
        jw, jh, _
        rgb_list, _
        jx, jy, _
         jw, jh, _
        0, _
        0
End Sub



Public Function GUI_Grh_Render(ByVal grh_index As Long, X As Long, Y As Long, Optional ByVal Angle As Single, Optional ByVal alpha_blend As Boolean, Optional ByVal Color As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/15/2003
'
'**************************************************************
    Dim temp_grh As Grh
    Dim rpg_list(3) As Long

    If Grh_Check(grh_index) = False Then
        Exit Function
    End If

    rpg_list(0) = Color
    rpg_list(1) = Color
    rpg_list(2) = Color
    rpg_list(3) = Color

    Grh_Initialize temp_grh, grh_index, alpha_blend, Angle
    
    Grh_Render temp_grh, X, Y, rpg_list
    
    GUI_Grh_Render = True
End Function

Public Sub Animations_Initialize(ByVal AnimationSpeed As Single, ByVal tile_size As Integer)
    base_tile_size = tile_size
    AnimBaseSpeed = AnimationSpeed
End Sub

Public Sub AnimSpeedCalculate(ByVal timer_elapsed_time As Single)
    timer_ticks_per_frame = AnimBaseSpeed * timer_elapsed_time
End Sub

Public Function Grh_Check(ByVal grh_index As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'**************************************************************
    'check grh_index
    If grh_index > 0 And grh_index <= grh_count Then
        If grh_list(grh_index).Active Then
            Grh_Check = True
        End If
    End If
End Function

Public Function GetMMColor(ByVal GrhIndex As Long) As Long
GetMMColor = grh_list(GrhIndex).MiniMap_color
End Function


