Attribute VB_Name = "modOreEngine"
'Particle Groups


 
'RGB Type
Public Type RGB
    R As Long
    G As Long
    B As Long
End Type
 

'index de la particula que debe ser = que le pusimos al server
Public Enum ParticulaMedit
    CHICO = 34
    MEDIANO = 35
    GRANDE = 37
    XGRANDE = 38
    XXGRANDE = 39
End Enum
 
'Old fashion BitBlt function
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
'Added by Juan Martín Sotuyo Dodero
Public Declare Function StretchBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

Sub CargarParticulas()
    Dim StreamFile As String
    Dim LoopC As Long
    Dim i As Long
    Dim GrhListing As String
    Dim TempSet As String
    Dim ColorSet As Long
   
    StreamFile = DirIndex & "Particles.ini"
    TotalStreams = Val(general_var_get(StreamFile, "INIT", "Total"))
 
    'resize StreamData array
    ReDim StreamData(1 To TotalStreams) As Stream
 
    'fill StreamData array with info from Particles.ini
    For LoopC = 1 To TotalStreams
        StreamData(LoopC).Name = general_var_get(StreamFile, Val(LoopC), "Name")
        StreamData(LoopC).NumOfParticles = general_var_get(StreamFile, Val(LoopC), "NumOfParticles")
        StreamData(LoopC).x1 = general_var_get(StreamFile, Val(LoopC), "X1")
        StreamData(LoopC).y1 = general_var_get(StreamFile, Val(LoopC), "Y1")
        StreamData(LoopC).x2 = general_var_get(StreamFile, Val(LoopC), "X2")
        StreamData(LoopC).y2 = general_var_get(StreamFile, Val(LoopC), "Y2")
        StreamData(LoopC).Angle = general_var_get(StreamFile, Val(LoopC), "Angle")
        StreamData(LoopC).vecx1 = general_var_get(StreamFile, Val(LoopC), "VecX1")
        StreamData(LoopC).vecx2 = general_var_get(StreamFile, Val(LoopC), "VecX2")
        StreamData(LoopC).vecy1 = general_var_get(StreamFile, Val(LoopC), "VecY1")
        StreamData(LoopC).vecy2 = general_var_get(StreamFile, Val(LoopC), "VecY2")
        StreamData(LoopC).life1 = general_var_get(StreamFile, Val(LoopC), "Life1")
        StreamData(LoopC).life2 = general_var_get(StreamFile, Val(LoopC), "Life2")
        StreamData(LoopC).friction = general_var_get(StreamFile, Val(LoopC), "Friction")
        StreamData(LoopC).Spin = general_var_get(StreamFile, Val(LoopC), "Spin")
        StreamData(LoopC).spin_speedL = general_var_get(StreamFile, Val(LoopC), "Spin_SpeedL")
        StreamData(LoopC).spin_speedH = general_var_get(StreamFile, Val(LoopC), "Spin_SpeedH")
        StreamData(LoopC).AlphaBlend = general_var_get(StreamFile, Val(LoopC), "AlphaBlend")
        StreamData(LoopC).gravity = general_var_get(StreamFile, Val(LoopC), "Gravity")
        StreamData(LoopC).grav_strength = general_var_get(StreamFile, Val(LoopC), "Grav_Strength")
        StreamData(LoopC).bounce_strength = general_var_get(StreamFile, Val(LoopC), "Bounce_Strength")
        StreamData(LoopC).XMove = general_var_get(StreamFile, Val(LoopC), "XMove")
        StreamData(LoopC).YMove = general_var_get(StreamFile, Val(LoopC), "YMove")
        StreamData(LoopC).move_x1 = general_var_get(StreamFile, Val(LoopC), "move_x1")
        StreamData(LoopC).move_x2 = general_var_get(StreamFile, Val(LoopC), "move_x2")
        StreamData(LoopC).move_y1 = general_var_get(StreamFile, Val(LoopC), "move_y1")
        StreamData(LoopC).move_y2 = general_var_get(StreamFile, Val(LoopC), "move_y2")
        StreamData(LoopC).life_counter = general_var_get(StreamFile, Val(LoopC), "life_counter")
        StreamData(LoopC).Speed = Val(general_var_get(StreamFile, Val(LoopC), "Speed"))
        StreamData(LoopC).NumGrhs = general_var_get(StreamFile, Val(LoopC), "NumGrhs")
       
        ReDim StreamData(LoopC).grh_list(1 To StreamData(LoopC).NumGrhs)
        GrhListing = general_var_get(StreamFile, Val(LoopC), "Grh_List")
       
        For i = 1 To StreamData(LoopC).NumGrhs
            StreamData(LoopC).grh_list(i) = general_field_read(Str(i), GrhListing, 44)
        Next i
        StreamData(LoopC).grh_list(i - 1) = StreamData(LoopC).grh_list(i - 1)
        For ColorSet = 1 To 4
            TempSet = general_var_get(StreamFile, Val(LoopC), "ColorSet" & ColorSet)
            StreamData(LoopC).colortint(ColorSet - 1).R = general_field_read(1, TempSet, 44)
            StreamData(LoopC).colortint(ColorSet - 1).G = general_field_read(2, TempSet, 44)
            StreamData(LoopC).colortint(ColorSet - 1).B = general_field_read(3, TempSet, 44)
        Next ColorSet
    Next LoopC
 
End Sub



