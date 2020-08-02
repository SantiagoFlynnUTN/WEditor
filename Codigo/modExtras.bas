Attribute VB_Name = "modExtras"
Option Explicit
Public SelTexWe As Integer
Public SelTexFrame As Integer
Public PutX As Integer
Public PutY As Integer
Public SelTexRecort As Boolean
Public SelInicialX(1 To 256) As Integer
Public SelInicialY(1 To 256) As Integer
Public SelTexIndex(1 To 256) As Integer


Public Type IWEin
    Num As Integer
    X As Integer
    Y As Integer
End Type
Public Type tTexWE
    Name As String
    Ancho As Integer
    Largo As Integer
    tw As Single
    th As Single
    NumIndex As Integer
    index() As IWEin
End Type
Public NumTexWe As Integer
Public TexWE() As tTexWE

Public Graficos(1 To 33000) As Boolean
Public GrhIndex(1 To 40000) As Boolean
Public Objetos(1 To 1000) As Boolean
Public Npcs(1 To 1000) As Boolean
Public MapsT(1 To 165) As Boolean
Private nNpcs As Integer
Private nObjs As Integer
Private nGri As Integer
Private nGraficos As Integer

Public Type tNewIndex
    Estatic As Integer ' Info de estatica
    Dinamica As Integer ' Animacion
    OverWriteGrafico As Integer ' Grafico
End Type
Public Type ttNewIndex
    Estatic As Integer ' Info de estatica
    temp As Byte
    Dinamica As Integer ' Animacion
    OverWriteGrafico As Integer ' Grafico
    Replace As Integer
End Type
Public Type tReajustados
    Reajustado As Boolean
    NuevoG As Integer
    Custom As Byte
End Type
Public Type tiReaj
        YaIndexado As Boolean
    Indice As Integer
End Type
Public iReaj(1 To 40000) As tiReaj
Public Reaj(1 To 33000) As tReajustados
Public nReaj As Integer
Public Type tNewEstatic
    W As Integer
    H As Integer
    L As Integer
    t As Integer
    th As Single
    tw As Single
End Type
Public Type ttNewEstatic
    W As Integer
    H As Integer
    L As Integer
    t As Integer
    th As Single
    tw As Single
    Replace As Integer
End Type
Public NewIndexData() As tNewIndex
Public numNewIndex As Integer
Public EstaticData() As tNewEstatic
Public TempEstatic() As ttNewEstatic
Public TempIndex() As ttNewIndex
Public ntEstatic As Integer
Public ntIndex As Integer

Public numNewEstatic As Integer
Public Sub LoadTexWe()
Dim j As Long
Dim P As Long
Dim s As String
s = App.PATH & "\Resources\INIT\TexWe.dat"
    NumTexWe = Val(GetVar(s, "INIT", "NUM"))
    
    If NumTexWe > 0 Then
    ReDim TexWE(1 To NumTexWe)
        For P = 1 To NumTexWe
        
            With TexWE(P)
            
                .Name = GetVar(s, CStr(P), "Name")
                .NumIndex = Val(GetVar(s, CStr(P), "NumIndex"))
                If .NumIndex > 0 Then
                    ReDim .index(1 To .NumIndex)
                    For j = 1 To .NumIndex
                        .index(j).Num = Val(GetVar(s, CStr(P), "Index" & j & "_Num"))
                        .index(j).X = Val(GetVar(s, CStr(P), "Index" & j & "_X"))
                        .index(j).Y = Val(GetVar(s, CStr(P), "Index" & j & "_Y"))
                    Next j
                End If
                .Ancho = Val(GetVar(s, CStr(P), "Ancho"))
                .Largo = Val(GetVar(s, CStr(P), "Largo"))
                .th = Val(GetVar(s, CStr(P), "TH"))
                .tw = Val(GetVar(s, CStr(P), "TW"))
                
            
                frmMain.lListado(0).AddItem .Name & " - [" & P & "]"
            End With
    

        Next P
    End If
    
End Sub
Public Sub GuardarTex(ByVal TODAS As Boolean, Optional ByVal CUAL As Long)
Dim P As Long
Dim s As String
s = App.PATH & "\Resources\init\texwe.dat"
WriteVar s, "INIT", "NUM", CStr(NumTexWe)
If TODAS Then
    For P = 1 To NumTexWe
        d_guardarTex s, P
    Next P
Else
    d_guardarTex s, CUAL
End If
End Sub
Public Sub d_guardarTex(ByVal s As String, ByVal P As Long)
Dim k As Long
With TexWE(P)


    Call WriteVar(s, CStr(P), "Name", .Name)
    Call WriteVar(s, CStr(P), "Ancho", CStr(.Ancho))
    Call WriteVar(s, CStr(P), "Largo", CStr(.Largo))
    Call WriteVar(s, CStr(P), "NumIndex", CStr(.NumIndex))
    For k = 1 To TexWE(P).NumIndex
    
        Call WriteVar(s, CStr(P), "INDEX" & k & "_NUM", CStr(.index(k).Num))
        Call WriteVar(s, CStr(P), "INDEX" & k & "_x", CStr(.index(k).X))
        Call WriteVar(s, CStr(P), "INDEX" & k & "_y", CStr(.index(k).Y))
        
        
    Next k

End With


End Sub
Public Sub SepararDecors()
Dim P As Long
Dim t As Integer
Dim s As String
Dim k As String
Dim z As String
Dim numDecor As Integer
For P = 1 To 1000

    If Objetos(P) Then
        t = FreeFile
        z = vbNullString
        k = vbNullString
        
            Open App.PATH & "\Resources\Dats\Obj.dat" For Input As #t
                Do
                    Line Input #t, s
                    If left$(s, 4 + Len(CStr(P))) = "[OBJ" & P Then
                        numDecor = numDecor + 1
                        k = s & "'" & grh_list(ObjData(P).grh_index).texture_index
                        FileCopy App.PATH & "\Resources\Graficos\" & grh_list(ObjData(P).grh_index).texture_index & ".png", App.PATH & "\Resources\Graficos\DEcor\" & grh_list(ObjData(P).grh_index).texture_index & ".png"
                        Do Until EOF(t)
                            Line Input #t, s
                            If left$(s, 4) = "[OBJ" Then Exit Do
                            If left$(s, 1) <> "'" Then
                            z = z & vbCrLf & s
                            End If
                        Loop
                        

                        Exit Do
                    End If
                
                Loop
                
            Close #t
            If LenB(k) > 0 Then
                t = FreeFile
                k = k & z
                Open App.PATH & "\Resources\Dats\obj_Dec.dat" For Append As #t
                    Print #t, k
                    'Print #T, K
                
                Close #t
                t = FreeFile
                Open App.PATH & "\Resources\Dats\Dec.dat" For Append As #t
                   Print #t, "[DEC" & numDecor & "]"
                   Print #t, z
                
                Close #t
            End If
    End If

Next P


End Sub






Public Sub copiargraficosusados()
Dim P As Long
For P = 1 To 33000
    If Graficos(P) And Not Reaj(P).Reajustado Then
    
        FileCopy App.PATH & "\Resources\Graficos\" & P & ".png", App.PATH & "\Resources\Graficos Usados\" & P & ".png"
    
    End If
Next P
End Sub
Public Sub ReIndexarEstaticos()
Dim P As Long
Dim E As Integer
Dim i As Integer
For P = 1 To 40000

    If GrhIndex(P) Then
        If grh_list(P).frame_count > 1 Then
'            If grh_list(grh_list(P).frame_list(1)).texture_index = 20 Then
'                E = DameEstatico(32, 32, 0, 0)
'                i = DameIndex(E, grh_list(P).texture_index)
'            Else
            E = DameDinamico(P)
            i = DameIndex(E, grh_list(grh_list(P).frame_list(1)).texture_index, True)
'            End If
        Else
            If grh_list(P).src_height > 0 And grh_list(P).src_width > 0 Then
                'Se reajustó el grafico, no es normal.
                If Reaj(grh_list(P).texture_index).Reajustado Then
                    i = ReindexarReajustado(grh_list(P).texture_index, P)
                Else
                    E = DameEstatico(grh_list(P).src_height, grh_list(P).src_width, grh_list(P).Src_X, grh_list(P).Src_Y)
                    i = DameIndex(E, grh_list(P).texture_index)
                End If
            End If
        End If
    End If
Next P

End Sub
Public Function ReindexarReajustado(ByVal G As Integer, ByVal i As Integer) As Integer
Dim tX As Integer
Dim tY As Integer
Dim pGraf As Integer
Dim E As Integer
Dim k As Integer
Dim z As Integer
If iReaj(i).YaIndexado Then
    ReindexarReajustado = iReaj(i).Indice
    Exit Function
End If

If Reaj(G).Custom Then
    If G = 215 Then
    
        If grh_list(i).Src_X >= 512 Or grh_list(i).Src_Y >= 512 Then
        
            If grh_list(i).Src_Y >= 512 And grh_list(i).Src_X < 256 Then
                tX = grh_list(i).Src_X \ 128
                tY = (grh_list(i).Src_Y - 512) \ 64
                z = tX + tY + 1
                tX = grh_list(i).Src_X + 128
                tY = (z - 1) * 64
                E = DameEstatico(grh_list(i).src_height, grh_list(i).src_width, tX, tY)
                k = DameIndex(E, 215)

            ElseIf grh_list(i).Src_Y < 256 And grh_list(i).Src_X >= 512 Then
                tX = grh_list(i).Src_X - 512
                tY = grh_list(i).Src_Y
                
                E = DameEstatico(grh_list(i).src_height, grh_list(i).src_width, tX, tY)
                k = DameIndex(E, 215)
            ElseIf grh_list(i).Src_Y >= 512 And grh_list(i).Src_X < 512 Then
                tX = (grh_list(i).Src_X - 256) \ 128
                tY = (grh_list(i).Src_Y - 512) \ 64
                z = tX + tY + 1
                tX = grh_list(i).Src_X + 128
                tY = ((z - 1) * 64)
                E = DameEstatico(grh_list(i).src_height, grh_list(i).src_width, tX, tY)
                k = DameIndex(E, 254)
            ElseIf grh_list(i).Src_Y < 512 And grh_list(i).Src_X >= 512 Then
                tX = grh_list(i).Src_X - 512
                tY = grh_list(i).Src_Y - 256
                
                E = DameEstatico(grh_list(i).src_height, grh_list(i).src_width, tX, tY)
                k = DameIndex(E, 254)
            ElseIf grh_list(i).Src_X >= 512 And grh_list(i).Src_Y >= 512 Then
                tX = grh_list(i).Src_X - 512
                tY = grh_list(i).Src_Y - 512
                E = DameEstatico(grh_list(i).src_height, grh_list(i).src_width, tX, tY)
                k = DameIndex(E, 255)
            
            End If
            
            iReaj(i).YaIndexado = True
            iReaj(i).Indice = k
            ReindexarReajustado = k
            Exit Function
        End If
    End If
    If G = 8007 Then
        If grh_list(i).Src_Y >= 256 Then
            tX = grh_list(i).Src_X
            tY = grh_list(i).Src_Y - 256
            E = DameEstatico(grh_list(i).src_height, grh_list(i).src_width, tX, tY)
            k = DameIndex(E, 257)
        Else
            tX = grh_list(i).Src_X
            tY = grh_list(i).Src_Y
            E = DameEstatico(grh_list(i).src_height, grh_list(i).src_width, tX, tY)
            k = DameIndex(E, 256)
        End If
            iReaj(i).YaIndexado = True
            iReaj(i).Indice = k
            ReindexarReajustado = k
            Exit Function
    End If

    If G = 7522 Then
        If grh_list(i).Src_X < 192 Then
            E = DameEstatico(grh_list(i).src_height, grh_list(i).src_width, grh_list(i).Src_X, grh_list(i).Src_Y)
            k = DameIndex(E, 258)
        Else
            If grh_list(i).Src_Y >= 96 Then
                tX = 96
            Else
                tX = 0
            End If
            tY = 128
            
            E = DameEstatico(grh_list(i).src_height, grh_list(i).src_width, tX, tY)
            k = DameIndex(E, 257)
            
        End If
    
                iReaj(i).YaIndexado = True
            iReaj(i).Indice = k
            ReindexarReajustado = k
            Exit Function
    
    End If
    If G = 7100 Then
        If grh_list(i).Src_X >= 256 Then
            tX = grh_list(i).Src_X - 256
            tY = grh_list(i).Src_Y + 128
        Else
            tX = grh_list(i).Src_X
            tY = grh_list(i).Src_Y
        End If
        E = DameEstatico(grh_list(i).src_height, grh_list(i).src_width, tX, tY)
        k = DameIndex(E, 259)
        iReaj(i).YaIndexado = True
        iReaj(i).Indice = k
        ReindexarReajustado = k
        Exit Function
    End If
    If G = 6008 Then
        If grh_list(i).Src_X >= 256 Then
            tX = grh_list(i).Src_X - 256
            tY = grh_list(i).Src_Y + 128
        Else
            tX = grh_list(i).Src_X
            tY = grh_list(i).Src_Y
        End If
        E = DameEstatico(grh_list(i).src_height, grh_list(i).src_width, tX, tY)
        k = DameIndex(E, 260)
        iReaj(i).YaIndexado = True
        iReaj(i).Indice = k
        ReindexarReajustado = k
        Exit Function
    End If
    If G = 6007 Then
        If grh_list(i).Src_X >= 256 Then
            tX = grh_list(i).Src_X - 256
            tY = grh_list(i).Src_Y + 128
        Else
            tX = grh_list(i).Src_X
            tY = grh_list(i).Src_Y
        End If
        E = DameEstatico(grh_list(i).src_height, grh_list(i).src_width, tX, tY)
        k = DameIndex(E, 261)
        iReaj(i).YaIndexado = True
        iReaj(i).Indice = k
        ReindexarReajustado = k
        Exit Function
    End If
    If G = 6006 Then
        If grh_list(i).Src_X >= 256 Then
            tX = grh_list(i).Src_X - 256
            tY = grh_list(i).Src_Y + 128
        Else
            tX = grh_list(i).Src_X
            tY = grh_list(i).Src_Y
        End If
        E = DameEstatico(grh_list(i).src_height, grh_list(i).src_width, tX, tY)
        k = DameIndex(E, 262)
        iReaj(i).YaIndexado = True
        iReaj(i).Indice = k
        ReindexarReajustado = k
        Exit Function
    End If
    If G = 6005 Then
        If grh_list(i).Src_X >= 256 Then
            tX = grh_list(i).Src_X - 256
            tY = grh_list(i).Src_Y + 128
        Else
            tX = grh_list(i).Src_X
            tY = grh_list(i).Src_Y
        End If
        E = DameEstatico(grh_list(i).src_height, grh_list(i).src_width, tX, tY)
        k = DameIndex(E, 263)
        iReaj(i).YaIndexado = True
        iReaj(i).Indice = k
        ReindexarReajustado = k
        Exit Function
    End If
    If G = 6001 Then
        If grh_list(i).Src_X >= 256 Then
            tX = grh_list(i).Src_X - 256
            tY = grh_list(i).Src_Y + 128
        Else
            tX = grh_list(i).Src_X
            tY = grh_list(i).Src_Y
        End If
        E = DameEstatico(grh_list(i).src_height, grh_list(i).src_width, tX, tY)
        k = DameIndex(E, 264)
        iReaj(i).YaIndexado = True
        iReaj(i).Indice = k
        ReindexarReajustado = k
        Exit Function
    End If
    If G = 6000 Then
        If grh_list(i).Src_X >= 256 Then
            tX = grh_list(i).Src_X - 256
            tY = grh_list(i).Src_Y + 128
        Else
            tX = grh_list(i).Src_X
            tY = grh_list(i).Src_Y
        End If
        E = DameEstatico(grh_list(i).src_height, grh_list(i).src_width, tX, tY)
        k = DameIndex(E, 265)
        iReaj(i).YaIndexado = True
        iReaj(i).Indice = k
        ReindexarReajustado = k
        Exit Function
    End If
    If G = 27066 Then
        tX = 150
        tY = 0
    
        E = DameEstatico(grh_list(i).src_height, grh_list(i).src_width, tX, tY)
        k = DameIndex(E, 266)
        iReaj(i).YaIndexado = True
        iReaj(i).Indice = k
        ReindexarReajustado = k
    End If
    If G = 27033 Then
        E = DameEstatico(grh_list(i).src_height, grh_list(i).src_width, 0, 0)
        k = DameIndex(E, 266)
        iReaj(i).YaIndexado = True
        iReaj(i).Indice = k
        ReindexarReajustado = k
    End If
    If G = 27087 Then
        tY = 130
        tX = 0
        E = DameEstatico(grh_list(i).src_height, grh_list(i).src_width, tX, tY)
        k = DameIndex(E, 266)
        iReaj(i).YaIndexado = True
        iReaj(i).Indice = k
        ReindexarReajustado = k
    End If
    If G >= 27205 And G <= 27208 Then
        tX = (G - 27205) * 64
        tY = 0
        E = DameEstatico(grh_list(i).src_height, grh_list(i).src_width, tX, tY)
        k = DameIndex(E, 267)
        iReaj(i).YaIndexado = True
        iReaj(i).Indice = k
        ReindexarReajustado = k
    End If
    If G = 6002 Then
        If grh_list(i).Src_X >= 256 Then
            tX = grh_list(i).Src_X - 256
            tY = grh_list(i).Src_Y + 128
        Else
            tX = grh_list(i).Src_X
            tY = grh_list(i).Src_Y
        End If
        E = DameEstatico(grh_list(i).src_height, grh_list(i).src_width, tX, tY)
        k = DameIndex(E, 268)
        iReaj(i).YaIndexado = True
        iReaj(i).Indice = k
        ReindexarReajustado = k
        Exit Function
    End If
    If G = 6003 Then
        If grh_list(i).Src_X >= 256 Then
            tX = grh_list(i).Src_X - 256
            tY = grh_list(i).Src_Y + 128
        Else
            tX = grh_list(i).Src_X
            tY = grh_list(i).Src_Y
        End If
        E = DameEstatico(grh_list(i).src_height, grh_list(i).src_width, tX, tY)
        k = DameIndex(E, 269)
        iReaj(i).YaIndexado = True
        iReaj(i).Indice = k
        ReindexarReajustado = k
        Exit Function
    End If
    If G = 6004 Then
        If grh_list(i).Src_X >= 256 Then
            tX = grh_list(i).Src_X - 256
            tY = grh_list(i).Src_Y + 128
        Else
            tX = grh_list(i).Src_X
            tY = grh_list(i).Src_Y
        End If
        E = DameEstatico(grh_list(i).src_height, grh_list(i).src_width, tX, tY)
        k = DameIndex(E, 270)
        iReaj(i).YaIndexado = True
        iReaj(i).Indice = k
        ReindexarReajustado = k
        Exit Function
    End If
    If G = 6020 Then
        If grh_list(i).Src_X >= 256 Then
            tX = grh_list(i).Src_X - 256
            tY = grh_list(i).Src_Y + 128
        Else
            tX = grh_list(i).Src_X
            tY = grh_list(i).Src_Y
        End If
        E = DameEstatico(grh_list(i).src_height, grh_list(i).src_width, tX, tY)
        k = DameIndex(E, 271)
        iReaj(i).YaIndexado = True
        iReaj(i).Indice = k
        ReindexarReajustado = k
        Exit Function
    End If
        If G = 7006 Then
        If grh_list(i).Src_X >= 256 Then
            tX = grh_list(i).Src_X - 256
            tY = 0
        End If
        E = DameEstatico(grh_list(i).src_height, grh_list(i).src_width, tX, tY)
        k = DameIndex(E, 272)
        iReaj(i).YaIndexado = True
        iReaj(i).Indice = k
        ReindexarReajustado = k
        Exit Function
    End If
        If G = 7005 Then
        If grh_list(i).Src_X >= 256 Then
            tX = 0
            tY = 0
        End If
        E = DameEstatico(grh_list(i).src_height, grh_list(i).src_width, tX, tY)
        k = DameIndex(E, 272)
        iReaj(i).YaIndexado = True
        iReaj(i).Indice = k
        ReindexarReajustado = k
        Exit Function
    End If
        If G = 7007 Then
        If grh_list(i).Src_X >= 256 Then
            tX = grh_list(i).Src_X - 256
            tY = 0
        End If
        E = DameEstatico(grh_list(i).src_height, grh_list(i).src_width, tX, tY)
        k = DameIndex(E, 273)
        iReaj(i).YaIndexado = True
        iReaj(i).Indice = k
        ReindexarReajustado = k
        Exit Function
    End If
        If G = 7008 Then
        If grh_list(i).Src_X >= 256 Then
            tX = 0
            tY = 0
        End If
        E = DameEstatico(grh_list(i).src_height, grh_list(i).src_width, tX, tY)
        k = DameIndex(E, 273)
        iReaj(i).YaIndexado = True
        iReaj(i).Indice = k
        ReindexarReajustado = k
        Exit Function
    End If
        If G = 16000 Then
        If grh_list(i).Src_X >= 256 Then
            tX = 0
            tY = 0
        End If
        E = DameEstatico(grh_list(i).src_height, grh_list(i).src_width, tX, tY)
        k = DameIndex(E, 274)
        iReaj(i).YaIndexado = True
        iReaj(i).Indice = k
        ReindexarReajustado = k
        Exit Function
    End If
        If G = 16006 Then
        If grh_list(i).Src_X >= 256 Then
            tX = 0
            tY = 0
        End If
        E = DameEstatico(grh_list(i).src_height, grh_list(i).src_width, tX, tY)
        k = DameIndex(E, 275)
        iReaj(i).YaIndexado = True
        iReaj(i).Indice = k
        ReindexarReajustado = k
        Exit Function
    End If
        If G = 16007 Then
        If grh_list(i).Src_X >= 256 Then
            tX = 0
            tY = 0
        End If
        E = DameEstatico(grh_list(i).src_height, grh_list(i).src_width, tX, tY)
        k = DameIndex(E, 276)
        iReaj(i).YaIndexado = True
        iReaj(i).Indice = k
        ReindexarReajustado = k
        Exit Function
    End If
        If G = 5560 Then
        If grh_list(i).Src_X >= 256 Then
            tX = 0
            tY = 0
        End If
        E = DameEstatico(grh_list(i).src_height, grh_list(i).src_width, tX, tY)
        k = DameIndex(E, 277)
        iReaj(i).YaIndexado = True
        iReaj(i).Indice = k
        ReindexarReajustado = k
        Exit Function
    End If
            If G = 5561 Then
        If grh_list(i).Src_X >= 256 Then
            tX = 0
            tY = 0
        End If
        E = DameEstatico(grh_list(i).src_height, grh_list(i).src_width, tX, tY)
        k = DameIndex(E, 278)
        iReaj(i).YaIndexado = True
        iReaj(i).Indice = k
        ReindexarReajustado = k
        Exit Function
    End If
            If G = 5562 Then
        If grh_list(i).Src_X >= 256 Then
            tX = 0
            tY = 0
        End If
        E = DameEstatico(grh_list(i).src_height, grh_list(i).src_width, tX, tY)
        k = DameIndex(E, 279)
        iReaj(i).YaIndexado = True
        iReaj(i).Indice = k
        ReindexarReajustado = k
        Exit Function
    End If
            If G = 13044 Then
        If grh_list(i).Src_X >= 256 Then
            tX = 0
            tY = 0
        End If
        E = DameEstatico(grh_list(i).src_height, grh_list(i).src_width, tX, tY)
        k = DameIndex(E, 280)
        iReaj(i).YaIndexado = True
        iReaj(i).Indice = k
        ReindexarReajustado = k
        Exit Function
    End If
                If G = 13045 Then
        If grh_list(i).Src_X >= 256 Then
            tX = 0
            tY = 0
        End If
        E = DameEstatico(grh_list(i).src_height, grh_list(i).src_width, tX, tY)
        k = DameIndex(E, 281)
        iReaj(i).YaIndexado = True
        iReaj(i).Indice = k
        ReindexarReajustado = k
        Exit Function
    End If
                If G = 13046 Then
        If grh_list(i).Src_X >= 256 Then
            tX = 0
            tY = 0
        End If
        E = DameEstatico(grh_list(i).src_height, grh_list(i).src_width, tX, tY)
        k = DameIndex(E, 282)
        iReaj(i).YaIndexado = True
        iReaj(i).Indice = k
        ReindexarReajustado = k
        Exit Function
    End If
                If G = 20509 Then
        If grh_list(i).Src_X >= 256 Then
            tX = 0
            tY = 0
        End If
        E = DameEstatico(grh_list(i).src_height, grh_list(i).src_width, tX, tY)
        k = DameIndex(E, 283)
        iReaj(i).YaIndexado = True
        iReaj(i).Indice = k
        ReindexarReajustado = k
        Exit Function
    End If
    If G = 665 Then
        If grh_list(i).Src_X < 172 Then
            tX = grh_list(i).Src_X
            tY = 0
        Else
            tX = grh_list(i).Src_X - 172
            tY = 64
        End If
        E = DameEstatico(grh_list(i).src_height, grh_list(i).src_width, tX, tY)
        k = DameIndex(E, 665)
        iReaj(i).YaIndexado = True
        iReaj(i).Indice = k
        ReindexarReajustado = k
    
    
    End If
                If G = 1530 Then
        If grh_list(i).Src_X >= 256 Then
            tX = 0
            tY = 0
        End If
        E = DameEstatico(grh_list(i).src_height, grh_list(i).src_width, tX, tY)
        k = DameIndex(E, 284)
        iReaj(i).YaIndexado = True
        iReaj(i).Indice = k
        ReindexarReajustado = k
        Exit Function
    End If
                    If G = 1531 Then
        If grh_list(i).Src_X >= 256 Then
            tX = 0
            tY = 0
        End If
        E = DameEstatico(grh_list(i).src_height, grh_list(i).src_width, tX, tY)
        k = DameIndex(E, 285)
        iReaj(i).YaIndexado = True
        iReaj(i).Indice = k
        ReindexarReajustado = k
        Exit Function
    End If
                    If G = 1532 Then
        If grh_list(i).Src_X >= 256 Then
            tX = 0
            tY = 0
        End If
        E = DameEstatico(grh_list(i).src_height, grh_list(i).src_width, tX, tY)
        k = DameIndex(E, 286)
        iReaj(i).YaIndexado = True
        iReaj(i).Indice = k
        ReindexarReajustado = k
        Exit Function
    End If
                    If G = 6018 Then
        If grh_list(i).Src_X >= 256 Then
            tX = 0
            tY = 0
        End If
        E = DameEstatico(grh_list(i).src_height, grh_list(i).src_width, tX, tY)
        k = DameIndex(E, 287)
        iReaj(i).YaIndexado = True
        iReaj(i).Indice = k
        ReindexarReajustado = k
        Exit Function
    End If
                    If G = 6019 Then
        If grh_list(i).Src_X >= 256 Then
            tX = 0
            tY = 0
        End If
        E = DameEstatico(grh_list(i).src_height, grh_list(i).src_width, tX, tY)
        k = DameIndex(E, 288)
        iReaj(i).YaIndexado = True
        iReaj(i).Indice = k
        ReindexarReajustado = k
        Exit Function
    End If
                    If G = 7501 Then
        If grh_list(i).Src_X >= 256 Then
            tX = 0
            tY = 0
        End If
        E = DameEstatico(grh_list(i).src_height, grh_list(i).src_width, tX, tY)
        k = DameIndex(E, 289)
        iReaj(i).YaIndexado = True
        iReaj(i).Indice = k
        ReindexarReajustado = k
        Exit Function
    End If
                    If G = 7502 Then
        If grh_list(i).Src_X >= 256 Then
            tX = 0
            tY = 0
        End If
        E = DameEstatico(grh_list(i).src_height, grh_list(i).src_width, tX, tY)
        k = DameIndex(E, 290)
        iReaj(i).YaIndexado = True
        iReaj(i).Indice = k
        ReindexarReajustado = k
        Exit Function
    End If
                    If G = 7503 Then
        If grh_list(i).Src_X >= 256 Then
            tX = 0
            tY = 0
        End If
        E = DameEstatico(grh_list(i).src_height, grh_list(i).src_width, tX, tY)
        k = DameIndex(E, 291)
        iReaj(i).YaIndexado = True
        iReaj(i).Indice = k
        ReindexarReajustado = k
        Exit Function
    End If
                    If G = 7504 Then
        If grh_list(i).Src_X >= 256 Then
            tX = 0
            tY = 0
        End If
        E = DameEstatico(grh_list(i).src_height, grh_list(i).src_width, tX, tY)
        k = DameIndex(E, 292)
        iReaj(i).YaIndexado = True
        iReaj(i).Indice = k
        ReindexarReajustado = k
        Exit Function
    End If
                        If G = 7505 Then
        If grh_list(i).Src_X >= 256 Then
            tX = 0
            tY = 0
        End If
        E = DameEstatico(grh_list(i).src_height, grh_list(i).src_width, tX, tY)
        k = DameIndex(E, 293)
        iReaj(i).YaIndexado = True
        iReaj(i).Indice = k
        ReindexarReajustado = k
        Exit Function
    End If
                        If G = 7506 Then
        If grh_list(i).Src_X >= 256 Then
            tX = 0
            tY = 0
        End If
        E = DameEstatico(grh_list(i).src_height, grh_list(i).src_width, tX, tY)
        k = DameIndex(E, 294)
        iReaj(i).YaIndexado = True
        iReaj(i).Indice = k
        ReindexarReajustado = k
        Exit Function
    End If
        If G = 7507 Then
        If grh_list(i).Src_X >= 256 Then
            tX = 0
            tY = 0
        End If
        E = DameEstatico(grh_list(i).src_height, grh_list(i).src_width, tX, tY)
        k = DameIndex(E, 295)
        iReaj(i).YaIndexado = True
        iReaj(i).Indice = k
        ReindexarReajustado = k
        Exit Function
    End If
        If G = 7509 Then
            tX = 0
            tY = 0

        E = DameEstatico(grh_list(i).src_height, grh_list(i).src_width, tX, tY)
        k = DameIndex(E, 296)
        iReaj(i).YaIndexado = True
        iReaj(i).Indice = k
        ReindexarReajustado = k
        Exit Function
    End If
        If G = 7510 Then

            tX = 288
            tY = 0

        E = DameEstatico(grh_list(i).src_height, grh_list(i).src_width, tX, tY)
        k = DameIndex(E, 296)
        iReaj(i).YaIndexado = True
        iReaj(i).Indice = k
        ReindexarReajustado = k
        Exit Function
    End If
        If G = 7004 Then

            tX = 256
            tY = 0

        E = DameEstatico(grh_list(i).src_height, grh_list(i).src_width, tX, tY)
        k = DameIndex(E, 5517)
        iReaj(i).YaIndexado = True
        iReaj(i).Indice = k
        ReindexarReajustado = k
        Exit Function
    End If
End If

tX = (grh_list(i).Src_X \ 256)
tY = (grh_list(i).Src_Y \ 256)

pGraf = (tY * 2) + tX
pGraf = Reaj(G).NuevoG + pGraf

Dim sx As Integer
Dim sy As Integer
sx = grh_list(i).Src_X - (256 * tX)
sy = grh_list(i).Src_Y - (256 * tY)

E = DameEstatico(grh_list(i).src_height, grh_list(i).src_width, sx, sy)

k = DameIndex(E, pGraf)
ReindexarReajustado = k
iReaj(i).YaIndexado = True
iReaj(i).Indice = k


End Function
Public Function DameIndex(ByVal E As Integer, ByVal Grafico As Integer, Optional ByVal Dinamica As Boolean)
Dim P As Long
    If numNewIndex > 0 Then
        For P = 1 To numNewIndex
        If Dinamica = False Then
            If NewIndexData(P).Estatic = E Then
                If NewIndexData(P).OverWriteGrafico = Grafico Then
                    Exit For
                End If
            End If
        Else
            If NewIndexData(P).Dinamica = E Then
                If NewIndexData(P).OverWriteGrafico = Grafico Then
                    Exit For
                End If
            End If
        End If
        Next P
    Else
        P = 1
    End If
    If P <= numNewIndex Then
        DameIndex = P
    Else
        'Hay q indexar
        numNewIndex = P
        ReDim Preserve NewIndexData(1 To numNewIndex)
        With NewIndexData(P)
            If Dinamica Then
            .Dinamica = E
            Else
            .Estatic = E
            End If
            .OverWriteGrafico = Grafico
        End With
        If Not Dinamica Then
        WriteVar App.PATH & "\Resources\Init\NewIndex.dat", CStr(P), "Estatica", CStr(E)
        WriteVar App.PATH & "\Resources\Init\NewIndex.dat", CStr(P), "Size", EstaticData(E).L & "-" & EstaticData(E).t & "-" & EstaticData(E).W & "-" & EstaticData(E).H
        
        Else
        WriteVar App.PATH & "\Resources\Init\NewIndex.dat", CStr(P), "Dinamica", CStr(E)

        End If
        WriteVar App.PATH & "\Resources\Init\NewIndex.dat", CStr(P), "OverWriteGrafico", CStr(Grafico)
        WriteVar App.PATH & "\Resources\Init\NewIndex.dat", "INIT", "Num", CStr(P)
        DameIndex = P
        
    End If
End Function
Public Function DameDinamico(ByVal G As Integer) As Integer
Dim P As Long
Dim z As Long
If Num_NwAnim > 0 Then
    For P = 1 To Num_NwAnim

        If NewAnimationData(P).NumFrames = grh_list(G).frame_count Then
            If NewAnimationData(P).Height = grh_list(grh_list(G).frame_list(1)).src_height Then
                If NewAnimationData(P).Width = grh_list(grh_list(G).frame_list(1)).src_width Then
                    If grh_list(grh_list(G).frame_list(1)).texture_index = 20 Then
                                        If NewAnimationData(P).Indice(1).X = ((((G - 1505) * 4) Mod 8) * 32) And _
                                            NewAnimationData(P).Indice(1).Y = ((((G - 1505) * 4) \ 8) * 32) Then
                                                Exit For
                                        End If
                        
                        
                    Else

                    For z = 1 To NewAnimationData(P).NumFrames
                        If NewAnimationData(P).Indice(z).X <> grh_list(grh_list(G).frame_list(z)).Src_X And _
                            NewAnimationData(P).Indice(z).Y <> grh_list(grh_list(G).frame_list(z)).Src_Y Then
                                Exit For
                        End If
                    
                    Next z
                    If z > NewAnimationData(P).NumFrames Then Exit For
                    
                    End If
                End If
            End If
        End If
    
    Next P


Else

    P = 1
End If

If P <= Num_NwAnim Then
    DameDinamico = P

Else
    'hay q indexarlo.
    Num_NwAnim = Num_NwAnim + 1
    ReDim Preserve NewAnimationData(1 To Num_NwAnim)
    
    With NewAnimationData(Num_NwAnim)

        .Width = grh_list(grh_list(G).frame_list(1)).src_width
        .Height = grh_list(grh_list(G).frame_list(1)).src_height
        .TileHeight = .Height / 32
        .TileWidth = .Width / 32
        .NumFrames = grh_list(G).frame_count
        .Velocidad = 1

        
        Dim lx As Long
        Dim nc As Long
        Dim nf As Long
        Dim ly As Long
        Dim k As Long
        Dim kc As Long
        Dim j As Long
        Dim jc As Long
        Dim s As String
        s = App.PATH & "\Resources\Init\NewAnim.dat"
        If grh_list(grh_list(G).frame_list(1)).texture_index = 20 Then

            .Columnas = 8
            .Filas = 8
            nc = 8
            nf = 8
            ReDim .Indice(1 To .NumFrames)
            
            k = G - 1505
            k = 4 * k
            
            
            For z = 1 To .NumFrames
                .Indice(z).X = (((k + (z - 1)) Mod 8) * 32)
                .Indice(z).Y = ((k + (z - 1)) \ 8) * 32
                .Indice(z).Grafico = 0
            
            
            Next z
            k = k + 1
            
            WriteVar s, "ANIMACION" & Num_NwAnim, "Inicial", CStr(k)
        Else
        For z = 1 To .NumFrames
            k = (grh_list(grh_list(G).frame_list(z)).Src_X / grh_list(grh_list(G).frame_list(z)).src_width)
            If k > kc Then kc = k
            If grh_list(grh_list(G).frame_list(z)).Src_X > lx Then nc = nc + 1
            j = (grh_list(grh_list(G).frame_list(z)).Src_Y / grh_list(grh_list(G).frame_list(z)).src_height)
            If j > jc Then jc = j
            If grh_list(grh_list(G).frame_list(z)).Src_Y > ly Then nf = nf + 1
        Next z
        jc = jc + 1
        kc = kc + 1
        nc = nc + 1
        nf = nf + 1
        
        If jc > nf Then nf = jc
        If kc > nc Then nc = kc
        
        .Columnas = nc
        .Filas = nf
        ReDim .Indice(1 To .NumFrames)
        For z = 1 To .NumFrames
            .Indice(z).X = grh_list(grh_list(G).frame_list(z)).Src_X
            .Indice(z).Y = grh_list(grh_list(G).frame_list(z)).Src_Y
            .Indice(z).Grafico = ((z - 1) \ (nc * nf))
        
        Next z
        End If


        WriteVar s, "NW_ANIM", "NUM", CStr(Num_NwAnim)
        
        WriteVar s, "ANIMACION" & Num_NwAnim, "Filas", CStr(nf)
        WriteVar s, "ANIMACION" & Num_NwAnim, "Columnas", CStr(nc)
        WriteVar s, "ANIMACION" & Num_NwAnim, "Ancho", CStr(.Width)
        WriteVar s, "ANIMACION" & Num_NwAnim, "Alto", CStr(.Height)
        WriteVar s, "ANIMACION" & Num_NwAnim, "NumeroFrames", CStr(.NumFrames)
        WriteVar s, "ANIMACION" & Num_NwAnim, "Velocidad", CStr(.Velocidad)
        
        
        
        
        
    End With
    DameDinamico = P
End If
End Function
Public Function DameEstatico(ByVal H As Integer, ByVal W As Integer, ByVal X As Integer, ByVal Y As Integer) As Integer
Dim P As Long
    If numNewEstatic > 0 Then
        For P = 1 To numNewEstatic
            If EstaticData(P).L = X Then
                If EstaticData(P).t = Y Then
                    If EstaticData(P).H = H Then
                        If EstaticData(P).W = W Then
                            Exit For
                        End If
                    End If
                End If
            End If
        Next P
    Else
        P = 1
    End If
    
    If P <= numNewEstatic Then
        'Ya esta indexado.
        DameEstatico = P
    Else
        'Hay q indexarlo.
        numNewEstatic = P
        ReDim Preserve EstaticData(1 To numNewEstatic)
        With EstaticData(numNewEstatic)
            .L = X
            .t = Y
            .W = W
            .H = H
            .th = H / 32
            .tw = W / 32
       End With
       WriteVar App.PATH & "\Resources\Init\NewEstatics.dat", CStr(numNewEstatic), "Left", CStr(X)
       WriteVar App.PATH & "\Resources\Init\NewEstatics.dat", CStr(numNewEstatic), "Top", CStr(Y)
       WriteVar App.PATH & "\Resources\Init\NewEstatics.dat", CStr(numNewEstatic), "Width", CStr(W)
       WriteVar App.PATH & "\Resources\Init\NewEstatics.dat", CStr(numNewEstatic), "Height", CStr(H)
       WriteVar App.PATH & "\Resources\Init\NewEstatics.dat", "INIT", "NUM", CStr(numNewEstatic)
       DameEstatico = P
    End If
    
End Function
Public Sub GuardarNpcsUsados()
Dim P As Long
Dim k As Integer
nNpcs = Val(GetVar(App.PATH & "\REGISTROS\REGISTROS.txt", "REGISTROS", "nnpcs"))
nNpcs = nNpcs + 1
WriteVar App.PATH & "\REGISTROS\REGISTROS.txt", "REGISTROS", "nNpcs", CStr(nNpcs)
    AbrirArchivo "Npcs" & nNpcs & ".txt", k
    For P = 1 To 1000
        If Npcs(P) Then _
            RegistrarLinea CStr(P), k
    Next P
End Sub
Public Sub GuardarGriUsados()
Dim P As Long
Dim k As Integer
nGri = CInt(Val(GetVar(App.PATH & "\REGISTROS\REGISTROS.txt", "REGISTROS", "nGri")))
nGri = nGri + 1
WriteVar App.PATH & "\REGISTROS\REGISTROS.txt", "REGISTROS", "Ngri", CStr(nGri)
    AbrirArchivo "gri" & nGri & " .txt", k
    For P = 1 To 40000
        If GrhIndex(P) Then _
            RegistrarLinea CStr(P), k
    Next P
Close #k
End Sub
Public Sub GuardarGraficosUsados()
Dim P As Long
Dim k As Integer
nGraficos = Val(GetVar(App.PATH & "\REGISTROS\REGISTROS.txt", "REGISTROS", "ngraficos"))
nGraficos = nGraficos + 1
WriteVar App.PATH & "\REGISTROS\REGISTROS.txt", "REGISTROS", "NGraficos", CStr(nGraficos)
    AbrirArchivo "Graficos" & nGraficos & ".txt", k
    For P = 1 To 33000
        If Graficos(P) Then _
            RegistrarLinea CStr(P), k
    Next P
    Close #k
End Sub
Public Sub GuardarGrhTracked()
Dim k As Integer
Dim P As Long
Dim FrmExtrast2 As String

AbrirArchivo "GrhTracked" & FrmExtrast2 & ".txt", k

For P = 1 To 165
    If MapsT(P) Then
        RegistrarLinea CStr(P), k
        MapsT(P) = False
    End If
Next P
Close #k
End Sub
Public Sub GuardarObjetosUsados()
Dim P As Long
Dim k As Integer
nObjs = Val(GetVar(App.PATH & "\REGISTROS\REGISTROS.txt", "REGISTROS", "nObj"))
nObjs = nObjs + 1
WriteVar App.PATH & "\REGISTROS\REGISTROS.txt", "REGISTROS", "NOBJ", CStr(nObjs)
    AbrirArchivo "Objetos" & nObjs & ".txt", k
    For P = 1 To 1000
        If Objetos(P) Then _
            RegistrarLinea CStr(P), k
    Next P
    Close #k
End Sub
Public Sub AbrirArchivo(ByVal nombre As String, ByRef i As Integer)
    i = FreeFile
    
    Open App.PATH & "\Registros\" & nombre For Output As #i
    
    

End Sub
Public Sub RegistrarLinea(ByVal Linea As String, ByVal i As Integer)
    
    Print #i, Linea
End Sub
Public Sub Load_NewEstatics()
Dim s As String
Dim i As Long
Dim z As Long

s = App.PATH & "\Resources\INIT\NewEstatics.dat"

numNewEstatic = Val(GetVar(s, "INIT", "num"))

If numNewEstatic > 0 Then
ReDim EstaticData(1 To numNewEstatic)
For i = 1 To numNewEstatic
    With EstaticData(i)
        .L = Val(GetVar(s, CStr(i), "Left"))
        .t = Val(GetVar(s, CStr(i), "Top"))
        .W = Val(GetVar(s, CStr(i), "Width"))
        .H = Val(GetVar(s, CStr(i), "Height"))
        .tw = .W / 32
        .th = .H / 32
    End With
Next i
End If

End Sub
Public Sub Load_NewIndex()
Dim s As String
Dim i As Long
Dim z As Long

s = App.PATH & "\Resources\INIT\NewIndex.dat"

numNewIndex = Val(GetVar(s, "INIT", "num"))

If numNewIndex > 0 Then
ReDim NewIndexData(1 To numNewIndex)
For i = 1 To numNewIndex
    With NewIndexData(i)
        .Dinamica = Val(GetVar(s, CStr(i), "Dinamica"))
        .Estatic = Val(GetVar(s, CStr(i), "Estatica"))
        .OverWriteGrafico = Val(GetVar(s, CStr(i), "OverWriteGrafico"))
    frmMain.lListado(5).AddItem i & " - [" & .OverWriteGrafico & "]"
    End With

Next i
End If


End Sub
Public Sub Load_Reajuste()
Dim s As String
Dim P As Long
Dim Num As Integer
s = App.PATH & "\Resources\Reajuste.txt"

nReaj = Val(GetVar(s, "INIT", "Num"))

For P = 1 To nReaj

    Num = Val(GetVar(s, CStr(P), "G"))
    Reaj(Num).Reajustado = True
    Reaj(Num).Custom = IIf(Val(GetVar(s, CStr(P), "CUSTOM")) = 1, True, False)
    Reaj(Num).NuevoG = Val(GetVar(s, CStr(P), "NuevoG"))
    

Next P
End Sub
Public Function PoneIndexEnTex(ByVal Tex As Integer, ByVal aX As Integer, ByVal aY As Integer, ByVal oX As Integer, ByVal oY As Integer) As Integer
Dim P As Long
Dim dX As Integer
Dim dy As Integer
Dim j As Integer
Dim X As Integer
Dim Y As Integer
Dim lx As Integer
Dim ly As Integer

If SelTexWe = 0 Then Exit Function

X = ((TexWE(Tex).Ancho - 1) \ 32) + 1
Y = ((TexWE(Tex).Largo - 1) \ 32) + 1

dX = aX - oX
dy = aY - oY

If Not frmConfigSup.DespMosaic.value = vbChecked Then

X = dX Mod X
Y = dy Mod Y
Else
X = (dX + Val(frmConfigSup.DMAncho.Text)) Mod X
Y = (dy + Val(frmConfigSup.DMLargo.Text)) Mod Y
End If
If X < 0 Then X = (((TexWE(Tex).Ancho - 1) \ 32) + 1) + X
If Y < 0 Then Y = (((TexWE(Tex).Largo - 1) \ 32) + 1) + Y

j = ((Y) * 16) + (X + 1)
If j > 0 Then
    If SelTexIndex(j) > 0 Then
            
        lx = (SelInicialX(j) \ 32)
        ly = (SelInicialY(j) \ 32)
        
        dX = lx - X
        dy = ly - Y
        
        If (aX + dX) <= 0 Or (aX + dX) > 100 Then Exit Function
        If (aY + dy) <= 0 Or (aY + dy) > 100 Then Exit Function
        
        PoneIndexEnTex = TexWE(Tex).index(SelTexIndex(j)).Num 'return this value

    End If
End If


End Function
Public Function DameIndexEnTex(ByVal Tex As Integer, ByVal aX As Integer, ByVal aY As Integer, ByVal oX As Integer, ByVal oY As Integer) As Integer
Dim P As Long
Dim dX As Integer
Dim dy As Integer
Dim j As Integer
Dim X As Integer
Dim Y As Integer
Dim lx As Integer
Dim ly As Integer

If SelTexWe = 0 Then Exit Function

X = ((TexWE(Tex).Ancho - 1) \ 32) + 1
Y = ((TexWE(Tex).Largo - 1) \ 32) + 1

dX = aX - oX
dy = aY - oY

If frmConfigSup.DespMosaic.value = vbChecked Then
X = (dX + Val(frmConfigSup.DMAncho.Text)) Mod X
Y = (dy + Val(frmConfigSup.DMLargo.Text)) Mod Y
Else
X = dX Mod X
Y = dy Mod Y
End If

If X < 0 Then X = (((TexWE(Tex).Ancho - 1) \ 32) + 1) + X
If Y < 0 Then Y = (((TexWE(Tex).Largo - 1) \ 32) + 1) + Y

j = ((Y) * 16) + (X + 1)
If j > 0 Then
If SelTexIndex(j) > 0 Then
        
    lx = (SelInicialX(j) \ 32)
    ly = (SelInicialY(j) \ 32)
    
    dX = lx - X
    dy = ly - Y
    
    If (aX + dX) <= 0 Or (aX + dX) > 100 Then Exit Function
    If (aY + dy) <= 0 Or (aY + dy) > 100 Then Exit Function
    
    DameIndexEnTex = TexWE(Tex).index(SelTexIndex(j)).Num

    
End If
End If


End Function
Public Function DameIndexEnTexUL(ByVal Tex As Integer) As Integer

Dim P As Long
Dim lp As Integer

If Tex = 0 Then Exit Function
With TexWE(Tex)
    If .NumIndex > 0 Then
    
        For P = 1 To .NumIndex
        
            If .index(P).X = 0 And .index(P).Y = 0 Then
                lp = P
                Exit For
            Else
                If lp > 0 Then
                If .index(P).X <= .index(lp).X And .index(P).Y <= .index(P).Y Then
                    lp = P
                End If
                Else
                lp = P
                End If
            End If
            
        Next P
        DameIndexEnTexUL = .index(lp).Num

    Else: DameIndexEnTexUL = 0
    End If
End With


End Function

Public Sub AnalizeTexture(ByVal t As Integer)
Erase SelInicialX
Erase SelInicialY
Erase SelTexIndex

Dim P As Long
Dim j As Long
Dim X As Long
Dim Y As Long
Dim kx As Long
Dim ky As Long
If t = 0 Then Exit Sub
With TexWE(t)

For P = 1 To .NumIndex

    X = (.index(P).X \ 32) + 1
    Y = .index(P).Y \ 32

    If NewIndexData(.index(P).Num).Estatic > 0 Then
        For kx = X To X + ((EstaticData(NewIndexData(.index(P).Num).Estatic).W - 1) \ 32) '((.Ancho - 1) \ 32)
            For ky = Y To Y + ((EstaticData(NewIndexData(.index(P).Num).Estatic).H - 1) \ 32)
                
                j = (ky * 16) + kx
                SelInicialX(j) = .index(P).X
                SelInicialY(j) = .index(P).Y
                SelTexIndex(j) = P
            Next ky
        Next kx
    ElseIf NewIndexData(.index(P).Num).Dinamica > 0 Then
        For kx = X To X + ((NewAnimationData(NewIndexData(.index(P).Num).Dinamica).Width - 1) \ 32) '((.Ancho - 1) \ 32)
            For ky = Y To Y + ((NewAnimationData(NewIndexData(.index(P).Num).Dinamica).Height - 1) \ 32)
                
                j = (ky * 16) + kx
                SelInicialX(j) = .index(P).X
                SelInicialY(j) = .index(P).Y
                SelTexIndex(j) = P
            Next ky
        Next kx
    
    End If
Next P

End With


End Sub
Public Function DECOR_GETGRH_FROMDEFAULT(ByVal X As Integer, ByVal Y As Integer) As Integer
Dim j As Long
With MapData(X, Y)
        If .DecorInfo.EstadoDefault > 0 Then 'Asignamos el decorgrh correcto segun estado.
            If .DecorInfo.EstadoDefault <= 5 Then
                If DecorData(.DecorI).DecorGrh(.DecorInfo.EstadoDefault) > 0 Then
                    DECOR_GETGRH_FROMDEFAULT = DecorData(.DecorI).DecorGrh(.DecorInfo.EstadoDefault)
                Else
                    For j = (.DecorInfo.EstadoDefault) To 1 Step -1
                        If DecorData(.DecorI).DecorGrh(j) > 0 Then
                            DECOR_GETGRH_FROMDEFAULT = DecorData(.DecorI).DecorGrh(j)
                            Exit For
                        End If
                    Next j
                End If
            Else
                DECOR_GETGRH_FROMDEFAULT = DecorData(.DecorI).DecorGrh(1)
            End If
        Else
            DECOR_GETGRH_FROMDEFAULT = DecorData(.DecorI).DecorGrh(1)
        End If
End With
End Function
Public Function AsignarClave(ByVal TipoClave As Byte, ByVal TipoObjeto As Byte) As Long
Select Case TipoObjeto
    Case 1 'Decors
        Select Case TipoClave
            Case 1 'Normal.
                NumDecorKeys = NumDecorKeys + 1
                ReDim Preserve DecorKeys(1 To NumDecorKeys)
                DecorKeys(NumDecorKeys).Tipo_Objeto = 1
                DecorKeys(NumDecorKeys).Tipo_Clave = 1
                DecorKeys(NumDecorKeys).Clave = UltimaDecorKey_Comun + 1
                UltimaDecorKey_Comun = UltimaDecorKey_Comun + 1
                AsignarClave = NumDecorKeys
                
        End Select
End Select
End Function
Public Function ChequearTemporal() As Boolean
Dim X As Long
Dim Y As Long
For X = 1 To 100
    For Y = 1 To 100
        With MapData(X, Y)
            If .Graphic(1).index > NumRealIndex Or _
                .Graphic(2).index > NumRealIndex Or _
                .Graphic(3).index > NumRealIndex Or _
                .Graphic(4).index > NumRealIndex Or _
                .Graphic(5).index > NumRealIndex Then
                        
                ChequearTemporal = True
                Exit Function
        
        
            End If
        End With
    Next Y
Next X

End Function
