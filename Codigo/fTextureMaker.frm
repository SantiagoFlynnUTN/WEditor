VERSION 5.00
Begin VB.Form fTextureMaker 
   Caption         =   "Textura Workshop"
   ClientHeight    =   8550
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   20250
   LinkTopic       =   "Form2"
   ScaleHeight     =   570
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1350
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command 
      Caption         =   "Limpiar"
      Height          =   375
      Left            =   16200
      TabIndex        =   45
      Top             =   120
      Width           =   1215
   End
   Begin WorldEditor.lvButtons_H csel 
      Height          =   375
      Left            =   4080
      TabIndex        =   43
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "Seleccionar"
      CapAlign        =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   1
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.CommandButton Command5 
      Caption         =   "<"
      Height          =   3735
      Left            =   11640
      TabIndex        =   42
      Top             =   4680
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Caption         =   ">"
      Height          =   3855
      Left            =   11640
      TabIndex        =   37
      Top             =   600
      Width           =   255
   End
   Begin VB.TextBox Text13 
      Height          =   375
      Left            =   15000
      TabIndex        =   36
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Text12 
      Height          =   375
      Left            =   12480
      TabIndex        =   35
      Top             =   120
      Width           =   975
   End
   Begin VB.Frame Frame3 
      Caption         =   "Indice"
      Height          =   2535
      Left            =   2280
      TabIndex        =   20
      Top             =   600
      Width           =   1575
      Begin VB.CommandButton cmdIndex 
         Caption         =   "Indexar"
         Enabled         =   0   'False
         Height          =   195
         Left            =   120
         TabIndex        =   32
         Top             =   2280
         Width           =   1335
      End
      Begin VB.CommandButton cmdReci 
         Caption         =   "Reciclado"
         Enabled         =   0   'False
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   2080
         Width           =   1335
      End
      Begin VB.TextBox Text11 
         Height          =   375
         Left            =   720
         TabIndex        =   30
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   720
         TabIndex        =   28
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox Text9 
         Height          =   375
         Left            =   720
         TabIndex        =   26
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   720
         TabIndex        =   24
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   720
         TabIndex        =   22
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "Index:"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "Height:"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Width:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Top:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Left:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.PictureBox Nuevo 
      Height          =   7800
      Left            =   11880
      ScaleHeight     =   7740
      ScaleWidth      =   7620
      TabIndex        =   19
      Top             =   600
      Width           =   7680
   End
   Begin VB.Frame Frame2 
      Caption         =   "Texturas"
      Height          =   5295
      Left            =   120
      TabIndex        =   7
      Top             =   3120
      Width           =   3735
      Begin VB.CommandButton Command4 
         Caption         =   "Nueva"
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   4680
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Guardar"
         Height          =   255
         Left            =   1920
         TabIndex        =   40
         Top             =   4680
         Width           =   1455
      End
      Begin VB.TextBox Text14 
         Height          =   285
         Left            =   960
         TabIndex        =   39
         Top             =   2600
         Width           =   2655
      End
      Begin VB.ListBox List1 
         Height          =   1230
         Left            =   240
         TabIndex        =   17
         Top             =   3360
         Width           =   3135
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   1080
         TabIndex        =   16
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1080
         TabIndex        =   14
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1080
         TabIndex        =   12
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1080
         TabIndex        =   11
         Top             =   840
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label15 
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Num Index:"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Label Label6 
         Caption         =   "TH:"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "TW:"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Largo:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Ancho:"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Grafico"
      Height          =   2535
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2055
      Begin VB.CommandButton Command1 
         Caption         =   "Buscar"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1080
         TabIndex        =   4
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1080
         TabIndex        =   2
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Indice:"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Grafico:"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.PictureBox Actual 
      Height          =   7800
      Left            =   3960
      ScaleHeight     =   7740
      ScaleWidth      =   7620
      TabIndex        =   0
      Top             =   600
      Width           =   7680
   End
   Begin WorldEditor.lvButtons_H butgrilla 
      Height          =   375
      Left            =   5640
      TabIndex        =   44
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "Grilla"
      CapAlign        =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   1
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.Label Label14 
      Caption         =   "Top:"
      Height          =   375
      Left            =   14160
      TabIndex        =   34
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label13 
      Caption         =   "Left:"
      Height          =   375
      Left            =   11760
      TabIndex        =   33
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "fTextureMaker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private aGrafico As Integer
Private agraficow As Integer
Private agraficoh As Integer
Private atx As Integer
Private nfx As Integer
Private nfy As Integer
Private aty As Integer
Private tw As Integer
Private th As Integer
Dim tX As Integer
Dim tY As Integer
Dim iX As Integer
Dim iY As Integer
Private tNumIndex As Integer
Private tIndex() As IWEin
Private tAncho As Integer
Private tAlto As Integer
Private TexAc As Long
Private TexArray(1 To 256) As Integer
Private TexInicialX(1 To 256) As Integer
Private fSelecter As Boolean
Private TexInicialY(1 To 256) As Integer
Private ifX As Integer
Private ifY As Integer
Private iiFx As Integer
Private iiFy As Integer
Private lifX As Integer
Private lifY As Integer
Private liiFx As Integer
Private liiFy As Integer
Private bGrilla As Boolean
Private Selecto As Boolean
Private SelInd() As Integer
Private SelStartX() As Integer
Private SelStartY() As Integer

Private NumSel As Integer







Private Sub Actual_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If fSelecter Then
        If agraficow > 0 And agraficoh > 0 Then
            If (X / Screen.TwipsPerPixelX) <= agraficow And (Y / Screen.TwipsPerPixelY) <= agraficoh Then
                modDXEngine.DibujareEnHwnd3 Actual.hWnd, aGrafico, 0, 0, True
                If bGrilla Then AplicarGrilla 0, 0, agraficow, agraficoh, Actual, vbGreen
                Actual.ForeColor = vbWhite
                Actual.DrawWidth = 2
                Actual.Line (iiFx, iiFy)-(iiFx, Y)
                Actual.Line (iiFx, Y)-(X, Y)
                Actual.Line (X, iiFy)-(X, Y)
                Actual.Line (iiFx, iiFy)-(X, iiFy)
                Actual.ForeColor = vbBlack
                Actual.DrawWidth = 1
            End If
        End If
    End If
    
End Sub

Private Sub Actual_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim z As Integer

    Selecto = False
    Erase SelInd
    Erase SelStartX
    Erase SelStartY
    Dim EncontroGrafico As Boolean
    Dim FoundEstatico As Boolean
    NumSel = 0


    iX = (X / Screen.TwipsPerPixelX)
    iY = (Y / Screen.TwipsPerPixelY)
    Actual.Cls
    If iX < agraficow And iY < agraficoh Then
        modDXEngine.DibujareEnHwnd3 Actual.hWnd, aGrafico, 0, 0, True
        If Not fTextureMaker.csel.value Then
            If Button = vbLeftButton Then
                For z = 1 To numNewIndex
                    If NewIndexData(z).OverWriteGrafico = aGrafico Then
                        EncontroGrafico = True
                        If Not NewIndexData(z).Estatic = 0 Then
                            With EstaticData(NewIndexData(z).Estatic)
                                If iX >= .L And iX < .L + .W Then
                                    If iY >= .t And iY < .t + .H Then
                                        'Es este el indice.
                                        Actual.ForeColor = vbWhite
                                        Actual.Line ((.L * Screen.TwipsPerPixelX), (.t * Screen.TwipsPerPixelY))-(((.L + .W) * Screen.TwipsPerPixelX), (.t * Screen.TwipsPerPixelY))
                                        Actual.Line ((.L * Screen.TwipsPerPixelX), ((.t + .H) * Screen.TwipsPerPixelY))-(((.L + .W) * Screen.TwipsPerPixelX), ((.t + .H) * Screen.TwipsPerPixelY))
                                        Actual.Line (((.L + .W) * Screen.TwipsPerPixelX), (.t * Screen.TwipsPerPixelY))-(((.L + .W) * Screen.TwipsPerPixelX), ((.t + .H) * Screen.TwipsPerPixelY))
                                        Actual.Line ((.L * Screen.TwipsPerPixelX), (.t * Screen.TwipsPerPixelY))-((.L * Screen.TwipsPerPixelX), ((.t + .H) * Screen.TwipsPerPixelY))
                                        FoundEstatico = True
                                        Exit For
                        
                        
                                    End If
                                End If
            
        
                            End With
                        End If
                    End If
                Next z
                Dim j As Long
                Dim H As Long
                If EncontroGrafico And Not FoundEstatico Then
                    For z = 1 To numNewIndex
                        If NewIndexData(z).OverWriteGrafico = aGrafico Then
                            If NewIndexData(z).Dinamica > 0 Then
                                With NewAnimationData(NewIndexData(z).Dinamica)
                                    For j = 1 To .NumFrames
                                        If iX >= .Indice(j).X And iX < (.Indice(j).X + .Width) Then
                                            If iY >= .Indice(j).Y And iY < .Indice(j).Y + .Height Then
                                                If (.Indice(j).Grafico = aGrafico) Or (.Indice(j).Grafico = 0) Then

                                                    'ENCONTRADO!!!
                                                    'Si lo encontramos seleccionamos los frames q esten en ese grafico y si tocamos
                                                    'el de pasar, pasa solo el frame1.
                                                    For H = 1 To .NumFrames
                                                        AplicarGrilla .Indice(H).X, .Indice(H).Y, .Indice(H).X + .Width, .Indice(H).Y + .Height, fTextureMaker.Actual, vbGreen
                                        
                                                    Next H
                                                    Exit For
                                                End If
                                            End If
                                        End If
                                    Next j
                                    If j < .NumFrames Then Exit For
                                End With
                            End If
                        End If
                    Next z
        
                End If
    
                If tw = 0 Then tw = 32
                If th = 0 Then th = 32
    
    
                If z > numNewIndex Then
                    'No esta indexado.
                    tX = iX \ 32
                    tY = iY \ 32
                    tX = tX * 32
                    tY = tY * 32
    
                    Text2.Text = 0
                    Text7.Text = tX
                    Text8.Text = tY
                    Text9.Text = tw
                    Text10.Text = th
                    setcelda
                    Text11.Text = z
                    cmdIndex.Enabled = True
                    cmdReci.Enabled = True
        
                Else
                    If NewIndexData(z).Estatic > 0 Then
                        Text2.Text = z
                        Text7.Text = EstaticData(NewIndexData(z).Estatic).L
                        Text8.Text = EstaticData(NewIndexData(z).Estatic).t
                        Text9.Text = EstaticData(NewIndexData(z).Estatic).W
                        Text10.Text = EstaticData(NewIndexData(z).Estatic).H
                        Text11.Text = z
                        cmdIndex.Enabled = False
                        cmdReci.Enabled = False
                    ElseIf NewIndexData(z).Dinamica > 0 Then
                        Text2.Text = z
                        cmdIndex.Enabled = False
                        cmdReci.Enabled = False
            
                    End If
                End If
            ElseIf Button = vbRightButton Then

            End If
        Else
            If fSelecter = False Then
                fSelecter = True
                ifX = iX
                ifY = iY
                iiFx = X
                iiFy = Y
                Selecto = False
            Else
                If agraficoh > 0 And agraficow > 0 Then
                    lifX = iX
                    lifY = iY
                    liiFx = X
                    liiFy = Y
                    ProcesarLimites ifX, ifY, lifX, lifY
                    If lifX > ifX And lifY > ifY Then
                        modDXEngine.DibujareEnHwnd3 Actual.hWnd, aGrafico, 0, 0, True
                        AplicarGrilla ifX, ifY, lifX, lifY, Actual, vbBlue
    
                    End If
                    fSelecter = False
                    Selecto = True
    
                End If
            End If



        End If
    End If
End Sub
Public Sub ProcesarLimites(ByVal X As Integer, ByVal Y As Integer, ByVal W As Integer, ByVal H As Integer)
Dim z As Long
Dim iC As Integer
Dim fC As Integer
Dim FF As Integer
Dim inF As Integer
Dim i As Long
Dim j As Long
Dim EncontroIndex As Boolean
Dim k As Long
Dim agw As Integer
agw = agraficow \ 32

iC = (X - 1) \ 32
fC = (W - 1) \ 32
FF = (H - 1) \ 32
inF = (Y - 1) \ 32




For i = iC To fC
    For j = inF To FF
        For z = 1 To numNewIndex
            If NewIndexData(z).OverWriteGrafico = aGrafico Then
                If NewIndexData(z).Estatic > 0 Then
                    With EstaticData(NewIndexData(z).Estatic)
                        If (.L + .W) > (i * 32) And .L < ((i + 1) * 32) Then
                            If (.t + .H) > ((j) * 32) And .t < ((j + 1) * 32) Then
                                'Este grafico esta aca.
                                'Tenemos que reajustaR?
                                
                                If NumSel > 0 Then
                                    For k = 1 To NumSel
                                        If SelInd(k) = z Then Exit For
                                        
                                    Next k
                                Else
                                    k = 1
                                End If
                                If k > NumSel Then
                                    NumSel = k
                                    ReDim Preserve SelInd(1 To k)
                                    ReDim Preserve SelStartX(1 To k)
                                    ReDim Preserve SelStartY(1 To k)
                                    SelInd(k) = z
                                    SelStartX(k) = (.L \ 32)
                                    SelStartY(k) = (.t \ 32)
                                End If
                                    
                                
                                If ifX > .L Or ifX = -1 Then
                                    ifX = .L
                                    iiFx = ifX * Screen.TwipsPerPixelX
                                End If
                                If lifX < (.L + .W) Or lifX = -1 Then
                                    lifX = .L + .W
                                    liiFx = lifX * Screen.TwipsPerPixelX
                                End If
                                
                                If ifY > .t Or ifY = -1 Then
                                    ifY = .t
                                    iiFy = ifY * Screen.TwipsPerPixelY
                                End If
                                If lifY < (.t + .H) Or lifY = -1 Then
                                    lifY = .t + .H
                                    liiFy = lifY * Screen.TwipsPerPixelY
                                End If
                                
                                EncontroIndex = True
                                Exit For
                            
                            End If
                        End If
                    End With
                End If
            End If
        Next z
    Next j
Next i
If EncontroIndex = False Then
    lifY = 0
    liiFy = 0
    lifX = 0
    liiFx = 0
    ifX = 0
    iiFx = 0
    ifY = 0
    iiFy = 0
End If


End Sub


Private Sub butgrilla_Click()
    bGrilla = butgrilla.value
    If bGrilla Then
        If agraficow > 0 And agraficoh > 0 Then
            AplicarGrilla 0, 0, agraficow, agraficoh, Actual, vbGreen
        End If
    Else
        If agraficow > 0 And agraficoh > 0 Then

            modDXEngine.DibujareEnHwnd3 Actual.hWnd, aGrafico, 0, 0, True
        End If
    End If
End Sub
Private Sub IndexarTemporal()
    Dim P As Long
    Dim j As Long
    For P = 1 To NumRealEstatic
        If EstaticData(P).L = Val(Text7.Text) Then
            If EstaticData(P).t = Val(Text8.Text) Then
                If EstaticData(P).W = Val(Text9.Text) Then
                    If EstaticData(P).H = Val(Text10.Text) Then
                        Exit For
                    End If
                End If
            End If
        End If
    Next P
    If P > NumRealEstatic Then
        'No esta indexado la estatica, lo buscamos en las estaticas temporales.
        For j = 1 To ntEstatic
            With TempEstatic(j)
                If .L = Val(Text7.Text) Then
                    If .t = Val(Text8.Text) Then
                        If .W = Val(Text9.Text) Then
                            If .H = Val(Text10.Text) Then
                                Exit For
                            End If
                        End If
                    End If
                End If
            End With
        Next j
        P = j
        If j > ntEstatic Then
            'No esta indexado tampoco en las tempestatic, hay q indexarlo.
            ntEstatic = ntEstatic + 1
            ReDim Preserve TempEstatic(1 To ntEstatic)
            With TempEstatic(ntEstatic)
                .L = Val(Text7.Text)
                .t = Val(Text8.Text)
                .W = Val(Text9.Text)
                .H = Val(Text10.Text)
                .tw = .W / 32
                .th = .H / 32
                WriteVar App.PATH & "\Resources\InitTemp\TempEstatics.Dat", "INIT", "NUM", CStr(ntEstatic)
                WriteVar App.PATH & "\Resources\InitTemp\TempEstatics.Dat", CStr(j), "Left", CStr(.L)
                WriteVar App.PATH & "\Resources\InitTemp\TempEstatics.Dat", CStr(j), "Top", CStr(.t)
                WriteVar App.PATH & "\Resources\InitTemp\TempEstatics.Dat", CStr(j), "Width", CStr(.W)
                WriteVar App.PATH & "\Resources\InitTemp\TempEstatics.Dat", CStr(j), "Height", CStr(.H)
            End With
            numNewEstatic = numNewEstatic + 1
            ReDim Preserve EstaticData(1 To numNewEstatic)
            With EstaticData(numNewEstatic)
                .L = Val(Text7.Text)
                .t = Val(Text8.Text)
                .W = Val(Text9.Text)
                .H = Val(Text10.Text)
                .tw = .W / 32
                .th = .H / 32
            End With
        End If
    End If

    'Tenemos el valor de estatic, que es P. Si P = J es decir que el valor es un tempestatic
    'sino, es un estatic comun. Ahora tenemos q indexar el indice.

    ntIndex = ntIndex + 1

    ReDim Preserve TempIndex(1 To ntIndex)

    If P = j Then
        'TempEstatic
        TempIndex(ntIndex).temp = 1
    Else
        TempIndex(ntIndex).temp = 0
    End If

    TempIndex(ntIndex).Estatic = P
    TempIndex(ntIndex).OverWriteGrafico = Val(Text1.Text)



    WriteVar App.PATH & "\Resources\InitTemp\TempIndex.Dat", "INIT", "NUM", CStr(ntIndex)
    WriteVar App.PATH & "\Resources\InitTemp\TempIndex.Dat", CStr(ntIndex), "Estatica", CStr(P)
    WriteVar App.PATH & "\Resources\InitTemp\TempIndex.Dat", CStr(ntIndex), "OverWriteGrafico", CStr(TempIndex(ntIndex).OverWriteGrafico)
    WriteVar App.PATH & "\Resources\InitTemp\TempIndex.Dat", CStr(ntIndex), "Temp", CStr(TempIndex(ntIndex).temp)



    numNewIndex = numNewIndex + 1
    ReDim Preserve NewIndexData(1 To numNewIndex)
    With NewIndexData(numNewIndex)
        If TempIndex(ntIndex).temp = 1 Then
            .Estatic = NumRealEstatic + P
        Else
            .Estatic = P
        End If
        .OverWriteGrafico = Val(Text1.Text)
    End With

End Sub
Private Sub cmdIndex_Click()
    IndexarTemporal

    Exit Sub
    Dim P As Long
    If Val(Text11.Text) = numNewIndex + 1 Then
        numNewIndex = numNewIndex + 1
        ReDim Preserve NewIndexData(1 To numNewIndex)
        WriteVar App.PATH & "\Resources\Init\NewIndex.dat", "INIT", "NUM", CStr(numNewIndex)
    ElseIf Val(Text11.Text) > numNewIndex + 1 Then
        MsgBox "Estás dejando indices sin usar."
        Exit Sub
    End If


    For P = 1 To numNewEstatic
        If EstaticData(P).L = Val(Text7.Text) Then
            If EstaticData(P).t = Val(Text8.Text) Then
                If EstaticData(P).W = Val(Text9.Text) Then
                    If EstaticData(P).H = Val(Text10.Text) Then
                        Exit For
                    End If
                End If
            End If
        End If
    Next P
    If P > numNewEstatic Then
        'No esta indexado.
        numNewEstatic = numNewEstatic + 1
        EstaticData(P).L = Val(Text7.Text)
        EstaticData(P).t = Val(Text8.Text)
        EstaticData(P).W = Val(Text9.Text)
        EstaticData(P).H = Val(Text10.Text)
        WriteVar App.PATH & "\Resources\Init\NewEstatics.dat", "INIT", "NUM", CStr(numNewEstatic)
        WriteVar App.PATH & "\Resources\Init\NewEstatics.dat", CStr(P), "Left", CStr(EstaticData(P).L)
        WriteVar App.PATH & "\Resources\Init\NewEstatics.dat", CStr(P), "Top", CStr(EstaticData(P).t)
        WriteVar App.PATH & "\Resources\Init\NewEstatics.dat", CStr(P), "Width", CStr(EstaticData(P).W)
        WriteVar App.PATH & "\Resources\Init\NewEstatics.dat", CStr(P), "Heigth", CStr(EstaticData(P).H)
    End If

    NewIndexData(Val(Text11.Text)).Estatic = P
    NewIndexData(Val(Text11.Text)).OverWriteGrafico = Val(Text1.Text)
    
    WriteVar App.PATH & "\Resources\INIT\NewIndex.dat", Text11.Text, "Estatica", CStr(P)
    WriteVar App.PATH & "\Resources\INIT\NewIndex.dat", Text11.Text, "OverWriteGrafico", Val(Text1.Text)
    
    
    Text2.Text = (Text11.Text)
End Sub

Private Sub cmdReci_Click()
    Dim P As Long
    Text11.Text = "..."
    DoEvents
    For P = 5 To numNewIndex
        If UCase$(left$(GetVar(App.PATH & "\Resources\Init\NewIndex.dat", CStr(P), "OverWriteGrafico"), 1)) = "R" Then
            Text11.Text = P
            Exit For
        End If
    Next P
End Sub

Private Sub Combo1_Click()
    Dim P As Integer
    Dim Xn As Integer
    Dim Xx As Integer
    Dim Yn As Integer
    Dim Yx As Integer
    Dim Y As Long
    Dim X As Long
    List1.Clear
    TexAc = Combo1.ListIndex + 1
    tAlto = TexWE(TexAc).Largo
    tAncho = TexWE(TexAc).Ancho
    If TexWE(TexAc).NumIndex > 0 Then
        ReDim tIndex(1 To TexWE(TexAc).NumIndex)
        tNumIndex = TexWE(TexAc).NumIndex
        For P = 1 To TexWE(TexAc).NumIndex
        
            
            tIndex(P).Num = TexWE(TexAc).index(P).Num
            tIndex(P).X = TexWE(TexAc).index(P).X
            tIndex(P).Y = TexWE(TexAc).index(P).Y
            
            Xn = (tIndex(P).X \ 32) + 1
            Yn = (tIndex(P).Y \ 32) + 1
            Xx = Xn + ((EstaticData(NewIndexData(tIndex(P).Num).Estatic).W - 1) \ 32)
            Yx = Yn + ((EstaticData(NewIndexData(tIndex(P).Num).Estatic).H - 1) \ 32)
             
            For X = Xn To Xx
                For Y = Yn To Yx
                    
                    TexArray(((Y - 1) * 16) + (X)) = tIndex(P).Num
                    TexInicialX(((Y - 1) * 16) + (X)) = tIndex(P).X
                    TexInicialY(((Y - 1) * 16) + (X)) = tIndex(P).Y
                    
                    
                Next Y
            Next X
            
            
            List1.AddItem tIndex(P).Num & "[" & tIndex(P).X & "," & tIndex(P).Y & "]"
        Next P
    End If
    Text3 = tAncho
    Text4 = tAlto
    Text14 = TexWE(TexAc).Name
    
    updatenuevo
End Sub

Private Sub Command_Click()
Erase tIndex
tNumIndex = 0
Erase TexInicialX
Erase TexInicialY
Erase TexArray
List1.Clear
Nuevo.Cls

End Sub

Private Sub Command1_Click()
    'Busca grafico y le pone en el Actual.
    If Val(Text1.Text) > 0 Then
        'Esta escrito un grafico.
        aGrafico = Val(Text1.Text)
        If aGrafico < 232 Or aGrafico >= 650 Then
            MsgBox "El grafico debe estar entre 232 y 649."
            aGrafico = 0
            Exit Sub
        End If
    ElseIf Val(Text2.Text) > 0 Then
        'Esta escrito un index
        aGrafico = NewIndexData(Val(Text2.Text)).OverWriteGrafico
    End If
    Actual.Cls
    modDXEngine.DibujareEnHwnd3 Actual.hWnd, aGrafico, 0, 0, True
    agraficow = DameWidthTextura(aGrafico)
    agraficoh = DameHeightTextura(aGrafico)
    
End Sub

Private Sub Command2_Click()
    Dim Tmp As Integer
    Dim P As Long
    Dim X As Long
    Dim Y As Long
    Dim lx As Long

    Dim k As Integer
    Dim kTmp As Integer
    Tmp = (nfy * 16) + (nfx + 1)
    If Selecto = False Then
        If TexArray(Tmp) = 0 Then

            'Agregamos
            If Val(Text2.Text) > 0 And Val(Text2.Text) < numNewIndex Then
                If NewIndexData(Val(Text2.Text)).Estatic > 0 Then
    
                    With EstaticData(NewIndexData(Val(Text2.Text)).Estatic)
        
                        For Y = nfy To nfy + ((.H \ 32) - 1)
                
                            For X = nfx To nfx + ((.W \ 32) - 1)
                                Tmp = (Y * 16) + (X + 1)
                    
                                TexArray(Tmp) = Val(Text2.Text)
                    
                                TexInicialX(Tmp) = nfx * 32
                                TexInicialY(Tmp) = nfy * 32
                            Next X
                        Next Y
            
                        If ((nfx + ((.W \ 32))) * 32) > tAncho Then tAncho = ((nfx + ((.W \ 32))) * 32)
                        If ((nfy + ((.H \ 32))) * 32) > tAlto Then tAlto = ((nfy + ((.H \ 32))) * 32)
            
                        tNumIndex = tNumIndex + 1
                        ReDim Preserve tIndex(1 To tNumIndex)
                        tIndex(tNumIndex).Num = Val(Text2.Text)
                        tIndex(tNumIndex).X = nfx * 32
                        tIndex(tNumIndex).Y = nfy * 32
            
            
                    End With
                ElseIf NewIndexData(Val(Text2.Text)).Dinamica > 0 Then
                    With NewAnimationData(NewIndexData(Val(Text2.Text)).Dinamica)
                        For Y = nfy To nfy + ((.Height \ 32) - 1)
                
                            For X = nfx To nfx + ((.Width \ 32) - 1)
                                Tmp = (Y * 16) + (X + 1)
                                TexArray(Tmp) = Val(Text2.Text)
                                TexInicialX(Tmp) = nfx * 32
                                TexInicialY(Tmp) = nfy * 32
                            Next X
                        Next Y
                        If ((nfx + ((.Width \ 32))) * 32) > tAncho Then tAncho = ((nfx + ((.Width \ 32))) * 32)
                        If ((nfy + ((.Height \ 32))) * 32) > tAlto Then tAlto = ((nfy + ((.Height \ 32))) * 32)
                
                        tNumIndex = tNumIndex + 1
                        ReDim Preserve tIndex(1 To tNumIndex)
                        tIndex(tNumIndex).Num = Val(Text2.Text)
                        tIndex(tNumIndex).X = nfx * 32
                        tIndex(tNumIndex).Y = nfy * 32
                
                
                    End With
                End If
        
        
        
                updatenuevo
        
            End If



            Text3.Text = tAncho
            Text4.Text = tAlto
            Label7.Caption = "Num Index: " & tNumIndex
            List1.AddItem tNumIndex & " - " & Val(Text2.Text) & "[" & tIndex(tNumIndex).X & "," & tIndex(tNumIndex).Y & "]"
        End If

    Else
        Dim TmpS As Integer
        If NumSel > 0 Then
            'Check capacity.
            For k = 1 To NumSel
                TmpS = ((nfx + SelStartX(k)) * 32) + EstaticData(NewIndexData(SelInd(k)).Estatic).W
                If tAncho < TmpS Then tAncho = TmpS
                TmpS = ((nfy + SelStartY(k)) * 32) + EstaticData(NewIndexData(SelInd(k)).Estatic).H
                If tAlto < TmpS Then tAlto = TmpS
            Next k
    
            For k = 1 To NumSel
                Tmp = ((nfy + SelStartY(k)) * 16) + (nfx + SelStartX(k) + 1)
                If TexArray(Tmp) = 0 Then

                    'Agregamos
                    If SelInd(k) > 0 And SelInd(k) < numNewIndex Then
                        With EstaticData(NewIndexData(SelInd(k)).Estatic)
        
                            For Y = (nfy + SelStartY(k)) To (nfy + SelStartY(k)) + ((.H \ 32) - 1)
                
                                For X = (nfx + SelStartX(k)) To (nfx + SelStartX(k)) + ((.W \ 32) - 1)
                                    Tmp = (Y * 16) + (X + 1)
                    
                                    TexArray(Tmp) = SelInd(k)
                    
                                    TexInicialX(Tmp) = (nfx + SelStartX(k)) * 32
                                    TexInicialY(Tmp) = (nfy + SelStartY(k)) * 32
                                Next X
                            Next Y
            
                            tNumIndex = tNumIndex + 1
                            ReDim Preserve tIndex(1 To tNumIndex)
                            tIndex(tNumIndex).Num = SelInd(k)
                            tIndex(tNumIndex).X = (nfx + SelStartX(k)) * 32
                            tIndex(tNumIndex).Y = (nfy + SelStartY(k)) * 32
            
            
                        End With
                        updatenuevo
                    End If



                    Text3.Text = tAncho
                    Text4.Text = tAlto
                    Label7.Caption = "Num Index: " & tNumIndex
                    List1.AddItem tNumIndex & " - " & SelInd(k) & "[" & tIndex(tNumIndex).X & "," & tIndex(tNumIndex).Y & "]"
                End If
            Next k
        End If
    End If
End Sub
Public Sub updatenuevo()
    Nuevo.Cls
    If tAlto = 0 Or tAncho = 0 Then Exit Sub
    Dim P As Long
    Dim R As RECT
    Dim d As RECT
    If tNumIndex = 0 Then Exit Sub
    R.Bottom = tAlto
    R.Right = tAncho
    ddevice.Clear 1, R, D3DCLEAR_TARGET, &H0, ByVal 0, 0
    For P = 1 To tNumIndex
    
        modDXEngine.DibujareEnHwnd2 Nuevo.hWnd, tIndex(P).Num, R, tIndex(P).X, tIndex(P).Y, False


    Next P


    d.left = 0
    d.top = 0
    d.Bottom = tAlto
    d.Right = tAncho
    ddevice.Present R, d, Nuevo.hWnd, ByVal 0
End Sub

Private Sub Command3_Click()
    Dim P As Long
    If TexAc = 0 Or tNumIndex = 0 Then Exit Sub
    With TexWE(TexAc)
        .Ancho = tAncho
        .Largo = tAlto
        .Name = Text14.Text
        .NumIndex = tNumIndex
        ReDim .index(1 To tNumIndex)
        
        For P = 1 To tNumIndex
            .index(P).Num = tIndex(P).Num
            .index(P).X = tIndex(P).X
            .index(P).Y = tIndex(P).Y
        Next P
    End With
    Combo1.List(TexAc - 1) = Text14 & " [" & TexAc & "]"
    frmMain.lListado(0).Clear
    For P = 1 To NumTexWe
        frmMain.lListado(0).AddItem TexWE(P).Name & " - [" & P & "]"
    Next P

    GuardoTexturita = True
    GuardarTex False, TexAc
End Sub

Private Sub Command4_Click()
    If GuardoTexturita Then
        NumTexWe = NumTexWe + 1
        ReDim Preserve TexWE(1 To NumTexWe)
    End If
    TexAc = NumTexWe
    Combo1.AddItem Text14.Text & " [" & NumTexWe & "]"
    List1.Clear
    Text3.Text = 0
    Text4.Text = 0

    Erase TexArray
    Erase TexInicialX
    Erase TexInicialY
    tNumIndex = 0
    Erase tIndex

    GuardoTexturita = False


End Sub

Private Sub Command5_Click()
    Dim Tmp As Long
    Dim W As Integer
    Dim H As Integer
    Dim P As Long
    Dim j As Long
    Dim X As Long
    Dim Y As Long
    Tmp = (nfy * 16) + (nfx + 1)
    If TexArray(Tmp) > 0 Then
        nfx = TexInicialX(Tmp) \ 32
        nfy = TexInicialY(Tmp) \ 32
        H = EstaticData(NewIndexData(TexArray(Tmp)).Estatic).H
        W = EstaticData(NewIndexData(TexArray(Tmp)).Estatic).W
    Else
        Exit Sub
    End If
    For P = 1 To tNumIndex
        If tIndex(P).Num = TexArray(Tmp) Then
            If tIndex(P).X = TexInicialX(Tmp) Then
                If tIndex(P).Y = TexInicialY(Tmp) Then
                    'FOUND IT
                    For j = P To tNumIndex - 1
              
                        tIndex(j) = tIndex(j + 1)
              
                    Next j
                    tNumIndex = tNumIndex - 1
              
              
                    For X = nfx To (nfx - 1) + (((W - 1) \ 32) + 1)
                        For Y = nfy To (nfy - 1) + (((H - 1) \ 32) + 1)
                    
                            TexArray((Y * 16) + (X + 1)) = 0
                            TexInicialY((Y * 16) + (X + 1)) = 0
                            TexInicialX((Y * 16) + (X + 1)) = 0
                    
                
                        Next Y
                    Next X
              
                    'Buscamos por el MAS ALTO Y ANCHO
                    tAncho = 0
                    tAlto = 0
                    For j = 1 To tNumIndex
                        If (tIndex(j).X + EstaticData(NewIndexData(tIndex(j).Num).Estatic).W) > tAncho Then
                            tAncho = (tIndex(j).X + EstaticData(NewIndexData(tIndex(j).Num).Estatic).W)
                        End If
                        If (tIndex(j).Y + EstaticData(NewIndexData(tIndex(j).Num).Estatic).H) > tAlto Then
                            tAlto = (tIndex(j).Y + EstaticData(NewIndexData(tIndex(j).Num).Estatic).H)
                        End If
                    Next j
                    Exit For
                End If
            End If
        End If
    Next P


    List1.Clear
    For P = 1 To tNumIndex
        List1.AddItem tIndex(P).Num & "[" & tIndex(P).X & "," & tIndex(P).Y & "]"

    Next P
    Text4 = tAlto
    Text3 = tAncho
    updatenuevo
End Sub

Private Sub Form_Load()
    fSelecter = False
    Selecto = False
    GuardoTexturita = True
    
End Sub
Private Sub AplicarGrilla(ByVal X As Integer, ByVal Y As Integer, ByVal W As Integer, ByVal H As Integer, ByRef P As PictureBox, ByVal Color As Long)
    Dim i As Long
    Dim j As Long
    Dim iX As Long
    Dim iY As Long
    Dim Fila As Integer
    Dim Col As Integer
    Dim iW As Long
    Dim iH As Long
    Dim fFila As Integer
    Dim fCol As Integer
    X = X * Screen.TwipsPerPixelX
    Y = Y * Screen.TwipsPerPixelY
    W = W * Screen.TwipsPerPixelX
    H = H * Screen.TwipsPerPixelY



    Col = (X) \ (32 * Screen.TwipsPerPixelX)
    Fila = (Y) \ (32 * Screen.TwipsPerPixelY)

    iX = Col * (32 * Screen.TwipsPerPixelX)
    iY = Fila * (32 * Screen.TwipsPerPixelY)

    fCol = (W - (1 * Screen.TwipsPerPixelX)) \ (32 * Screen.TwipsPerPixelX)
    fFila = (H - (1 * Screen.TwipsPerPixelY)) \ (32 * Screen.TwipsPerPixelY)
    iW = fCol * (32 * Screen.TwipsPerPixelX)
    iH = fFila * (32 * Screen.TwipsPerPixelY)
    P.ForeColor = Color
    For i = Col To fCol
    
        For j = Fila To fFila
            
            P.Line ((i * (32 * Screen.TwipsPerPixelX)), (j * (32 * Screen.TwipsPerPixelY)))-((i * (32 * Screen.TwipsPerPixelX)) + (32 * Screen.TwipsPerPixelX), (j * (32 * Screen.TwipsPerPixelY)))
            P.Line ((i * (32 * Screen.TwipsPerPixelX)), (j * (32 * Screen.TwipsPerPixelY)))-((i * (32 * Screen.TwipsPerPixelX)), (j * (32 * Screen.TwipsPerPixelY)) + (32 * Screen.TwipsPerPixelY))
            P.Line ((i * (32 * Screen.TwipsPerPixelX)), (j * (32 * Screen.TwipsPerPixelY)) + (32 * Screen.TwipsPerPixelY))-((i * (32 * Screen.TwipsPerPixelX)) + (32 * Screen.TwipsPerPixelX), (j * (32 * Screen.TwipsPerPixelY)) + (32 * Screen.TwipsPerPixelY))
            P.Line ((i * (32 * Screen.TwipsPerPixelX)) + (32 * Screen.TwipsPerPixelX), (j * (32 * Screen.TwipsPerPixelY)))-((i * (32 * Screen.TwipsPerPixelX)) + (32 * Screen.TwipsPerPixelX), (j * (32 * Screen.TwipsPerPixelY)) + (32 * Screen.TwipsPerPixelY))
            
        Next j
    Next i


End Sub

Private Sub lvButtons_H1_Click()
    bGrilla = butgrilla.value

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If GuardoTexturita = False Then
        NumTexWe = NumTexWe - 1
        ReDim Preserve TexWE(1 To NumTexWe)
    End If
End Sub

Private Sub Nuevo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Tmp As Integer
    Dim W As Integer
    Dim H As Integer
    Nuevo.Cls
    updatenuevo
    atx = X / Screen.TwipsPerPixelX
    aty = Y / Screen.TwipsPerPixelY

    nfx = atx \ 32
    nfy = aty \ 32

    Tmp = (nfy * 16) + (nfx + 1)

    If TexArray(Tmp) > 0 Then
        nfx = TexInicialX(Tmp) \ 32
        nfy = TexInicialY(Tmp) \ 32
        H = EstaticData(NewIndexData(TexArray(Tmp)).Estatic).H
        W = EstaticData(NewIndexData(TexArray(Tmp)).Estatic).W
    Else
        H = 32
        W = 32
    End If

    Text12.Text = nfx * 32
    Text13.Text = nfy * 32

    Nuevo.Line (((nfx * 32) * Screen.TwipsPerPixelX), ((nfy * 32) * Screen.TwipsPerPixelY))-((((nfx * 32) + W) * Screen.TwipsPerPixelX), ((nfy * 32) * Screen.TwipsPerPixelY))
    Nuevo.Line (((nfx * 32) * Screen.TwipsPerPixelX), (((nfy * 32) + H) * Screen.TwipsPerPixelY))-((((nfx * 32) + W) * Screen.TwipsPerPixelX), (((nfy * 32) + H) * Screen.TwipsPerPixelY))
    Nuevo.Line (((nfx * 32) * Screen.TwipsPerPixelX), (((nfy * 32)) * Screen.TwipsPerPixelY))-((((nfx) * 32) * Screen.TwipsPerPixelX), (((nfy * 32) + H) * Screen.TwipsPerPixelY))
    Nuevo.Line ((((nfx * 32) + W) * Screen.TwipsPerPixelX), ((nfy * 32) * Screen.TwipsPerPixelY))-((((nfx * 32) + W) * Screen.TwipsPerPixelX), (((nfy * 32) + H) * Screen.TwipsPerPixelY))

End Sub

Private Sub Nuevo_Paint()
    updatenuevo

End Sub

Private Sub Text10_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Actual.Cls
        modDXEngine.DibujareEnHwnd3 Actual.hWnd, aGrafico, 0, 0, True
    
        setcelda
    End If
End Sub

Private Sub Text7_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Actual.Cls
        modDXEngine.DibujareEnHwnd3 Actual.hWnd, aGrafico, 0, 0, True
        setcelda
    End If
End Sub
Sub setcelda()

    th = Val(Text10.Text)
    tw = Val(Text9.Text)
    tX = Val(Text7.Text)
    tY = Val(Text8.Text)
        
    Actual.ForeColor = vbRed
    Actual.Line (tX * Screen.TwipsPerPixelX, tY * Screen.TwipsPerPixelY)-((tX + tw) * Screen.TwipsPerPixelX, tY * Screen.TwipsPerPixelY)
    Actual.Line (tX * Screen.TwipsPerPixelX, (tY + th) * Screen.TwipsPerPixelY)-((tX + tw) * Screen.TwipsPerPixelX, (tY + th) * Screen.TwipsPerPixelY)
    Actual.Line (tX * Screen.TwipsPerPixelX, tY * Screen.TwipsPerPixelY)-(tX * Screen.TwipsPerPixelX, (tY + th) * Screen.TwipsPerPixelY)
    Actual.Line ((tX + tw) * Screen.TwipsPerPixelX, tY * Screen.TwipsPerPixelY)-((tX + tw) * Screen.TwipsPerPixelX, (tY + th) * Screen.TwipsPerPixelY)

End Sub

Private Sub Text8_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Actual.Cls
        modDXEngine.DibujareEnHwnd3 Actual.hWnd, aGrafico, 0, 0, True
    
        setcelda
    End If
End Sub

Private Sub Text9_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Actual.Cls
        modDXEngine.DibujareEnHwnd3 Actual.hWnd, aGrafico, 0, 0, True
    
        setcelda
    End If
End Sub
