VERSION 5.00
Begin VB.Form Render 
   Caption         =   "Render"
   ClientHeight    =   2235
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8745
   LinkTopic       =   "Form2"
   ScaleHeight     =   149
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   583
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   360
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   89
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   6720
      TabIndex        =   7
      Text            =   "100"
      Top             =   480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   6720
      TabIndex        =   6
      Text            =   "100"
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generar "
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   1560
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.PictureBox ProgressBar1 
      Height          =   615
      Left            =   1200
      ScaleHeight     =   555
      ScaleWidth      =   6315
      TabIndex        =   2
      Top             =   840
      Width           =   6375
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9735
      Left            =   240
      ScaleHeight     =   649
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   833
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   12495
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   48000
      Left            =   -120
      ScaleHeight     =   3200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   3200
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   48000
   End
   Begin VB.Label Label2 
      Caption         =   "Alto:"
      Height          =   255
      Left            =   5880
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Ancho:"
      Height          =   255
      Left            =   5880
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
End
Attribute VB_Name = "Render"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Render(Optional ByVal MiniMapa As Boolean)
Dim nombre As String
    Set s = ddevice.GetBackBuffer(0, D3DBACKBUFFER_TYPE_MONO)
    
    Dim sr As RECT
    
    Dim i As Long
    Dim z As Long
    Dim LastX As Long
    Dim LastY As Long
    Dim C As Long
    sr.Bottom = 256
    sr.Right = 256
    Picture1.Cls
    Me.Caption = "Produciendo...."
    guardobmp = True
        Picture1.AutoSize = True
    Picture1.AutoRedraw = True
    Picture3.AutoRedraw = True
    If Not MiniMapa Then
    nombre = GetVar(App.PATH & "\Render\Render.txt", "INIT", "NUM")
    Else
    nombre = GetVar(App.PATH & "\Render\Render.txt", "INIT", "NUM_MINIMAPA")
    End If
    For z = 1 To 13
    For i = 1 To 13

        If z = 1 Then
        sr.top = 32
        Else
        sr.top = 0
        End If
        
        If i = 1 Then
        sr.left = 32
        Else
        sr.left = 0
        End If
        ddevice.Clear 0, ByVal 0&, D3DCLEAR_TARGET, D3DColorXRGB(0, 0, 0), 0, 0
        If TIPOMAPAX = 0 Then
        RenderScreen ((8 * (i - 1))) - 1, (8 * (z - 1)) - 1, 0, 0
        ElseIf TIPOMAPAX = 1 Then
        RenderNewMap ((8 * (i - 1))) - 1, (8 * (z - 1)) - 1, 0, 0
        End If
        
    Dim t As PALETTEENTRY
    
    
    d3dx.SaveSurfaceToFile App.PATH & "\Render\" & i & "_" & z & ".bmp", D3DXIFF_BMP, s, t, sr

    Picture2.Cls
    Picture2.AutoSize = True
    
    Set Picture2.Picture = LoadPicture(App.PATH & "\Render\" & i & "_" & z & ".bmp")
    

    Picture2.AutoSize = True
    
    Picture1.PaintPicture Picture2.Image, _
    LastX, LastY, Picture2.ScaleWidth, Picture2.ScaleHeight, _
    0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight
    
        LastX = LastX + Picture2.ScaleWidth
    Kill App.PATH & "\Render\" & i & "_" & z & ".bmp"
    Next i

        Me.Caption = "Produciendo... " & z & "/13"
        'Me.ProgressBar1.value = Me.ProgressBar1.value + 1
        Me.Refresh
        LastY = LastY + Picture2.ScaleHeight
        LastX = 0
    
    Next z

    If Not MiniMapa Then
    
        SavePicture Picture1.Image, App.PATH & "\Render\" & nombre & ".bmp"
        MsgBox "Imagen guardada como " & App.PATH & "\Render\" & nombre & ".bmp"
        WriteVar App.PATH & "\Render\Render.txt", "INIT", "NUM", Val(nombre) + 1
    Else
        Picture3.Width = DameSize(Val(Me.Text1.Text))
        Picture3.Height = DameSize(Val(Me.Text2.Text))
        Picture3.ClipControls = True
        Picture3.PaintPicture Picture1.Image, 0, 0, Val(Me.Text1.Text), Val(Me.Text2.Text) ', 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight
        
        SavePicture Picture3.Image, App.PATH & "\Render\MiniMapas\MiniMapa" & nombre & ".bmp"
        MsgBox "Imagen guardada como " & App.PATH & "\Render\MiniMapas\MiniMapa" & nombre & ".bmp"
        WriteVar App.PATH & "\Render\Render.txt", "INIT", "NUM_MINIMAPA", Val(nombre) + 1
    End If
    guardobmp = False

    Unload Me
End Sub
Public Function DameSize(ByVal nv As Long) As Long
If nv <= 32 Then
    DameSize = 32
ElseIf nv <= 64 Then
    DameSize = 64
ElseIf nv <= 128 Then
    DameSize = 128
ElseIf nv <= 256 Then
    DameSize = 256
ElseIf nv <= 512 Then
    DameSize = 512
Else
    DameSize = nv
End If


End Function
Private Sub Command1_Click()
    Render True
End Sub
