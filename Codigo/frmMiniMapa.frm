VERSION 5.00
Begin VB.Form frmMiniMapa 
   Caption         =   "Mini MAPA"
   ClientHeight    =   10950
   ClientLeft      =   -4260
   ClientTop       =   -1740
   ClientWidth     =   17190
   LinkTopic       =   "Form1"
   ScaleHeight     =   730
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1146
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3840
      Left            =   0
      ScaleHeight     =   3840
      ScaleWidth      =   3840
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   3840
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Generar Mapa"
      Height          =   255
      Left            =   15360
      TabIndex        =   2
      Top             =   2640
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   48000
      Left            =   13560
      ScaleHeight     =   3200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   3200
      TabIndex        =   1
      Top             =   360
      Width           =   48000
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   12120
      Left            =   120
      ScaleHeight     =   808
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   800
      TabIndex        =   0
      Top             =   0
      Width           =   12000
   End
End
Attribute VB_Name = "frmMiniMapa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command2_Click()

    Set s = ddevice.GetBackBuffer(0, D3DBACKBUFFER_TYPE_MONO)
    
    Dim sr As RECT
    
    Dim i As Long
    Dim z As Long


    sr.Bottom = 256
    sr.Right = 256
    Picture2.Cls
    frmMiniMapa.Caption = "Produciendo...."
    guardobmp = True
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
        RenderScreen ((8 * (i - 1))) - 1, (8 * (z - 1)) - 1, 0, 0
        
        
        
    Dim t As PALETTEENTRY
    
    
    d3dx.SaveSurfaceToFile App.PATH & "\Render\" & z & "_" & i & ".bmp", D3DXIFF_BMP, s, t, sr

    Picture3.Cls
    Set Picture3.Picture = LoadPicture(App.PATH & "\Render\" & i & "_" & z & ".bmp")
    
    Picture3.ScaleMode = 3
    Picture3.AutoSize = True
    Picture2.AutoSize = True
    Picture1.Cls
    Picture1.PaintPicture Picture3.Image, _
    Picture1.ScaleLeft, Picture1.ScaleTop, Picture1.ScaleWidth, Picture1.ScaleHeight, _
    Picture3.ScaleLeft, Picture3.ScaleTop, Picture3.ScaleWidth, Picture3.ScaleHeight
    Picture1.Picture = Picture1.Image
    
    Picture2.PaintPicture Picture1.Image, (Picture1.ScaleWidth) * (i - 1), (Picture1.ScaleHeight) * (z - 1)
    
    Next i
    Next z

    
    SavePicture Picture2.Image, App.PATH & "\" & Nombre_Mapa & "_mini.bmp"
    frmMiniMapa.Caption = "Listo...."
    guardobmp = False
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyU Then
    Stop
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyU Then
    Picture3.Visible = Not Picture3.Visible
End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = vbRightButton Then
        Picture3.Visible = Not Picture3.Visible
    End If
End Sub

