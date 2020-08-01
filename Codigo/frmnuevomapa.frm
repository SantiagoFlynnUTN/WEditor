VERSION 5.00
Begin VB.Form NuevoMapa 
   Caption         =   "Nuevo Mapa"
   ClientHeight    =   1335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6885
   LinkTopic       =   "Form2"
   ScaleHeight     =   1335
   ScaleWidth      =   6885
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Crear"
      Height          =   495
      Left            =   2640
      TabIndex        =   2
      Top             =   840
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Top             =   90
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Textura de base:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "NuevoMapa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Combo1.ListIndex <= 0 Then

NuevoTex = 0

Else

NuevoTex = Val(ReadField(2, Combo1.List(Combo1.ListIndex), Asc("[")))
End If
NuevoOk = True
Unload Me

End Sub

Private Sub Form_Load()
Dim P As Long
    Combo1.AddItem "Ninguna"
    For P = 1 To NumTexWe
        Combo1.AddItem TexWE(P).Name & " - [" & P & "]"
    Next P
End Sub
