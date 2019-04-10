VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "MapHeader"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6375
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Guardar"
      Height          =   360
      Left            =   2760
      TabIndex        =   7
      Top             =   2520
      Width           =   990
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2640
      TabIndex        =   5
      Text            =   "0"
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox tSom 
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Text            =   "1"
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox tVer 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Text            =   "2"
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Si esta en 0, automaticamente el cliente va a ponerle 21 + MapNumber."
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   1320
      Width           =   5160
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grafico Minimapa:"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   1995
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sombras:"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   705
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    MapHead.GraficoMapa = Val(Text1)
    MapHead.Version = Val(tVer)
    MapHead.SombrasAmbientales = Val(tSom)
End Sub

