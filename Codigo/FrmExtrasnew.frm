VERSION 5.00
Begin VB.Form FrmExtras 
   Caption         =   "Extras"
   ClientHeight    =   5640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8025
   LinkTopic       =   "Form2"
   ScaleHeight     =   5640
   ScaleWidth      =   8025
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "General"
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
      Begin VB.CommandButton Command1 
         Caption         =   "Actualizar Temporal"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "Revisa el mapa para saber si es un mapa temporal o no."
         Top             =   240
         Width           =   2295
      End
   End
End
Attribute VB_Name = "FrmExtras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    MapaTemporal = ChequearTemporal
    If MapaTemporal Then
        MsgBox "Este mapa es un mapa temporal."
    Else
        MsgBox "Este mapa no es temporal."
    End If
End Sub
