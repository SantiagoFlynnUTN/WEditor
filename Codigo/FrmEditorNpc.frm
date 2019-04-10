VERSION 5.00
Begin VB.Form FrmEditorNpc 
   Caption         =   "Editor NPC"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6000
   LinkTopic       =   "Form2"
   ScaleHeight     =   5865
   ScaleWidth      =   6000
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Respawn"
      Height          =   975
      Left            =   120
      TabIndex        =   32
      Top             =   4200
      Width           =   5775
      Begin VB.ComboBox mSamePos 
         Height          =   315
         ItemData        =   "FrmEditorNpc.frx":0000
         Left            =   1320
         List            =   "FrmEditorNpc.frx":000A
         TabIndex        =   38
         Top             =   600
         Width           =   855
      End
      Begin VB.ComboBox mRespawn 
         Height          =   315
         ItemData        =   "FrmEditorNpc.frx":0016
         Left            =   1200
         List            =   "FrmEditorNpc.frx":0020
         TabIndex        =   37
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox xRespawnTime 
         Height          =   495
         Left            =   4440
         TabIndex        =   36
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label16 
         Caption         =   "Respawn Time:"
         Height          =   255
         Left            =   3240
         TabIndex        =   35
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label15 
         Caption         =   "Misma Posicion:"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label14 
         Caption         =   "Respawnea:"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Stats"
      Height          =   2775
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   5775
      Begin VB.CommandButton cStats 
         Caption         =   "Ver stats"
         Height          =   255
         Left            =   1200
         TabIndex        =   31
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox xNivel 
         Height          =   405
         Left            =   600
         TabIndex        =   20
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lExp 
         Caption         =   "0"
         Height          =   255
         Left            =   480
         TabIndex        =   30
         Top             =   2040
         Width           =   2775
         WordWrap        =   -1  'True
      End
      Begin VB.Label lOro 
         Caption         =   "0"
         Height          =   255
         Left            =   480
         TabIndex        =   29
         Top             =   1680
         Width           =   2775
         WordWrap        =   -1  'True
      End
      Begin VB.Label lVida 
         Caption         =   "0"
         Height          =   255
         Left            =   600
         TabIndex        =   28
         Top             =   600
         Width           =   2655
         WordWrap        =   -1  'True
      End
      Begin VB.Label lDef 
         Caption         =   "0"
         Height          =   255
         Left            =   840
         TabIndex        =   27
         Top             =   960
         Width           =   2415
         WordWrap        =   -1  'True
      End
      Begin VB.Label lDaño 
         Caption         =   "0"
         Height          =   255
         Left            =   600
         TabIndex        =   26
         Top             =   1320
         Width           =   2775
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label12 
         Caption         =   "Exp:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "Oro:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Daño:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Defensa:"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Vida:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label13 
         Caption         =   "Nivel:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lModExp 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   4800
         TabIndex        =   18
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label lModGold 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   4800
         TabIndex        =   17
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label lModHp 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   4800
         TabIndex        =   16
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lModDef 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   4800
         TabIndex        =   15
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lModDaño 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   4800
         TabIndex        =   14
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Modificador Exp:"
         Height          =   255
         Left            =   3480
         TabIndex        =   13
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Modificador Oro:"
         Height          =   255
         Left            =   3480
         TabIndex        =   12
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Modificador Def:"
         Height          =   255
         Left            =   3480
         TabIndex        =   11
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Modificador HP:"
         Height          =   255
         Left            =   3480
         TabIndex        =   10
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Modificador Daño:"
         Height          =   255
         Left            =   3480
         TabIndex        =   9
         Top             =   1320
         Width           =   1455
      End
   End
   Begin VB.CommandButton csave 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   5400
      Width           =   2295
   End
   Begin VB.CommandButton cexit 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   5400
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Caption         =   "General"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      Begin VB.ComboBox bHeading 
         Height          =   315
         ItemData        =   "FrmEditorNpc.frx":002C
         Left            =   960
         List            =   "FrmEditorNpc.frx":003C
         TabIndex        =   7
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox xNum 
         Height          =   405
         Left            =   1560
         TabIndex        =   2
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Heading:"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lNombre 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2160
         TabIndex        =   3
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label Label1 
         Caption         =   "Numero de NPC:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
   End
End
Attribute VB_Name = "FrmEditorNpc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Parse()
    
    If TipoSeleccionado = 2 Then
        With MapData(ObjetoSeleccionado.X, ObjetoSeleccionado.Y)
            xNum.Text = .NPCIndex
            lNombre.Caption = NpcData(.NPCIndex).Name
            bHeading.ListIndex = NpcData(.NPCIndex).Heading - 1
            xNivel.Text = .NpcInfo.Nivel
            If xNivel.Text = 0 Then
                xNivel.Text = 1
                .NpcInfo.Nivel = 1
            End If
            ShowStats
    
        
        
            mRespawn.ListIndex = .NpcInfo.Respawn
            mSamePos.ListIndex = .NpcInfo.RespawnSamePos
            xRespawnTime.Text = .NpcInfo.RespawnTime

        End With
    End If
End Sub

Public Sub ShowStats()

    If TipoSeleccionado = 2 Then
        With NpcData(MapData(ObjetoSeleccionado.X, ObjetoSeleccionado.Y).NPCIndex)
            If Val(xNivel) <= 0 Then xNivel.Text = "1"
            lVida.Caption = (.MaxHP + (.ModHp * (Val(xNivel) - 1))) & " - (" & .MaxHP & " + " & (.ModHp * (Val(xNivel) - 1)) & ")"
            lDaño.Caption = (.MinHit + (.ModHit * (Val(xNivel) - 1))) & "/" & (.MaxHit + (.ModHit * (Val(xNivel) - 1))) & " - (" & .MinHit & "/" & .MaxHit & " + " & (.ModHp * (Val(xNivel) - 1)) & ")"
            lDef.Caption = .Def + (.ModDef * (Val(xNivel) - 1)) & " - (" & .Def & " + " & (.ModDef * (Val(xNivel) - 1)) & ")"
            lOro.Caption = .Oro + (.ModOro * (Val(xNivel) - 1)) & " - (" & .Oro & " + " & (.ModOro * (Val(xNivel) - 1)) & ")"
            lExp.Caption = .Exp + (.ModExp * (Val(xNivel) - 1)) & " - (" & .Exp & " + " & (.ModExp * (Val(xNivel) - 1)) & ")"
            
            lModGold.Caption = .ModOro
            lModExp.Caption = .ModExp
            lModHp.Caption = .ModHp
            lModDaño.Caption = .ModHit
            lModDef.Caption = .ModDef

        End With
    End If

End Sub

Private Sub cexit_Click()
    Unload Me
    
End Sub

Private Sub csave_Click()
    If TipoSeleccionado = 2 Then
    
        With MapData(ObjetoSeleccionado.X, ObjetoSeleccionado.Y)
            
            If Val(xNum) <> .NPCIndex Then
                If Val(xNum) <= UBound(NpcData) And Val(xNum) > 0 Then
                    .NPCIndex = Val(xNum)
                    lNombre.Caption = NpcData(.NPCIndex).Name
                    MakeChar .CHarIndex, NpcData(.NPCIndex).Body, NpcData(.NPCIndex).Head, NpcData(.NPCIndex).Heading, CInt(ObjetoSeleccionado.X), CInt(ObjetoSeleccionado.Y)
                    ShowStats
                End If
            End If
            .NpcInfo.Heading = bHeading.ListIndex + 1
            CharList(.CHarIndex).Heading = bHeading.ListIndex + 1
            If .NpcInfo.Nivel <> Val(xNivel) Then
                .NpcInfo.Nivel = Val(xNivel)
                ShowStats
            End If
            .NpcInfo.Respawn = mRespawn.ListIndex
            .NpcInfo.RespawnSamePos = mSamePos.ListIndex
            .NpcInfo.RespawnTime = Val(xRespawnTime.Text)
            
            
            
        End With
    End If
    
End Sub

Private Sub cStats_Click()
    ShowStats
End Sub

