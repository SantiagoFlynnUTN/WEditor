VERSION 5.00
Begin VB.Form frmEditorDecor 
   Caption         =   "Editor de decor"
   ClientHeight    =   3585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   239
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "General"
      Height          =   855
      Left            =   240
      TabIndex        =   8
      Top             =   240
      Width           =   4215
      Begin VB.TextBox xDefault 
         Height          =   285
         Left            =   3480
         TabIndex        =   13
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox xTipo 
         Height          =   285
         Left            =   1200
         TabIndex        =   10
         Top             =   220
         Width           =   615
      End
      Begin VB.TextBox xNumero 
         Height          =   285
         Left            =   1440
         TabIndex        =   9
         Top             =   510
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Estado default:"
         Height          =   255
         Left            =   2400
         TabIndex        =   14
         Top             =   255
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Tipo de decor:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Numero de decor:"
         Height          =   165
         Left            =   120
         TabIndex        =   11
         Top             =   570
         Width           =   1455
      End
   End
   Begin VB.CommandButton bExit 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CommandButton bSave 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Clave"
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   2160
      Width           =   4215
      Begin VB.TextBox xClave 
         Height          =   285
         Left            =   3480
         TabIndex        =   15
         Top             =   120
         Width           =   615
      End
      Begin VB.CommandButton bNextClave 
         Caption         =   "Siguiente clave"
         Height          =   320
         Left            =   2160
         TabIndex        =   5
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox xNumClave 
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   510
         Width           =   615
      End
      Begin VB.TextBox xTipoClave 
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Top             =   220
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Clave:"
         Height          =   165
         Left            =   2160
         TabIndex        =   16
         Top             =   180
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Numero de clave:"
         Height          =   165
         Left            =   120
         TabIndex        =   3
         Top             =   570
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo de clave:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmEditorDecor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private AgregoClave As Boolean
Private ModificoClave As Boolean

Private Sub bExit_Click()
       If AgregoClave Then
        'Modifico clave pero no guardo el decor, raro.
        NumDecorKeys = NumDecorKeys - 1
        Select Case DecorKeys(NumDecorKeys + 1).Tipo_Clave
            Case 1 'Normal
                UltimaDecorKey_Comun = UltimaDecorKey_Comun - 1
        End Select
        ReDim Preserve DecorKeys(1 To NumDecorKeys)
        AgregoClave = False
    End If
    
    Unload Me
    
End Sub

Public Sub Parse()
    If TipoSeleccionado = 1 Then
        If MapData(ObjetoSeleccionado.X, ObjetoSeleccionado.Y).DecorI > 0 Then
            With MapData(ObjetoSeleccionado.X, ObjetoSeleccionado.Y)
                xTipo = DecorData(.DecorI).DecorType
                xNumero = .DecorI
                xDefault = .DecorInfo.EstadoDefault
                xTipoClave = .DecorInfo.TipoClave
                xClave = .DecorInfo.Clave
            End With
        End If
    End If
End Sub

Private Sub bNextClave_Click()
    If AgregoClave Then
        MsgBox "Ya has asignado una nueva clave a este decor."
        Exit Sub
    End If
    If Val(xTipoClave) = 0 Then
        MsgBox "Debes introducir el tipo de clave primero."
        Exit Sub
    End If
    
    xNumClave = AsignarClave(Val(xTipoClave), 1)
    With DecorKeys(Val(xNumClave))
        xClave = .Clave
        .X = ObjetoSeleccionado.X
        .Y = ObjetoSeleccionado.Y
        .Contenedor = 0
    End With
    
    AgregoClave = True
    
End Sub

Private Sub bSave_Click()
Dim j As Long
        If TipoSeleccionado = 1 Then
        If MapData(ObjetoSeleccionado.X, ObjetoSeleccionado.Y).DecorI > 0 Then
            With MapData(ObjetoSeleccionado.X, ObjetoSeleccionado.Y)

                If .DecorInfo.Clave <> Val(xNumClave) Then
                    If AgregoClave Or ModificoClave Then
                        If Val(xNumClave) <> NumDecorKeys Then
                            'Agrego una clave pero no la esta usando.
                            'Cuando llegue al final del sub la va a borrar
                            'Lo que me interesa ver aca es si esta usando una clave o no.
                            If Val(xNumClave) > NumDecorKeys Then
                                'Modifico una clave y le puso algo q esta fuera de rango.
                                MsgBox "La clave que asignaste es invalida."
                                xNumClave = 0 'Se la seteamos en 0
                            ElseIf Val(xNumClave) > 0 Then
                                If Val(xTipoClave) > 0 Then
                                    With DecorKeys(Val(xNumClave))
                                        If Val(xTipoClave) <> .Tipo_Clave Then
                                            'Es distinto el tipo clave, tenemos q asignar un numero de clave.
                                            'Si fuera el mismo sirve el numero de clave.
                                            Select Case Val(xTipoClave)
                                                Case 1
                                                    UltimaDecorKey_Comun = UltimaDecorKey_Comun + 1
                                                    .Clave = UltimaDecorKey_Comun
                                            End Select
                                            .Tipo_Clave = Val(xTipoClave)

                                        End If
                                        .X = ObjetoSeleccionado.X
                                        .Y = ObjetoSeleccionado.Y
                                        .Contenedor = 0
                                        .Tipo_Objeto = 1
                                    End With
                                Else
                                    MsgBox "Asignaste una clave con tipo clave invalida."
                                    xNumClave = 0
                                End If
                            End If
                            modGameIni.Save_DecoKeys Val(xNumClave)
                            
                        Else
                            'Agrego una clave y la esta usando ! Guardamos la data.
                            modGameIni.Save_DecoKeys NumDecorKeys
                            AgregoClave = False
                        End If
                    End If
                    End If
                    .DecorInfo.Clave = Val(xNumClave)
                    .DecorInfo.EstadoDefault = Val(xDefault)

                If Val(xNumero) <> .DecorI Then
                    If Val(xNumero) <= UBound(DecorData) Then
                        .DecorI = Val(xNumero)
                    End If
                End If
                .DecorGrh.index = DECOR_GETGRH_FROMDEFAULT(ObjetoSeleccionado.X, ObjetoSeleccionado.Y)
            End With
        End If
    End If
    If AgregoClave Then
        'Modifico clave pero no guardo el decor, raro.
        NumDecorKeys = NumDecorKeys - 1
        Select Case DecorKeys(NumDecorKeys + 1).Tipo_Clave
            Case 1 'Normal
                UltimaDecorKey_Comun = UltimaDecorKey_Comun - 1
        End Select
        ReDim Preserve DecorKeys(1 To NumDecorKeys)
        AgregoClave = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If AgregoClave Then
        'Modifico clave pero no guardo el decor, raro.
        NumDecorKeys = NumDecorKeys - 1
        Select Case DecorKeys(NumDecorKeys + 1).Tipo_Clave
            Case 1 'Normal
                UltimaDecorKey_Comun = UltimaDecorKey_Comun - 1
        End Select
        ReDim Preserve DecorKeys(1 To NumDecorKeys)
        AgregoClave = False
    End If
End Sub
