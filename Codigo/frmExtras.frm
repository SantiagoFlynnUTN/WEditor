VERSION 5.00
Begin VB.Form frmExtras 
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6525
   LinkTopic       =   "Form2"
   ScaleHeight     =   5070
   ScaleWidth      =   6525
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check15 
      Caption         =   "Borrar Npcs"
      Height          =   195
      Left            =   360
      TabIndex        =   21
      Top             =   3720
      Width           =   1815
   End
   Begin VB.CheckBox Check14 
      Caption         =   "Bucle Lista"
      Height          =   195
      Left            =   360
      TabIndex        =   20
      Top             =   3480
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Mod Index"
      Height          =   255
      Left            =   3480
      TabIndex        =   19
      Top             =   840
      Width           =   1695
   End
   Begin VB.CheckBox Check13 
      Caption         =   "Modificar"
      Height          =   195
      Left            =   360
      TabIndex        =   18
      Top             =   3240
      Width           =   1815
   End
   Begin VB.TextBox t2 
      Height          =   285
      Left            =   2520
      TabIndex        =   17
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CheckBox Check12 
      Caption         =   "Track Grafico"
      Height          =   195
      Left            =   360
      TabIndex        =   16
      Top             =   3000
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Mover Graficos Used"
      Height          =   255
      Left            =   3480
      TabIndex        =   15
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Recargar NewEstatic"
      Height          =   255
      Left            =   3480
      TabIndex        =   14
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Recargar NewIndex"
      Height          =   255
      Left            =   3480
      TabIndex        =   13
      Top             =   120
      Width           =   1695
   End
   Begin VB.CheckBox Check11 
      Caption         =   "Separar Decors"
      Height          =   195
      Left            =   360
      TabIndex        =   12
      Top             =   2760
      Width           =   1815
   End
   Begin VB.CheckBox Check10 
      Caption         =   "Guardar Mapas"
      Height          =   195
      Left            =   360
      TabIndex        =   11
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CheckBox Check9 
      Caption         =   "Load NuevoIndex"
      Height          =   195
      Left            =   360
      TabIndex        =   10
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CheckBox Check8 
      Caption         =   "Reemplazar Grhindex"
      Height          =   195
      Left            =   360
      TabIndex        =   9
      Top             =   2040
      Width           =   1815
   End
   Begin VB.CheckBox Check7 
      Caption         =   "Separar Graficos"
      Height          =   195
      Left            =   360
      TabIndex        =   8
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CheckBox Check6 
      Caption         =   "Reindexar Estaticos"
      Height          =   195
      Left            =   360
      TabIndex        =   7
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Todos los mapas"
      Height          =   195
      Left            =   360
      TabIndex        =   5
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CheckBox Check5 
      Caption         =   "Obtener Npcs"
      Height          =   195
      Left            =   360
      TabIndex        =   4
      Top             =   840
      Width           =   1815
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Obtener Objetos"
      Height          =   195
      Left            =   360
      TabIndex        =   3
      Top             =   600
      Width           =   1815
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Obtener Grhindex"
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   360
      Width           =   1815
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Obtener Graficos"
      Height          =   195
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Procesar"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   3960
      Width           =   1695
   End
End
Attribute VB_Name = "frmExtras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim p As Long
Dim Mapa As Integer
Dim L As String
Dim j  As Integer
Dim n As Integer
Dim X As String
Dim k As Long
Dim u As Long
If frmExtras.Check14.value Then
    u = Val(GetVar(App.PATH & "\RESOURCES\INIT\EQUIV.dat", "INIT", "NUM"))
Else
    u = 1
End If
    
For k = 1 To u

If Check14.value Then
    frmExtras.t2.Text = GetVar(App.PATH & "\RESOURCES\INIT\EQUIV.dat", CStr(k), "E")
    Command5_Click
End If

Erase Objetos
Erase Graficos
Erase GrhIndex
Erase Npcs
If Check9.value Then
    X = "C:\Maps\NUEVO_INDEX\"
Else
    X = "C:\Maps\"
End If

If Check4.value Then
For p = 1 To 165 '
   ReDim MapData(1 To 100, 1 To 100)
   Erase CharList
   Label1.Caption = "Procesando " & p & "/160"
   DoEvents
   modMapIO.AbrirMapaComun X & p & ".map"
   Procesar p

Next p

    Label1.Caption = "Listo!."
Else

L = (InputBox("¿ Que mapa quieres procesar ?"))
If InStr(L, "<") Then
n = Val(ReadField(2, L, Asc("<")))

For p = 1 To n '
   ReDim MapData(1 To 100, 1 To 100)
   Erase CharList
   Label1.Caption = "Procesando " & p & "/" & n
   DoEvents
   modMapIO.AbrirMapaComun X & p & ".map"
   Procesar p

Next p

    Label1.Caption = "Listo!."
ElseIf InStr(L, "-") Then
    j = Val(ReadField(1, L, Asc("-")))
    n = Val(ReadField(2, L, Asc("-")))
    For p = j To n
   ReDim MapData(1 To 100, 1 To 100)
   Erase CharList
   Label1.Caption = "Procesando " & p & "/" & n
   DoEvents
   modMapIO.AbrirMapaComun X & p & ".map"
   Procesar p
    
    Next p
        Label1.Caption = "Listo!."
ElseIf InStr(L, "I") Then
    Mapa = Val(ReadField(2, L, Asc("I")))
   ReDim MapData(1 To 100, 1 To 100)
   Erase CharList
   Label1.Caption = "Procesando 1/1."
   DoEvents
   
   modMapIO.AbrirMapaComun X & "I" & Mapa & ".map"
   Procesar Mapa, True
   Label1.Caption = "Listo."
Else
    Mapa = Val(L)
   ReDim MapData(1 To 100, 1 To 100)
   Erase CharList
   Label1.Caption = "Procesando 1/1."
   DoEvents
   
   modMapIO.AbrirMapaComun X & Mapa & ".map"
   Procesar Mapa
   Label1.Caption = "Listo."
End If





End If

If Check11.value Then
    Label1.Caption = "Separando decors"
    DoEvents
    modExtras.SepararDecors
    Label1.Caption = "Listo!"
End If

If Check14.value Then
    frmExtras.Caption = k
    DoEvents
    
End If

Next k
End Sub

Private Sub Command2_Click()
Load_NewIndex
MsgBox "OK"
End Sub


Private Sub Command3_Click()
Load_NewEstatics
MsgBox "OK"
End Sub

Private Sub Command4_Click()
Dim i As Long

For i = 1 To numNewIndex
    If Not FileExist(App.PATH & "\Resources\Graficos Usados\" & NewIndexData(i).OverWriteGrafico & ".png", vbNormal) Then
        If FileExist(App.PATH & "\Resources\Graficos\" & NewIndexData(i).OverWriteGrafico & ".png", vbNormal) Then
            FileCopy App.PATH & "\Resources\Graficos\" & NewIndexData(i).OverWriteGrafico & ".png", App.PATH & "\Resources\Graficos Usados\" & NewIndexData(i).OverWriteGrafico & ".png"
        Else
            Stop
        End If
    End If
Next i
MsgBox "listo"
End Sub

Private Sub Command5_Click()
Dim p As Long
Dim j As Long
Dim oG As Integer
Dim nG As Integer
Dim k As Boolean
k = True
oG = Val(ReadField(1, frmExtras.t2.Text, Asc("-")))
nG = Val(ReadField(2, frmExtras.t2.Text, Asc("-")))
If Not k Then
For p = 1 To numNewIndex
    With NewIndexData(p)
    If .OverWriteGrafico = oG Then
        If EstaticData(.Estatic).L >= 256 Then
            For j = 1 To numNewIndex
                
                If NewIndexData(j).OverWriteGrafico = nG Then
                    If EstaticData(NewIndexData(j).Estatic).W = EstaticData(.Estatic).W And EstaticData(NewIndexData(j).Estatic).H = EstaticData(.Estatic).H Then
                        If EstaticData(NewIndexData(j).Estatic).L = EstaticData(.Estatic).L - 256 And EstaticData(NewIndexData(j).Estatic).t = EstaticData(.Estatic).t + 128 Then
                            WriteVar App.PATH & "\Resources\init\NewIndex.dat", CStr(p), "OverWriteGrafico", "Reciclable"
                            Exit For
                        End If
                    End If
                End If
            Next j
            If j > numNewIndex Then
                WriteVar App.PATH & "\Resources\init\NewIndex.dat", CStr(p), "OverWriteGrafico", CStr(nG)
                NewIndexData(p).OverWriteGrafico = nG
                For j = 1 To numNewEstatic
                    If EstaticData(j).W = EstaticData(.Estatic).W And EstaticData(j).H = EstaticData(.Estatic).H Then
                        If EstaticData(j).L = EstaticData(.Estatic).L - 256 And EstaticData(j).t = EstaticData(.Estatic).t + 128 Then
                            WriteVar App.PATH & "\Resources\init\NewIndex.dat", CStr(p), "Estatica", CStr(j)
                            NewIndexData(p).Estatic = j
                            Exit For
                        End If
                    End If
                Next j
                If j > numNewEstatic Then Stop
            End If
        
        Else
            For j = 1 To numNewIndex
                
                If NewIndexData(j).OverWriteGrafico = nG Then
                    If NewIndexData(j).Estatic = .Estatic Then
                        WriteVar App.PATH & "\Resources\init\NewIndex.dat", CStr(p), "OverWriteGrafico", "Reciclable"
                        Exit For
                    End If
                End If
            Next j
            If j > numNewIndex Then
                WriteVar App.PATH & "\Resources\init\NewIndex.dat", CStr(p), "OverWriteGrafico", CStr(nG)
                NewIndexData(p).OverWriteGrafico = (nG)
            End If
        End If
    End If
    End With
Next p
Else
For p = 1 To numNewIndex
    With NewIndexData(p)
    If .OverWriteGrafico = oG Then
        If EstaticData(.Estatic).t >= 256 Then
            For j = 1 To numNewIndex
                
                If NewIndexData(j).OverWriteGrafico = nG + 1 Then
                    If EstaticData(NewIndexData(j).Estatic).W = EstaticData(.Estatic).W And EstaticData(NewIndexData(j).Estatic).H = EstaticData(.Estatic).H Then
                        If EstaticData(NewIndexData(j).Estatic).L = EstaticData(.Estatic).L And EstaticData(NewIndexData(j).Estatic).t = EstaticData(.Estatic).t - 256 Then
                            WriteVar App.PATH & "\Resources\init\NewIndex.dat", CStr(p), "OverWriteGrafico", "Reciclable"
                            Exit For
                        End If
                    End If
                End If
            Next j
            If j > numNewIndex Then
                WriteVar App.PATH & "\Resources\init\NewIndex.dat", CStr(p), "OverWriteGrafico", CStr(nG + 1)
                NewIndexData(p).OverWriteGrafico = nG + 1
                For j = 1 To numNewEstatic
                    If EstaticData(j).W = EstaticData(.Estatic).W And EstaticData(j).H = EstaticData(.Estatic).H Then
                        If EstaticData(j).L = EstaticData(.Estatic).L And EstaticData(j).t = EstaticData(.Estatic).t - 256 Then
                            WriteVar App.PATH & "\Resources\init\NewIndex.dat", CStr(p), "Estatica", CStr(j)
                            NewIndexData(p).Estatic = j
                            Exit For
                        End If
                    End If
                Next j
                If j > numNewEstatic Then Stop
            End If
        
        Else
            For j = 1 To numNewIndex
                
                If NewIndexData(j).OverWriteGrafico = nG Then
                    If NewIndexData(j).Estatic = .Estatic Then
                        WriteVar App.PATH & "\Resources\init\NewIndex.dat", CStr(p), "OverWriteGrafico", "Reciclable"
                        Exit For
                    End If
                End If
            Next j
            If j > numNewIndex Then
                WriteVar App.PATH & "\Resources\init\NewIndex.dat", CStr(p), "OverWriteGrafico", CStr(nG)
                NewIndexData(p).OverWriteGrafico = (nG)
            End If
        End If
    End If
    End With
Next p


End If

End Sub

