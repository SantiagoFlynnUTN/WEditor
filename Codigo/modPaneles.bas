Attribute VB_Name = "modPaneles"
'**************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'**************************************************************

''
' modPaneles
'
' @remarks Funciones referentes a los Paneles de Funcion
' @author gshaxor@gmail.com
' @version 0.3.28
' @date 20060530

Option Explicit

''
' Activa/Desactiva el Estado de la Funcion en el Panel Superior
'
' @param Numero Especifica en numero de funcion
' @param Activado Especifica si esta o no activado

Public Sub EstSelectPanel(ByVal Numero As Byte, ByVal Activado As Boolean)
'*************************************************
'Author: ^[GS]^
'Last modified: 30/05/06
'*************************************************
    If Activado = True Then
        frmMain.SelectPanel(Numero).GradientMode = lv_Bottom2Top
        frmMain.SelectPanel(Numero).HoverBackColor = frmMain.SelectPanel(Numero).GradientColor
        If frmMain.mnuVerAutomatico.Checked = True Then
            Select Case Numero
                Case 0
                    If cCapaSel = 4 Then
                        frmMain.mnuVerCapa4.Tag = CInt(frmMain.mnuVerCapa4.Checked)
                        frmMain.mnuVerCapa4.Checked = True
                    ElseIf cCapaSel = 3 Then
                        frmMain.mnuVerCapa3.Tag = CInt(frmMain.mnuVerCapa3.Checked)
                        frmMain.mnuVerCapa3.Checked = True
                    ElseIf cCapaSel = 2 Then
                        frmMain.mnuVerCapa2.Tag = CInt(frmMain.mnuVerCapa2.Checked)
                        frmMain.mnuVerCapa2.Checked = True
                    ElseIf cCapaSel = 9 Then
                        frmMain.MnuVerCapa9.Tag = CInt(frmMain.MnuVerCapa9.Checked)
                        frmMain.MnuVerCapa9.Checked = True
                        
                    End If
                Case 2
                    frmMain.cVerBloqueos.Tag = CInt(frmMain.cVerBloqueos.value)
                    frmMain.cVerBloqueos.value = True
                    frmMain.mnuVerBloqueos.Checked = frmMain.cVerBloqueos.value
                Case 6
                    frmMain.cVerTriggers.Tag = CInt(frmMain.cVerTriggers.value)
                    frmMain.cVerTriggers.value = True
                    frmMain.mnuVerTriggers.Checked = frmMain.cVerTriggers.value
            End Select
        End If
    Else
        frmMain.SelectPanel(Numero).HoverBackColor = frmMain.SelectPanel(Numero).BackColor
        frmMain.SelectPanel(Numero).GradientMode = lv_NoGradient
        If frmMain.mnuVerAutomatico.Checked = True Then
            Select Case Numero
                Case 0
                    If cCapaSel = 4 Then
                        If LenB(frmMain.mnuVerCapa3.Tag) <> 0 Then frmMain.mnuVerCapa4.Checked = CBool(-1)
                    ElseIf cCapaSel = 3 Then
                        If LenB(frmMain.mnuVerCapa3.Tag) <> 0 Then frmMain.mnuVerCapa3.Checked = CBool(frmMain.mnuVerCapa3.Tag)
                    ElseIf cCapaSel = 2 Then
                        If LenB(frmMain.mnuVerCapa2.Tag) <> 0 Then frmMain.mnuVerCapa2.Checked = CBool(frmMain.mnuVerCapa2.Tag)
                    ElseIf cCapaSel = 9 Then
                        If LenB(frmMain.MnuVerCapa9.Tag) <> 0 Then frmMain.MnuVerCapa9.Checked = CBool(frmMain.MnuVerCapa9.Tag)
                        
                    End If
                Case 2
                    If LenB(frmMain.cVerBloqueos.Tag) = 0 Then frmMain.cVerBloqueos.Tag = 0
                    frmMain.cVerBloqueos.value = CBool(frmMain.cVerBloqueos.Tag)
                    frmMain.mnuVerBloqueos.Checked = frmMain.cVerBloqueos.value
                Case 6
                    If LenB(frmMain.cVerTriggers.Tag) = 0 Then frmMain.cVerTriggers.Tag = 0
                    frmMain.cVerTriggers.value = CBool(frmMain.cVerTriggers.Tag)
                    frmMain.mnuVerTriggers.Checked = frmMain.cVerTriggers.value
            End Select
        End If
    End If
End Sub

''
' Muestra los controles que componen a la funcion seleccionada del Panel
'
' @param Numero Especifica el numero de Funcion
' @param Ver Especifica si se va a ver o no
' @param Normal Inidica que ahi que volver todo No visible

Public Sub VerFuncion(ByVal Numero As Byte, ByVal Ver As Boolean, Optional Normal As Boolean)
'*************************************************
'Author: ^[GS]^
'Last modified: 24/11/08
'*************************************************
    If Normal = True Then
        Call VerFuncion(vMostrando, False, False)
    End If
    frmMain.chkDecorBloq.Visible = False
    Select Case Numero
        Case 0 ' Superficies
            frmMain.cFiltro(0).Visible = Ver
            frmMain.cCapas.Visible = Ver
            frmMain.cGrh.Visible = Ver
            frmMain.cQuitarEnEstaCapa.Visible = Ver
            frmMain.cQuitarEnTodasLasCapas.Visible = Ver
            frmMain.cSeleccionarSuperficie.Visible = Ver
            frmMain.lbFiltrar(0).Visible = Ver
            frmMain.lbCapas.Visible = Ver
            frmMain.lbGrh.Visible = Ver
            frmMain.PreviewGrh.Visible = Ver
            frmMain.StatTxt.Visible = Not Ver
            frmMain.Statz = Not Ver
            If frmMain.Statz Then
                frmMain.VistaStat.Caption = "Vista Previa"
            Else
                frmMain.VistaStat.Caption = "Stats"
            End If
            frmMain.bI.Visible = Ver
            frmMain.cGrill.Visible = Ver
            If Ver = True Then
                If frmMain.bI.value Then
                    frmMain.lListado(0).Visible = False
                    frmMain.lListado(5).Visible = True
                Else
                    frmMain.lListado(0).Visible = True
                    frmMain.lListado(5).Visible = False
                End If
            Else
                frmMain.lListado(0).Visible = False
                frmMain.lListado(5).Visible = False
            End If
        Case 1 ' Translados
            frmMain.lMapN.Visible = Ver
            frmMain.lXhor.Visible = Ver
            frmMain.lYver.Visible = Ver
            frmMain.tTMapa.Visible = Ver
            frmMain.tTX.Visible = Ver
            frmMain.tTY.Visible = Ver
            frmMain.cInsertarTrans.Visible = Ver
            frmMain.cInsertarTransOBJ.Visible = Ver
            frmMain.cUnionManual.Visible = Ver
            frmMain.cUnionAuto.Visible = Ver
            frmMain.cQuitarTrans.Visible = Ver
        Case 2 ' Bloqueos
            frmMain.cQuitarBloqueo.Visible = Ver
            frmMain.cInsertarBloqueo.Visible = Ver
            frmMain.cVerBloqueos.Visible = Ver
        Case 3  ' NPCs
            frmMain.lListado(1).Visible = Ver
            frmMain.lListado(7).Visible = False
            frmMain.cFiltro(1).Visible = Ver
            frmMain.lbFiltrar(1).Visible = Ver
            frmMain.lNumFunc(Numero - 3).Visible = Ver
            frmMain.cNumFunc(Numero - 3).Visible = Ver
            frmMain.cInsertarFunc(Numero - 3).Visible = Ver
            frmMain.cQuitarFunc(Numero - 3).Visible = Ver
            frmMain.lCantFunc(Numero - 3).Visible = Ver
            frmMain.cCantFunc(Numero - 3).Visible = Ver
            frmMain.decorb.Caption = "Hostiles"
            frmMain.decorb.value = False
            frmMain.decorb.Visible = Ver
            
        Case 4 ' NPCs Hostiles
            'frmMain.lListado(1).Visible = Ver
            'frmMain.cFiltro(1).Visible = Ver
            'frmMain.lbFiltrar(1).Visible = Ver
            'frmMain.lNumFunc(Numero - 3).Visible = Ver
            'frmMain.cNumFunc(Numero - 3).Visible = Ver
            'frmMain.cInsertarFunc(Numero - 3).Visible = Ver
            'frmMain.cQuitarFunc(Numero - 3).Visible = Ver
            'frmMain.lCantFunc(Numero - 3).Visible = Ver
            'frmMain.cCantFunc(Numero - 3).Visible = Ver
        Case 5 ' OBJs
            frmMain.lListado(3).Visible = Ver
            frmMain.lListado(6).Visible = False
            frmMain.cFiltro(3).Visible = Ver
            frmMain.lbFiltrar(3).Visible = Ver
            frmMain.lNumFunc(Numero - 3).Visible = Ver
            frmMain.cNumFunc(Numero - 3).Visible = Ver
            frmMain.cInsertarFunc(Numero - 3).Visible = Ver
            frmMain.cQuitarFunc(Numero - 3).Visible = Ver
            frmMain.lCantFunc(Numero - 3).Visible = Ver
            frmMain.cCantFunc(Numero - 3).Visible = Ver
            frmMain.decorb.Visible = Ver
            frmMain.decorb.value = False
            frmMain.decorb.Caption = "Decors"
        Case 6 ' Triggers
            If Ver Then
                SobreIndex = 2
            End If
            frmMain.cQuitarTrigger.Visible = Ver
            frmMain.cInsertarTrigger.Visible = Ver
            frmMain.cVerTriggers.Visible = Ver
            frmMain.cVerTriggers.Caption = "Mostrar Triggers"
            frmMain.cInsertarTrigger.Caption = "Insertar Triggers"
            frmMain.cQuitarTrigger.Caption = "Quitar Triggers"
            frmMain.lListado(4).Visible = Ver
            frmMain.decorb.Visible = Ver
            If Ver Then
                frmMain.decorb.Caption = "Tipo Terreno"
            End If
            frmMain.lListado(8).Visible = False
        
        Case 7
            frmMain.cLuces.Visible = Ver
        Case 8
            frmMain.cParticulas.Visible = Ver
        Case 9
            If Ver Then
            frmMain.frmSPOTLIGHTS.Visible = True
            Else
            frmMain.frmSPOTLIGHTS.Visible = False
            End If
        Case 10
            If Ver Then
                frmMain.pPaneles.Visible = False
                frmMain.MpNw.Visible = True
                frmMain.LayerC.ListIndex = 0
                frmMain.SizeC.ListIndex = 0
                frmMain.StatTxt.Visible = False
            Else
                frmMain.MpNw.Visible = False
                frmMain.pPaneles.Visible = True
                frmMain.StatTxt.Visible = True
            End If
    End Select
    If Ver = True Then
        vMostrando = Numero
        If Numero < 0 Or Numero > 6 Then Exit Sub
        If frmMain.SelectPanel(Numero).value = False Then
            frmMain.SelectPanel(Numero).value = True
        End If
    Else
        If Numero < 0 Or Numero > 6 Then Exit Sub
        If frmMain.SelectPanel(Numero).value = True Then
            frmMain.SelectPanel(Numero).value = False
        End If
    End If
End Sub

''
' Filtra del Listado de Elementos de una Funcion
'
' @param Numero Indica la funcion a Filtrar

Public Sub Filtrar(ByVal Numero As Byte)
'*************************************************
'Author: ^[GS]^
'Last modified: 26/05/06
'*************************************************

    Dim vMaximo As Integer
    Dim vDatos As String
    Dim NumI As Integer
    Dim vMin As Integer
    Dim i As Integer
    Dim j As Integer
    Dim Tipo As Byte
    
    Tipo = Numero
    If Numero = 1 And frmMain.decorb.value Then Tipo = 7
    If Numero = 3 And frmMain.decorb.value Then Tipo = 6
    
    
    vDatos = frmMain.cFiltro(Numero).Text
    If LenB(frmMain.cFiltro(Numero).Text) > 0 Then
    If frmMain.cFiltro(Numero).ListCount > 0 Then
        For j = 0 To frmMain.cFiltro(Numero).ListCount - 1
            If UCase$(frmMain.cFiltro(Numero).List(j)) = UCase$(frmMain.cFiltro(Numero).Text) Then
                Exit For
            End If
        Next j
        
        If j > frmMain.cFiltro(Numero).ListCount - 1 Then 'Sino estaba borramos el mas viejo.
            If frmMain.cFiltro(Numero).ListCount > 5 Then
                frmMain.cFiltro(Numero).RemoveItem 0
            End If
            frmMain.cFiltro(Numero).AddItem frmMain.cFiltro(Numero).Text
        Else
        
            If j < (frmMain.cFiltro(Numero).ListCount - 1) Then
                Debug.Print (frmMain.cFiltro(Numero).ListCount - 1)
                For i = j To frmMain.cFiltro(Numero).ListCount - 2
                    frmMain.cFiltro(Numero).List(i) = frmMain.cFiltro(Numero).List(i + 1)
                Next i
                Debug.Print frmMain.cFiltro(Numero).Text
                frmMain.cFiltro(Numero).List(frmMain.cFiltro(Numero).ListCount - 1) = vDatos
            End If
        End If
    Else
            frmMain.cFiltro(Numero).AddItem frmMain.cFiltro(Numero).Text
    End If
    End If
    frmMain.lListado(Tipo).Clear
    frmMain.cFiltro(Numero).Text = vDatos
    vMin = 1
    Select Case Tipo
        Case 0 ' superficie
            vMaximo = NumTexWe
        Case 1 ' NPCs
            vMaximo = NumNPCs
        Case 3 ' Objetos
            vMaximo = NumOBJs
        Case 7 'Hostiels
            vMaximo = NumNPCsHOST
            vMin = 500
        Case 6 'DEcors
            vMaximo = numDecor
    End Select
    
    For i = vMin To vMaximo
    
        Select Case Tipo
            Case 0 ' superficie
                vDatos = TexWE(i).Name
                NumI = i
            Case 1 ' NPCs
                vDatos = NpcData(i).Name
                NumI = i
            Case 3 ' Objetos
                vDatos = ObjData(i).Name
                NumI = i
            Case 7
                vDatos = NpcData(i).Name
                NumI = i
            Case 6
                vDatos = DecorData(i).Name
                NumI = i
        End Select
        
        For j = 1 To Len(vDatos)
            If UCase$(mid$(vDatos & Str(i), j, Len(frmMain.cFiltro(Numero).Text))) = UCase$(frmMain.cFiltro(Numero).Text) Or LenB(frmMain.cFiltro(Numero).Text) = 0 Then
                frmMain.lListado(Tipo).AddItem vDatos & " - [" & NumI & "]"
                Exit For
            End If
        Next
    Next
End Sub

Public Function DameGrhIndex(ByVal GrhIn As Integer) As Integer
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************

DameGrhIndex = SupData(GrhIn).Grh

If SupData(GrhIn).Width > 0 Then
    frmConfigSup.MOSAICO.value = vbChecked
    frmConfigSup.mAncho.Text = SupData(GrhIn).Width
    frmConfigSup.mLargo.Text = SupData(GrhIn).Height
Else
    frmConfigSup.MOSAICO.value = vbUnchecked
    frmConfigSup.mAncho.Text = "0"
    frmConfigSup.mLargo.Text = "0"
End If



End Function

Public Sub fPreviewGrh(ByVal GrhIn As Integer)
'*************************************************
'Author: Unkwown
'Last modified: 22/05/06
'*************************************************

If Val(GrhIn) < 1 Then
  frmMain.cGrh.Text = MaxGrhs
  Exit Sub
End If

If Val(GrhIn) > MaxGrhs Then
  frmMain.cGrh.Text = 1
  Exit Sub
End If

'Change CurrentGrh
CurrentGrh.grh_index = GrhIn
CurrentGrh.Started = 1
CurrentGrh.frame_counter = 1
CurrentGrh.frame_speed = grh_list(CurrentGrh.grh_index).frame_speed

End Sub

''
' Indica la accion de mostrar Vista Previa de la Superficie seleccionada
'

Public Sub VistaPreviaDeSup()
frmMain.PreviewGrh.Cls
If SelTexWe = 0 Then Exit Sub
With TexWE(SelTexWe)
If .Ancho = 0 Or .Largo = 0 Then Exit Sub
Dim P As Long
Dim R As RECT
Dim d As RECT
If .NumIndex = 0 Then Exit Sub
R.Bottom = .Largo
R.Right = .Ancho
ddevice.Clear 1, R, D3DCLEAR_TARGET, &H0, ByVal 0, 0
For P = 1 To .NumIndex
    
    'dibuja en el vista previa
    modDXEngine.DibujareEnHwnd2 frmMain.PreviewGrh.hWnd, .index(P).Num, R, .index(P).X, .index(P).Y, False


Next P


d.left = 0
d.top = 0
d.Bottom = .Largo
d.Right = .Ancho
ddevice.Present R, d, frmMain.PreviewGrh.hWnd, ByVal 0


If frmMain.cGrill.value Then
frmMain.PreviewGrh.ForeColor = vbCyan
frmMain.PreviewGrh.DrawWidth = 1
For P = 1 To .NumIndex

    frmMain.PreviewGrh.Line (.index(P).X * Screen.TwipsPerPixelX, .index(P).Y * Screen.TwipsPerPixelY)-(.index(P).X * Screen.TwipsPerPixelX, (.index(P).Y + EstaticData(NewIndexData(.index(P).Num).Estatic).H) * Screen.TwipsPerPixelY)
    frmMain.PreviewGrh.Line (.index(P).X * Screen.TwipsPerPixelX, .index(P).Y * Screen.TwipsPerPixelY)-((.index(P).X + EstaticData(NewIndexData(.index(P).Num).Estatic).W) * Screen.TwipsPerPixelX, .index(P).Y * Screen.TwipsPerPixelY)
    frmMain.PreviewGrh.Line ((.index(P).X + EstaticData(NewIndexData(.index(P).Num).Estatic).W) * Screen.TwipsPerPixelX, .index(P).Y * Screen.TwipsPerPixelY)-((.index(P).X + EstaticData(NewIndexData(.index(P).Num).Estatic).W) * Screen.TwipsPerPixelX, (.index(P).Y + EstaticData(NewIndexData(.index(P).Num).Estatic).H) * Screen.TwipsPerPixelY)
    frmMain.PreviewGrh.Line (.index(P).X * Screen.TwipsPerPixelX, (.index(P).Y + EstaticData(NewIndexData(.index(P).Num).Estatic).H) * Screen.TwipsPerPixelY)-((.index(P).X + EstaticData(NewIndexData(.index(P).Num).Estatic).W) * Screen.TwipsPerPixelX, (.index(P).Y + EstaticData(NewIndexData(.index(P).Num).Estatic).H) * Screen.TwipsPerPixelY)


Next P
End If
If SelTexFrame > 0 Then
P = SelTexFrame
frmMain.PreviewGrh.ForeColor = vbYellow
frmMain.PreviewGrh.DrawWidth = 2
    frmMain.PreviewGrh.Line (.index(P).X, .index(P).Y)-(.index(P).X, (.index(P).Y + EstaticData(NewIndexData(.index(P).Num).Estatic).H))
    frmMain.PreviewGrh.Line (.index(P).X, .index(P).Y)-((.index(P).X + EstaticData(NewIndexData(.index(P).Num).Estatic).W), .index(P).Y)
    frmMain.PreviewGrh.Line ((.index(P).X + EstaticData(NewIndexData(.index(P).Num).Estatic).W), .index(P).Y)-((.index(P).X + EstaticData(NewIndexData(.index(P).Num).Estatic).W), (.index(P).Y + EstaticData(NewIndexData(.index(P).Num).Estatic).H))
    frmMain.PreviewGrh.Line (.index(P).X, (.index(P).Y + EstaticData(NewIndexData(.index(P).Num).Estatic).H))-((.index(P).X + EstaticData(NewIndexData(.index(P).Num).Estatic).W), (.index(P).Y + EstaticData(NewIndexData(.index(P).Num).Estatic).H))



End If

End With
End Sub

