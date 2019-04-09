Attribute VB_Name = "modIndices"
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
' modIndices
'
' @remarks Funciones Especificas al Trabajo con Indices
' @author gshaxor@gmail.com
' @version 0.1.05
' @date 20060530

Option Explicit

' Carga los indices de Superficie
'

Public Sub CargarIndicesSuperficie()
'*************************************************
'Author: ^[GS]^
'Last modified: 29/05/06
'*************************************************

On Error GoTo Fallo
    If General_File_Exist(inipath & "GrhIndex\indices.ini", vbArchive) = False Then
        MsgBox "Falta el archivo 'GrhIndex\indices.ini'", vbCritical
        End
    End If
    Dim Leer As New clsIniReader
    Dim i As Integer
    Leer.Initialize inipath & "GrhIndex\indices.ini"
    MaxSup = Leer.GetValue("INIT", "Referencias")
    ReDim SupData(MaxSup) As SupData
    frmMain.lListado(0).Clear
    For i = 0 To MaxSup
        SupData(i).Name = Leer.GetValue("REFERENCIA" & i, "Nombre")
        SupData(i).Grh = Val(Leer.GetValue("REFERENCIA" & i, "GrhIndice"))
        SupData(i).Width = Val(Leer.GetValue("REFERENCIA" & i, "Ancho"))
        SupData(i).Height = Val(Leer.GetValue("REFERENCIA" & i, "Alto"))
        SupData(i).Block = IIf(Val(Leer.GetValue("REFERENCIA" & i, "Bloquear")) = 1, True, False)
        SupData(i).Capa = Val(Leer.GetValue("REFERENCIA" & i, "Capa"))
        frmMain.lListado(0).AddItem SupData(i).Name & " - #" & i
    Next
    DoEvents
    Exit Sub
Fallo:
    MsgBox "Error al intentar cargar el indice " & i & " de GrhIndex\indices.ini" & vbCrLf & "Err: " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly
End Sub

''
' Carga los indices de Objetos
'

Public Sub CargarIndicesOBJ()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

On Error GoTo Fallo
    If General_File_Exist(DirDats & "\OBJ.dat", vbArchive) = False Then
        MsgBox "Falta el archivo 'OBJ.dat' en " & DirDats, vbCritical
        End
    End If
    Dim Obj As Integer
    Dim Leer As New clsIniReader
    Call Leer.Initialize(DirDats & "\OBJ.dat")
    frmMain.lListado(3).Clear
    NumOBJs = Val(Leer.GetValue("INIT", "NumOBJs"))
    ReDim ObjData(1 To NumOBJs) As ObjData
    For Obj = 1 To NumOBJs
        frmCargando.X.Caption = "Cargando Datos de Objetos..." & Obj & "/" & NumOBJs
        DoEvents
        ObjData(Obj).Name = Leer.GetValue("OBJ" & Obj, "Name")
        ObjData(Obj).grh_index = Val(Leer.GetValue("OBJ" & Obj, "GrhIndex"))
        ObjData(Obj).ObjType = Val(Leer.GetValue("OBJ" & Obj, "ObjType"))
        ObjData(Obj).Ropaje = Val(Leer.GetValue("OBJ" & Obj, "NumRopaje"))
        ObjData(Obj).Info = Leer.GetValue("OBJ" & Obj, "Info")
        ObjData(Obj).WeaponAnim = Val(Leer.GetValue("OBJ" & Obj, "Anim"))
        ObjData(Obj).Texto = Leer.GetValue("OBJ" & Obj, "Texto")
        ObjData(Obj).GrhSecundario = Val(Leer.GetValue("OBJ" & Obj, "GrhSec"))
        frmMain.lListado(3).AddItem ObjData(Obj).Name & " - #" & Obj
    Next Obj
    Exit Sub
Fallo:
MsgBox "Error al intentar cargar el Objteto " & Obj & " de OBJ.dat en " & DirDats & vbCrLf & "Err: " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly

End Sub

''
' Carga los indices de Triggers
'

Public Sub CargarIndicesTriggers()
'*************************************************
'Author: ^[GS]^
'Last modified: 28/05/06
'*************************************************

On Error GoTo Fallo
    If General_File_Exist(DirIndex & "Triggers.ini", vbArchive) = False Then
        MsgBox "Falta el archivo 'Triggers.ini' en " & DirIndex, vbCritical
        End
    End If
    Dim NumT As Integer
    Dim t As Integer
    Dim Leer As New clsIniReader
    Call Leer.Initialize(DirIndex & "Triggers.ini")
    frmMain.lListado(4).Clear
    NumT = Val(Leer.GetValue("INIT", "NumTriggers"))
    For t = 1 To NumT
         frmMain.lListado(4).AddItem Leer.GetValue("Trig" & t, "Name") & " - #" & (t - 1)
    Next t

Exit Sub
Fallo:
    MsgBox "Error al intentar cargar el Trigger " & t & " de Triggers.ini en " & DirIndex & vbCrLf & "Err: " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly

End Sub

''
' Carga los indices de Cuerpos
'




''
' Carga los indices de NPCs
'

Public Sub CargarIndicesNPC()
'*************************************************
'Author: ^[GS]^
'Last modified: 26/05/06
'*************************************************
On Error Resume Next
'On Error GoTo Fallo
    If General_File_Exist(DirDats & "\NPCs.dat", vbArchive) = False Then
        MsgBox "Falta el archivo 'NPCs.dat' en " & DirDats, vbCritical
        End
    End If
    'If general_file_exist(DirDats & "\NPCs-HOSTILES.dat", vbArchive) = False Then
    '    MsgBox "Falta el archivo 'NPCs-HOSTILES.dat' en " & DirDats, vbCritical
    '    End
    'End If
    Dim Trabajando As String
    Dim NPC As Integer
    Dim Leer As New clsIniReader
    Dim Leer2 As New clsIniReader
    frmMain.lListado(1).Clear
    frmMain.lListado(2).Clear
    Call Leer.Initialize(DirDats & "\NPCs.dat")
    NumNPCs = Val(Leer.GetValue("INIT", "NumNPCs"))
    Call Leer2.Initialize(DirDats & "\NPCs-HOSTILES.dat")
    NumNPCsHOST = Val(Leer2.GetValue("INIT", "NumNPCs"))
    ReDim NpcData(1000) As NpcData
    Trabajando = "Dats\NPCs.dat"
    'Call Leer.Initialize(DirDats & "\NPCs.dat")
    'MsgBox "  "
    For NPC = 1 To NumNPCs
        NpcData(NPC).Name = Leer.GetValue("NPC" & NPC, "Name")
        
        NpcData(NPC).Body = Val(Leer.GetValue("NPC" & NPC, "Body"))
        NpcData(NPC).Head = Val(Leer.GetValue("NPC" & NPC, "Head"))
        NpcData(NPC).Heading = Val(Leer.GetValue("NPC" & NPC, "Heading"))
        
        NpcData(NPC).MaxHP = Val(Leer2.GetValue("NPC" & (NPC), "MaxHP"))
        NpcData(NPC).MinHit = Val(Leer2.GetValue("NPC" & (NPC), "MinHit"))
        NpcData(NPC).MaxHit = Val(Leer2.GetValue("NPC" & (NPC), "MaxHit"))
        NpcData(NPC).Def = Val(Leer2.GetValue("NPC" & (NPC), "Def"))
        NpcData(NPC).Exp = Val(Leer2.GetValue("NPC" & (NPC), "GiveExp"))
        NpcData(NPC).Oro = Val(Leer2.GetValue("NPC" & (NPC), "GiveGLD"))
        
                NpcData(NPC).ModHp = Val(Leer2.GetValue("NPC" & (NPC), "ModHP"))
        NpcData(NPC).ModHit = Val(Leer2.GetValue("NPC" & (NPC), "ModHit"))
        NpcData(NPC).ModDef = Val(Leer2.GetValue("NPC" & (NPC), "ModDef"))
        NpcData(NPC).ModExp = Val(Leer2.GetValue("NPC" & (NPC), "ModExp"))
        NpcData(NPC).ModOro = Val(Leer2.GetValue("NPC" & (NPC), "ModOro"))
        
        If LenB(NpcData(NPC).Name) <> 0 Then frmMain.lListado(1).AddItem NpcData(NPC).Name & " - #" & NPC
    Next NPC
    'MsgBox "  "
    'Trabajando = "Dats\NPCs-HOSTILES.dat"

    For NPC = 500 To NumNPCsHOST
        NpcData(NPC).Name = Leer2.GetValue("NPC" & (NPC), "Name")
        NpcData(NPC).Body = Val(Leer2.GetValue("NPC" & (NPC), "Body"))
        NpcData(NPC).Head = Val(Leer2.GetValue("NPC" & (NPC), "Head"))
        NpcData(NPC).Heading = Val(Leer2.GetValue("NPC" & (NPC), "Heading"))
        If LenB(NpcData(NPC).Name) <> 0 Then frmMain.lListado(7).AddItem NpcData(NPC).Name & " - #" & (NPC)
    
        NpcData(NPC).MaxHP = Val(Leer2.GetValue("NPC" & (NPC), "MaxHP"))
        NpcData(NPC).MinHit = Val(Leer2.GetValue("NPC" & (NPC), "MinHit"))
        NpcData(NPC).MaxHit = Val(Leer2.GetValue("NPC" & (NPC), "MaxHit"))
        NpcData(NPC).Def = Val(Leer2.GetValue("NPC" & (NPC), "Def"))
        NpcData(NPC).Exp = Val(Leer2.GetValue("NPC" & (NPC), "GiveExp"))
        NpcData(NPC).Oro = Val(Leer2.GetValue("NPC" & (NPC), "GiveGLD"))
        
        NpcData(NPC).ModHp = Val(Leer2.GetValue("NPC" & (NPC), "ModHP"))
        NpcData(NPC).ModHit = Val(Leer2.GetValue("NPC" & (NPC), "ModHit"))
        NpcData(NPC).ModDef = Val(Leer2.GetValue("NPC" & (NPC), "ModDef"))
        NpcData(NPC).ModExp = Val(Leer2.GetValue("NPC" & (NPC), "ModExp"))
        NpcData(NPC).ModOro = Val(Leer2.GetValue("NPC" & (NPC), "ModOro"))
        
        
    
    Next NPC
    Exit Sub
Fallo:
    MsgBox "Error al intentar cargar el NPC " & NPC & " de " & Trabajando & " en " & DirDats & vbCrLf & "Err: " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly

End Sub

