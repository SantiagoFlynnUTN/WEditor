VERSION 5.00
Begin VB.Form frmMusica 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Musica"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5430
   Icon            =   "frmMusica.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   5430
   StartUpPosition =   2  'CenterScreen
   Begin WorldEditor.lvButtons_H cmdAplicarYCerrar 
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Top             =   1920
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
      CapAlign        =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFHover         =   0
      Focus           =   0   'False
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      ImgSize         =   0
      cBack           =   0
   End
   Begin WorldEditor.lvButtons_H CmdEscuchar 
      Height          =   615
      Left            =   3120
      TabIndex        =   1
      Top             =   120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1085
      CapAlign        =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFHover         =   0
      Focus           =   0   'False
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      ImgSize         =   0
      cBack           =   0
   End
   Begin VB.FileListBox fleMusicas 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1665
      Left            =   120
      Pattern         =   "*.mid"
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin WorldEditor.lvButtons_H cmddetener 
      Height          =   615
      Left            =   3120
      TabIndex        =   2
      Top             =   960
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1085
      CapAlign        =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFHover         =   0
      Focus           =   0   'False
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      ImgSize         =   0
      cBack           =   0
   End
End
Attribute VB_Name = "frmMusica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Option Explicit

Private MidiActual As String

''
' Aplica la Musica seleccionada y oculta la ventana
'

Private Sub cmdAplicarYCerrar_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
On Error Resume Next
If Len(MidiActual) >= 5 Then
    MapInfo.Music = left(MidiActual, Len(MidiActual) - 4)
    frmMapInfo.txtMapMusica.Text = MapInfo.Music
    MidiActual = Empty
End If
Me.Hide
End Sub


''
' Oculta la ventana
'

Private Sub cmdCerrar_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Me.Hide
End Sub

''
' Detiene la Musica que se encuentra Reproduciendo
'

Private Sub cmdDetener_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Audio.StopMidi
CmdEscuchar.Enabled = True
cmddetener.Enabled = False
Play = False
End Sub

''
' Inicia la reproduccion de la Musica Seleccionada
'

Private Sub cmdEscuchar_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Audio.PlayMIDI fleMusicas.List(fleMusicas.ListIndex)
cmddetener.Enabled = True
CmdEscuchar.Enabled = False
Play = True
End Sub

''
' Selecciona una nueva Musica del listado
'

Private Sub fleMusicas_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
MidiActual = fleMusicas.List(fleMusicas.ListIndex)
cmdAplicarYCerrar.Enabled = True
If Play = False Then CmdEscuchar.Enabled = True
End Sub

