VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WorldEditor"
   ClientHeight    =   13260
   ClientLeft      =   390
   ClientTop       =   840
   ClientWidth     =   21000
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   884
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1400
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   9
      Left            =   6240
      TabIndex        =   120
      Top             =   30
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   1826
      Caption         =   "SPOTS"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   65535
      LockHover       =   3
      cGradient       =   65535
      Gradient        =   1
      CapStyle        =   2
      Mode            =   1
      Value           =   0   'False
      CustomClick     =   1
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   0
      Left            =   4320
      TabIndex        =   23
      Top             =   30
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1826
      Caption         =   "&Superficie (F5)"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   0
      cFHover         =   0
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      ImgAlign        =   5
      Image           =   "frmMain.frx":000C
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin VB.PictureBox pPaneles 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000040&
      ForeColor       =   &H80000008&
      Height          =   4395
      Left            =   45
      ScaleHeight     =   4365
      ScaleWidth      =   4365
      TabIndex        =   3
      Top             =   1920
      Width           =   4395
      Begin VB.CheckBox chkDecorBloq 
         BackColor       =   &H00808080&
         Caption         =   "Bloquear Decor"
         Height          =   255
         Left            =   480
         MaskColor       =   &H00808080&
         TabIndex        =   174
         Top             =   3600
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.ListBox lListado 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   2460
         Index           =   8
         ItemData        =   "frmMain.frx":3552
         Left            =   120
         List            =   "frmMain.frx":3554
         Style           =   1  'Checkbox
         TabIndex        =   173
         Tag             =   "-1"
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.ListBox lListado 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   2580
         Index           =   7
         ItemData        =   "frmMain.frx":3556
         Left            =   120
         List            =   "frmMain.frx":3558
         TabIndex        =   171
         Tag             =   "-1"
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
      End
      Begin WorldEditor.lvButtons_H cInsertarFunc 
         Height          =   375
         Index           =   2
         Left            =   2400
         TabIndex        =   49
         Top             =   3480
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Caption         =   "&Insertar Objetos"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.ComboBox cCantFunc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   2
         ItemData        =   "frmMain.frx":355A
         Left            =   960
         List            =   "frmMain.frx":355C
         TabIndex        =   108
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ListBox lListado 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   2580
         Index           =   6
         ItemData        =   "frmMain.frx":355E
         Left            =   120
         List            =   "frmMain.frx":3560
         TabIndex        =   170
         Tag             =   "-1"
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
      End
      Begin WorldEditor.lvButtons_H decorb 
         Height          =   375
         Left            =   2400
         TabIndex        =   169
         Top             =   3840
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Caption         =   "Decor"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.CheckBox cGrill 
         BackColor       =   &H00000040&
         Caption         =   "Mostrar Grilla"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2400
         TabIndex        =   168
         Top             =   3960
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.ListBox lListado 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   2160
         Index           =   5
         ItemData        =   "frmMain.frx":3562
         Left            =   120
         List            =   "frmMain.frx":3564
         TabIndex        =   167
         Tag             =   "-1"
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.ListBox lListado 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   2160
         Index           =   0
         ItemData        =   "frmMain.frx":3566
         Left            =   120
         List            =   "frmMain.frx":3568
         Sorted          =   -1  'True
         TabIndex        =   50
         Tag             =   "-1"
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
      End
      Begin WorldEditor.lvButtons_H bI 
         Height          =   375
         Left            =   120
         TabIndex        =   166
         Top             =   2280
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         Caption         =   "Index"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.TextBox tTY 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Left            =   1200
         TabIndex        =   65
         Text            =   "1"
         Top             =   960
         Visible         =   0   'False
         Width           =   2900
      End
      Begin VB.TextBox tTX 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Left            =   1200
         TabIndex        =   64
         Text            =   "1"
         Top             =   600
         Visible         =   0   'False
         Width           =   2900
      End
      Begin VB.TextBox tTMapa 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Left            =   1200
         TabIndex        =   63
         Text            =   "1"
         Top             =   240
         Visible         =   0   'False
         Width           =   2900
      End
      Begin WorldEditor.lvButtons_H cInsertarTrans 
         Height          =   375
         Left            =   240
         TabIndex        =   66
         Top             =   1320
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         Caption         =   "&Insertar Translado"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cInsertarTransOBJ 
         Height          =   375
         Left            =   240
         TabIndex        =   67
         Top             =   1680
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         Caption         =   "Colocar automaticamente &Objeto"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cUnionManual 
         Height          =   375
         Left            =   240
         TabIndex        =   68
         Top             =   2160
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         Caption         =   "&Union con Mapa Adyacente (manual)"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.ComboBox cCapas 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         ItemData        =   "frmMain.frx":356A
         Left            =   1080
         List            =   "frmMain.frx":3580
         TabIndex        =   0
         TabStop         =   0   'False
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox cGrh 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         ItemData        =   "frmMain.frx":3596
         Left            =   2880
         List            =   "frmMain.frx":3598
         TabIndex        =   52
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox cFiltro 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   0
         Left            =   600
         TabIndex        =   51
         Top             =   2760
         Visible         =   0   'False
         Width           =   3615
      End
      Begin WorldEditor.lvButtons_H cQuitarEnTodasLasCapas 
         Height          =   375
         Left            =   120
         TabIndex        =   53
         Top             =   3840
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "Quitar en &Capas 2 y 3"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cQuitarEnEstaCapa 
         Height          =   375
         Left            =   120
         TabIndex        =   54
         Top             =   3480
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "&Quitar en esta Capa"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cSeleccionarSuperficie 
         Height          =   495
         Left            =   2400
         TabIndex        =   55
         Top             =   3480
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         Caption         =   "&Insertar Superficie"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.ComboBox cNumFunc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   2
         ItemData        =   "frmMain.frx":359A
         Left            =   3360
         List            =   "frmMain.frx":359C
         TabIndex        =   46
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox cFiltro 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   3
         Left            =   600
         TabIndex        =   44
         Top             =   2760
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.ComboBox cNumFunc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   0
         ItemData        =   "frmMain.frx":359E
         Left            =   3360
         List            =   "frmMain.frx":35A0
         TabIndex        =   36
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox cFiltro 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   1
         Left            =   600
         TabIndex        =   35
         Top             =   2760
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.ListBox lListado 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   2580
         Index           =   1
         ItemData        =   "frmMain.frx":35A2
         Left            =   120
         List            =   "frmMain.frx":35A4
         TabIndex        =   34
         Tag             =   "-1"
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.ListBox lListado 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   3180
         Index           =   4
         ItemData        =   "frmMain.frx":35A6
         Left            =   120
         List            =   "frmMain.frx":35A8
         Style           =   1  'Checkbox
         TabIndex        =   33
         Tag             =   "-1"
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.PictureBox Picture5 
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   5
         Top             =   0
         Width           =   0
      End
      Begin VB.PictureBox Picture6 
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   6
         Top             =   0
         Width           =   0
      End
      Begin VB.PictureBox Picture7 
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   7
         Top             =   0
         Width           =   0
      End
      Begin VB.PictureBox Picture8 
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   8
         Top             =   0
         Width           =   0
      End
      Begin VB.PictureBox Picture9 
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   9
         Top             =   0
         Width           =   0
      End
      Begin VB.PictureBox Picture11 
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   29
         Top             =   0
         Width           =   0
      End
      Begin WorldEditor.lvButtons_H cQuitarTrigger 
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   3840
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "&Quitar Trigger's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cVerTriggers 
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   3480
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "&Mostrar Trigger's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cInsertarTrigger 
         Height          =   375
         Left            =   2400
         TabIndex        =   32
         Top             =   3480
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Caption         =   "&Insertar Trigger"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cAgregarFuncalAzar 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   38
         Top             =   3480
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "Insetar NPC's al &Azar"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cQuitarFunc 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   39
         Top             =   3840
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "&Quitar NPC's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cInsertarFunc 
         Height          =   375
         Index           =   0
         Left            =   2400
         TabIndex        =   40
         Top             =   3480
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Caption         =   "&Insertar NPC's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cVerBloqueos 
         Height          =   495
         Left            =   120
         TabIndex        =   41
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   873
         Caption         =   "&Mostrar Bloqueos"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cInsertarBloqueo 
         Height          =   735
         Left            =   120
         TabIndex        =   42
         Top             =   720
         Visible         =   0   'False
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   1296
         Caption         =   "&Insertar Bloqueos"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cQuitarBloqueo 
         Height          =   735
         Left            =   120
         TabIndex        =   43
         Top             =   1560
         Visible         =   0   'False
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   1296
         Caption         =   "&Quitar Bloqueos"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cAgregarFuncalAzar 
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   47
         Top             =   3480
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "Insetar OBJ's al &Azar"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cQuitarFunc 
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   48
         Top             =   3840
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "&Quitar OBJ's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cInsertarFunc 
         Height          =   375
         Index           =   1
         Left            =   2400
         TabIndex        =   62
         Top             =   3480
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Caption         =   "&Insertar NPC's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cQuitarFunc 
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   61
         Top             =   3840
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "&Quitar NPC's"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cAgregarFuncalAzar 
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   60
         Top             =   3480
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "Insetar NPC's al &Azar"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.ComboBox cCantFunc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   1
         ItemData        =   "frmMain.frx":35AA
         Left            =   840
         List            =   "frmMain.frx":35AC
         TabIndex        =   56
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ComboBox cFiltro 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   2
         Left            =   600
         TabIndex        =   57
         Top             =   2760
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.ListBox lListado 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   2580
         Index           =   2
         ItemData        =   "frmMain.frx":35AE
         Left            =   120
         List            =   "frmMain.frx":35B0
         TabIndex        =   58
         Tag             =   "-1"
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.ComboBox cNumFunc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   1
         ItemData        =   "frmMain.frx":35B2
         Left            =   3360
         List            =   "frmMain.frx":35B4
         TabIndex        =   59
         Text            =   "500"
         Top             =   3120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Frame cParticulas 
         BackColor       =   &H80000007&
         Caption         =   "Particles"
         ForeColor       =   &H80000009&
         Height          =   3105
         Left            =   720
         TabIndex        =   88
         Top             =   360
         Visible         =   0   'False
         Width           =   3180
         Begin VB.CheckBox ChkInterior 
            Caption         =   "Ver Interiores"
            Height          =   195
            Left            =   1560
            TabIndex        =   119
            Top             =   1200
            Width           =   1455
         End
         Begin VB.CheckBox chkParticle 
            Caption         =   "Ver Particulas"
            Height          =   195
            Left            =   1560
            TabIndex        =   118
            Top             =   840
            Width           =   1455
         End
         Begin VB.TextBox txtInterior 
            Height          =   285
            Left            =   120
            TabIndex        =   117
            Text            =   "1"
            Top             =   2760
            Width           =   735
         End
         Begin VB.TextBox txtParticula 
            BackColor       =   &H80000006&
            ForeColor       =   &H80000005&
            Height          =   375
            Left            =   900
            TabIndex        =   89
            Text            =   "1"
            Top             =   255
            Width           =   555
         End
         Begin WorldEditor.lvButtons_H cInsertarParticula 
            Height          =   405
            Left            =   75
            TabIndex        =   90
            Top             =   765
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   714
            Caption         =   "Insertar Particula"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   1
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin WorldEditor.lvButtons_H cQuitarParticula 
            Height          =   390
            Left            =   75
            TabIndex        =   91
            Top             =   1215
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   688
            Caption         =   "Quitar Particula"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   1
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin WorldEditor.lvButtons_H CmdInteriorI 
            Height          =   405
            Left            =   120
            TabIndex        =   115
            Top             =   1680
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   714
            Caption         =   "Insertar Interior"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   1
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin WorldEditor.lvButtons_H CmdInteriorQ 
            Height          =   405
            Left            =   120
            TabIndex        =   116
            Top             =   2160
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   714
            Caption         =   "Quitar Interior"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   1
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin VB.Label Label1 
            BackColor       =   &H80000012&
            Caption         =   "Particula:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   150
            TabIndex        =   92
            Top             =   300
            Width           =   660
         End
      End
      Begin VB.Frame cLuces 
         BackColor       =   &H80000012&
         Caption         =   "Luces"
         ForeColor       =   &H80000009&
         Height          =   3645
         Left            =   165
         TabIndex        =   78
         Top             =   120
         Visible         =   0   'False
         Width           =   3975
         Begin VB.TextBox Combo1 
            Height          =   285
            Left            =   2520
            TabIndex        =   114
            Text            =   "14"
            Top             =   240
            Width           =   1215
         End
         Begin VB.CheckBox cVerLuces 
            BackColor       =   &H80000012&
            Caption         =   "Ver luces"
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   2400
            TabIndex        =   98
            Top             =   2400
            Width           =   1515
         End
         Begin VB.CheckBox cBorde 
            BackColor       =   &H80000012&
            Caption         =   "Bordeado"
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   2400
            TabIndex        =   96
            Top             =   2640
            Width           =   1515
         End
         Begin VB.TextBox tLuz 
            Height          =   285
            Left            =   2160
            TabIndex        =   95
            Text            =   "0"
            Top             =   2040
            Width           =   1575
         End
         Begin VB.ListBox lLuces 
            Height          =   1425
            ItemData        =   "frmMain.frx":35B6
            Left            =   2160
            List            =   "frmMain.frx":35ED
            TabIndex        =   94
            Top             =   600
            Width           =   1575
         End
         Begin VB.CheckBox LuzRedonda 
            BackColor       =   &H80000012&
            Caption         =   "Luces redondas"
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   2400
            TabIndex        =   87
            Top             =   2880
            Width           =   1515
         End
         Begin VB.Frame RGBCOLOR 
            BackColor       =   &H80000012&
            Caption         =   "RGB"
            ForeColor       =   &H00FFFFFF&
            Height          =   690
            Left            =   135
            TabIndex        =   81
            Top             =   120
            Width           =   1680
            Begin VB.TextBox G 
               BackColor       =   &H80000012&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000014&
               Height          =   315
               Left            =   600
               TabIndex        =   84
               Text            =   "1"
               Top             =   270
               Width           =   450
            End
            Begin VB.TextBox B 
               BackColor       =   &H80000012&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000014&
               Height          =   315
               Left            =   1095
               TabIndex        =   83
               Text            =   "1"
               Top             =   270
               Width           =   450
            End
            Begin VB.TextBox R 
               BackColor       =   &H80000012&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000014&
               Height          =   315
               Left            =   105
               TabIndex        =   82
               Text            =   "1"
               Top             =   270
               Width           =   450
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H80000012&
            Caption         =   "Rango"
            ForeColor       =   &H8000000E&
            Height          =   660
            Left            =   135
            TabIndex        =   79
            Top             =   840
            Width           =   1695
            Begin VB.TextBox cRango 
               BackColor       =   &H80000012&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000014&
               Height          =   315
               Left            =   105
               TabIndex        =   80
               Text            =   "0"
               Top             =   240
               Width           =   1035
            End
         End
         Begin WorldEditor.lvButtons_H cInsertarLuz 
            Height          =   360
            Left            =   120
            TabIndex        =   85
            Top             =   1920
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   635
            Caption         =   "Insertar Luz"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   1
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin WorldEditor.lvButtons_H cQuitarLuz 
            Height          =   360
            Left            =   120
            TabIndex        =   86
            Top             =   1560
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   635
            Caption         =   "Quitar Luz"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   1
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin WorldEditor.lvButtons_H cInsertarBorde 
            Height          =   360
            Left            =   120
            TabIndex        =   97
            Top             =   2280
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   635
            Caption         =   "Insertar Borde"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   1
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin WorldEditor.lvButtons_H cVertical 
            Height          =   360
            Left            =   120
            TabIndex        =   99
            Top             =   2760
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   635
            Caption         =   "|"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBhover         =   8454016
            LockHover       =   1
            cGradient       =   0
            Mode            =   1
            Value           =   0   'False
            cBack           =   8421631
         End
         Begin WorldEditor.lvButtons_H cHorizontal 
            Height          =   360
            Left            =   480
            TabIndex        =   100
            Top             =   2760
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   635
            Caption         =   "----"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBhover         =   8454016
            LockHover       =   1
            cGradient       =   0
            Mode            =   1
            Value           =   0   'False
            cBack           =   8421631
         End
         Begin WorldEditor.lvButtons_H cUL 
            Height          =   360
            Left            =   840
            TabIndex        =   101
            Top             =   2760
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   635
            Caption         =   "°"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBhover         =   8454016
            LockHover       =   1
            cGradient       =   0
            Mode            =   1
            Value           =   0   'False
            cBack           =   8421631
         End
         Begin WorldEditor.lvButtons_H cUR 
            Height          =   360
            Left            =   1200
            TabIndex        =   102
            Top             =   2760
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   635
            Caption         =   "°"
            CapAlign        =   1
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBhover         =   8454016
            LockHover       =   1
            cGradient       =   0
            Mode            =   1
            Value           =   0   'False
            cBack           =   8421631
         End
         Begin WorldEditor.lvButtons_H cBL 
            Height          =   360
            Left            =   840
            TabIndex        =   103
            Top             =   3120
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   635
            Caption         =   "."
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBhover         =   8454016
            LockHover       =   1
            cGradient       =   0
            Mode            =   1
            Value           =   0   'False
            cBack           =   8421631
         End
         Begin WorldEditor.lvButtons_H cBR 
            Height          =   360
            Left            =   1200
            TabIndex        =   104
            Top             =   3120
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   635
            Caption         =   "."
            CapAlign        =   1
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBhover         =   8454016
            LockHover       =   1
            cGradient       =   0
            Mode            =   1
            Value           =   0   'False
            cBack           =   8421631
         End
         Begin WorldEditor.lvButtons_H cCROSSUR 
            Height          =   360
            Left            =   480
            TabIndex        =   105
            Top             =   3120
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   635
            Caption         =   ".  °"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBhover         =   8454016
            LockHover       =   1
            cGradient       =   0
            Mode            =   1
            Value           =   0   'False
            cBack           =   8421631
         End
         Begin WorldEditor.lvButtons_H cCROSSUL 
            Height          =   360
            Left            =   120
            TabIndex        =   106
            Top             =   3120
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   635
            Caption         =   "°  ."
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBhover         =   8454016
            LockHover       =   1
            cGradient       =   0
            Mode            =   1
            Value           =   0   'False
            cBack           =   8421631
         End
         Begin WorldEditor.lvButtons_H cALLC 
            Height          =   360
            Left            =   1560
            TabIndex        =   107
            Top             =   2760
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   635
            Caption         =   ": :"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBhover         =   8454016
            LockHover       =   1
            cGradient       =   0
            Mode            =   1
            Value           =   0   'False
            cBack           =   8421631
         End
         Begin WorldEditor.lvButtons_H cNotUL 
            Height          =   360
            Left            =   1560
            TabIndex        =   109
            Top             =   3120
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   635
            Caption         =   ".  :"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBhover         =   8454016
            LockHover       =   1
            cGradient       =   0
            Mode            =   1
            Value           =   0   'False
            cBack           =   8421631
         End
         Begin WorldEditor.lvButtons_H cNotUR 
            Height          =   360
            Left            =   1920
            TabIndex        =   110
            Top             =   2760
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   635
            Caption         =   ":  ."
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBhover         =   8454016
            LockHover       =   1
            cGradient       =   0
            Mode            =   1
            Value           =   0   'False
            cBack           =   8421631
         End
         Begin WorldEditor.lvButtons_H cNotBL 
            Height          =   360
            Left            =   1920
            TabIndex        =   111
            Top             =   3120
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   635
            Caption         =   "°  :"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBhover         =   8454016
            LockHover       =   1
            cGradient       =   0
            Mode            =   1
            Value           =   0   'False
            cBack           =   8421631
         End
         Begin WorldEditor.lvButtons_H cNotBR 
            Height          =   360
            Left            =   2280
            TabIndex        =   112
            Top             =   3120
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   635
            Caption         =   ":  °"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBhover         =   8454016
            LockHover       =   1
            cGradient       =   0
            Mode            =   1
            Value           =   0   'False
            cBack           =   8421631
         End
         Begin WorldEditor.lvButtons_H cINV 
            Height          =   360
            Left            =   2760
            TabIndex        =   113
            Top             =   3120
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   635
            Caption         =   "INVERTIR"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBhover         =   8454016
            LockHover       =   1
            cGradient       =   0
            Mode            =   1
            Value           =   0   'False
            cBack           =   8421631
         End
      End
      Begin WorldEditor.lvButtons_H cQuitarTrans 
         Height          =   375
         Left            =   240
         TabIndex        =   70
         Top             =   3000
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         Caption         =   "&Quitar Translados"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   1
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cUnionAuto 
         Height          =   375
         Left            =   240
         TabIndex        =   69
         Top             =   2520
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         Caption         =   "Union con Mapas &Adyacentes (auto)"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.ComboBox cCantFunc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   330
         Index           =   0
         ItemData        =   "frmMain.frx":36F0
         Left            =   840
         List            =   "frmMain.frx":36F2
         TabIndex        =   37
         Text            =   "1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ListBox lListado 
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   2580
         Index           =   3
         ItemData        =   "frmMain.frx":36F4
         Left            =   120
         List            =   "frmMain.frx":36F6
         TabIndex        =   45
         Tag             =   "-1"
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.Frame frmSPOTLIGHTS 
         BackColor       =   &H00000000&
         Height          =   4335
         Left            =   0
         TabIndex        =   121
         Top             =   0
         Visible         =   0   'False
         Width           =   3975
         Begin VB.CheckBox MarcarsPOT 
            Caption         =   "Marcar SPOTS"
            Height          =   195
            Left            =   1560
            TabIndex        =   146
            Top             =   240
            Width           =   1815
         End
         Begin VB.TextBox COLOR_CUSTOM_EXTRA 
            Height          =   285
            Left            =   2280
            TabIndex        =   142
            Top             =   3555
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.TextBox COLOR_CUSTOM_SPOT 
            Height          =   285
            Left            =   2280
            TabIndex        =   141
            Top             =   3240
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.ComboBox SPOT_ANIM 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   130
            Top             =   480
            Width           =   2055
         End
         Begin VB.ComboBox COLORSPOT 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   129
            Top             =   1200
            Width           =   2055
         End
         Begin VB.ComboBox COLOREXTRA 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   128
            Top             =   1560
            Width           =   2055
         End
         Begin VB.TextBox GRAFICO_SPOT 
            Height          =   285
            Left            =   2280
            TabIndex        =   127
            Text            =   "43"
            Top             =   840
            Width           =   1335
         End
         Begin VB.TextBox SPOT_INTENSIDAD 
            Height          =   285
            Left            =   2280
            TabIndex        =   126
            Text            =   "1"
            Top             =   2250
            Width           =   1335
         End
         Begin VB.TextBox GRAFICO_SPOT_COLOR 
            Height          =   285
            Left            =   2280
            TabIndex        =   125
            Top             =   1920
            Width           =   1335
         End
         Begin VB.TextBox SPOT_OFFSETX 
            Height          =   285
            Left            =   2280
            TabIndex        =   124
            Top             =   2570
            Width           =   1335
         End
         Begin VB.TextBox SPOT_OFFSETY 
            Height          =   285
            Left            =   2280
            TabIndex        =   123
            Top             =   2880
            Width           =   1335
         End
         Begin WorldEditor.lvButtons_H QUITARSPOT 
            Height          =   375
            Left            =   120
            TabIndex        =   122
            Top             =   3840
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            Caption         =   "QUITAR SPOT"
            CapAlign        =   2
            BackStyle       =   3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBhover         =   12632064
            cGradient       =   12632064
            Gradient        =   1
            Mode            =   1
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin WorldEditor.lvButtons_H PONERSPOT 
            Height          =   375
            Left            =   1440
            TabIndex        =   131
            Top             =   3840
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            Caption         =   "PONER SPOT"
            CapAlign        =   2
            BackStyle       =   3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBhover         =   12632064
            cGradient       =   12632064
            Gradient        =   1
            Mode            =   1
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin WorldEditor.lvButtons_H SPOTEDITAR 
            Height          =   375
            Left            =   2640
            TabIndex        =   145
            Top             =   3840
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            Caption         =   "EDITAR"
            CapAlign        =   2
            BackStyle       =   3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBhover         =   12632064
            cGradient       =   12632064
            Gradient        =   1
            Mode            =   0
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Color custom extra:"
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   240
            TabIndex        =   144
            Top             =   3555
            Visible         =   0   'False
            Width           =   1740
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Color custom:"
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   240
            TabIndex        =   143
            Top             =   3315
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Animacion:"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   240
            TabIndex        =   140
            Top             =   600
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Color SPOT:"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   240
            TabIndex        =   139
            Top             =   1320
            Width           =   1140
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Color adicional:"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   240
            TabIndex        =   138
            Top             =   1680
            Width           =   660
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Grafico luz:"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   240
            TabIndex        =   137
            Top             =   960
            Width           =   795
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Grafico del color:"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   0
            TabIndex        =   136
            Top             =   0
            Width           =   660
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Grafico del color:"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   240
            TabIndex        =   135
            Top             =   2040
            Width           =   1200
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Intensidad:"
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   240
            TabIndex        =   134
            Top             =   2310
            Width           =   1020
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Offset X:"
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   240
            TabIndex        =   133
            Top             =   2640
            Width           =   1020
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Offset Y:"
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   240
            TabIndex        =   132
            Top             =   2880
            Width           =   1020
         End
      End
      Begin VB.Label lYver 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Y vertical:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Left            =   240
         TabIndex        =   73
         Top             =   1005
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lXhor 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "X horizontal:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Left            =   240
         TabIndex        =   72
         Top             =   645
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label lMapN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Mapa:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Left            =   240
         TabIndex        =   71
         Top             =   285
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Label lbFiltrar 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Filtrar:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   2820
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label lbCapas 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Capa Actual:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Left            =   120
         TabIndex        =   20
         Top             =   3195
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Label lbGrh 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Sup Actual:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Left            =   2040
         TabIndex        =   19
         Top             =   3195
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label lNumFunc 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Numero de NPC:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   1
         Left            =   2160
         TabIndex        =   18
         Top             =   3195
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label lCantFunc 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Cantidad:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   3195
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lbFiltrar 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Filtrar:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   2820
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label lNumFunc 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Numero de OBJ:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   2
         Left            =   2160
         TabIndex        =   15
         Top             =   3195
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label lCantFunc 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Cantidad:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   3195
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lbFiltrar 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Filtrar:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   3
         Left            =   120
         TabIndex        =   13
         Top             =   2820
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label lCantFunc 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Cantidad:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   3195
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lNumFunc 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Numero de NPC:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   0
         Left            =   2160
         TabIndex        =   11
         Top             =   3195
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label lbFiltrar 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Filtrar:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   2820
         Visible         =   0   'False
         Width           =   450
      End
   End
   Begin WorldEditor.lvButtons_H VistaStat 
      Height          =   3810
      Left            =   3975
      TabIndex        =   172
      Top             =   6315
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   6720
      Caption         =   "Vista Previa"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.PictureBox picRadar 
      BackColor       =   &H00000040&
      BorderStyle     =   0  'None
      Height          =   1590
      Left            =   120
      ScaleHeight     =   106
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   107
      TabIndex        =   93
      Top             =   120
      Width           =   1605
      Begin VB.Shape ApuntadorRadar 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   6  'Inside Solid
         DrawMode        =   6  'Mask Pen Not
         FillColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   600
         Top             =   600
         Width           =   375
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000000&
         Height          =   1365
         Left            =   120
         Top             =   105
         Width           =   1365
      End
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   8
      Left            =   14520
      TabIndex        =   77
      Top             =   30
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   1826
      Caption         =   "Particula e Interiores"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      CustomClick     =   1
      ImgAlign        =   5
      Image           =   "frmMain.frx":36F8
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin VB.PictureBox MainViewPic 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   11520
      Left            =   4560
      ScaleHeight     =   768
      ScaleMode       =   0  'User
      ScaleWidth      =   1024
      TabIndex        =   75
      Top             =   1440
      Width           =   15360
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   6
      Left            =   10995
      TabIndex        =   28
      Top             =   30
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1826
      Caption         =   "Tri&gger's (F12)"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      CustomClick     =   1
      ImgAlign        =   5
      Image           =   "frmMain.frx":3CBE
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   5
      Left            =   9840
      TabIndex        =   27
      Top             =   30
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   1826
      Caption         =   "&Objetos (F11)"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      CustomClick     =   1
      ImgAlign        =   5
      Image           =   "frmMain.frx":4284
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   3
      Left            =   8640
      TabIndex        =   26
      Top             =   30
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1826
      Caption         =   "&NPC's (F8)"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      CustomClick     =   1
      ImgAlign        =   5
      Image           =   "frmMain.frx":4785
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   2
      Left            =   7200
      TabIndex        =   25
      Top             =   30
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   1826
      Caption         =   "&Bloqueos (F7)"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      CustomClick     =   1
      ImgAlign        =   5
      Image           =   "frmMain.frx":4B39
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin VB.Timer TimAutoGuardarMapa 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   3960
      Top             =   1920
   End
   Begin VB.TextBox StatTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3825
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "frmMain.frx":4EBA
      Top             =   6300
      Width           =   3825
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   5460
      Left            =   60
      ScaleHeight     =   364
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   297
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   6270
      Width           =   4455
      Begin VB.PictureBox PreviewGrh 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         FillColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   3840
         Left            =   45
         ScaleHeight     =   3840
         ScaleWidth      =   3840
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   3840
      End
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   2565
      Top             =   2025
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   675
      Index           =   4
      Left            =   9840
      TabIndex        =   74
      Top             =   240
      Visible         =   0   'False
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   1191
      Caption         =   "none"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      CustomClick     =   1
      ImgAlign        =   5
      Image           =   "frmMain.frx":4EFA
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   7
      Left            =   12240
      TabIndex        =   76
      Top             =   30
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1826
      Caption         =   "Luces "
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      CustomClick     =   1
      ImgAlign        =   5
      Image           =   "frmMain.frx":52AE
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   10
      Left            =   13440
      TabIndex        =   148
      Top             =   30
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   1826
      Caption         =   "NW"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      ImgAlign        =   5
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin VB.Frame MpNw 
      BackColor       =   &H00000000&
      Caption         =   "Mapas graficos"
      ForeColor       =   &H00FFFFFF&
      Height          =   8415
      Left            =   0
      TabIndex        =   147
      Top             =   1920
      Visible         =   0   'False
      Width           =   4455
      Begin VB.CheckBox cTipoTerreno 
         BackColor       =   &H00000000&
         Caption         =   "ver tipo terreno"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1920
         TabIndex        =   165
         Top             =   6360
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00000000&
         Caption         =   "Costa"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   163
         Top             =   6360
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00000000&
         Caption         =   "Lava"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   162
         Top             =   6120
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00000000&
         Caption         =   "Agua (Lago)"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   161
         Top             =   5880
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00000000&
         Caption         =   "Normal"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   160
         Top             =   5640
         Width           =   1335
      End
      Begin VB.CheckBox ccursor 
         BackColor       =   &H00000000&
         Caption         =   "Ver cursor"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   360
         TabIndex        =   159
         Top             =   3000
         Width           =   1455
      End
      Begin VB.TextBox txtNumIndice 
         Height          =   495
         Left            =   360
         TabIndex        =   157
         Top             =   2400
         Width           =   1455
      End
      Begin VB.PictureBox picsur 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1920
         Left            =   120
         ScaleHeight     =   128
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   128
         TabIndex        =   156
         Top             =   3360
         Width           =   1920
      End
      Begin VB.CheckBox cVerIndices 
         BackColor       =   &H00000000&
         Caption         =   "Ver indices"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   480
         TabIndex        =   154
         Top             =   2040
         Width           =   1455
      End
      Begin WorldEditor.lvButtons_H cInsertarSurface 
         Height          =   495
         Left            =   2160
         TabIndex        =   152
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         Caption         =   "Insertar"
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
         cGradient       =   0
         Mode            =   2
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.ComboBox SizeC 
         Height          =   315
         ItemData        =   "frmMain.frx":5874
         Left            =   360
         List            =   "frmMain.frx":5887
         TabIndex        =   151
         Top             =   1560
         Width           =   1455
      End
      Begin VB.ComboBox LayerC 
         Height          =   315
         ItemData        =   "frmMain.frx":58A2
         Left            =   360
         List            =   "frmMain.frx":58B5
         TabIndex        =   150
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox txtNumSurface 
         Height          =   495
         Left            =   360
         TabIndex        =   149
         Top             =   600
         Width           =   1455
      End
      Begin WorldEditor.lvButtons_H cBorrarSurface 
         Height          =   495
         Left            =   2160
         TabIndex        =   153
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         Caption         =   "Borrar"
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
         cGradient       =   0
         Mode            =   2
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cBorrarSobrante 
         Height          =   495
         Left            =   2160
         TabIndex        =   155
         Top             =   1800
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         Caption         =   "Borrar Sobrante"
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
         cGradient       =   0
         Mode            =   2
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cEditarIndice 
         Height          =   495
         Left            =   2160
         TabIndex        =   158
         Top             =   2400
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         Caption         =   "Editar Indice"
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
         cGradient       =   0
         Mode            =   2
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin WorldEditor.lvButtons_H cAplicarTerreno 
         Height          =   615
         Left            =   1920
         TabIndex        =   164
         Top             =   5640
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1085
         Caption         =   "Aplicar Tipo Terreno"
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
         cGradient       =   0
         Mode            =   2
         Value           =   0   'False
         cBack           =   -2147483633
      End
   End
   Begin WorldEditor.lvButtons_H SelectPanel 
      Height          =   1035
      Index           =   1
      Left            =   4680
      TabIndex        =   24
      Top             =   30
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   1826
      Caption         =   "&Translados   (F6)"
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LockHover       =   1
      cGradient       =   8421631
      Mode            =   1
      Value           =   0   'False
      ImgAlign        =   5
      Image           =   "frmMain.frx":58E6
      ImgSize         =   24
      cBack           =   -2147483633
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   1680
      TabIndex        =   176
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   1800
      TabIndex        =   175
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label MapPest 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mapa1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   4560
      TabIndex        =   22
      Top             =   1080
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Menu FileMnu 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuArchivoLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNuevoMapa 
         Caption         =   "&Nuevo Mapa"
         Shortcut        =   ^N
      End
      Begin VB.Menu OpMpGr 
         Caption         =   "&Abrir Mapa Grafico"
      End
      Begin VB.Menu mnuAbrirMapa 
         Caption         =   "&Abrir Mapa"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuArchivoLine2 
         Caption         =   "-"
      End
      Begin VB.Menu chkGuardarInf 
         Caption         =   "Guardar .Inf"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuGuardarMapa 
         Caption         =   "&Guardar Mapa"
         Shortcut        =   ^G
      End
      Begin VB.Menu grdNuevoMapa 
         Caption         =   "Guardar Mapa como"
      End
      Begin VB.Menu mnuArchivoLine5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGuardarcomoBMP 
         Caption         =   "Guardar Render en MiniMapa"
      End
      Begin VB.Menu mnuGuardarcomoJPG 
         Caption         =   "Guardar Render 3200x3200"
      End
      Begin VB.Menu mnuArchivoLine7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
      End
      Begin VB.Menu mnuArchivoLine6 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuEdicion 
      Caption         =   "&Edición"
      Begin VB.Menu mnuComo 
         Caption         =   "¿ Como seleccionar ? ---- Mantener SHIFT y arrastrar el cursor."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuCortar 
         Caption         =   "C&ortar Selección"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopiarOld 
         Caption         =   "&Copiar Selección"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuCopiar 
         Caption         =   "&Copiar Selección sin traslados"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPegar 
         Caption         =   "&Pegar Selección"
         Shortcut        =   ^V
      End
      Begin VB.Menu cInterio 
         Caption         =   "Copiar Interiores"
         Shortcut        =   ^I
      End
      Begin VB.Menu pInter 
         Caption         =   "Pegar Interiores"
         Shortcut        =   ^K
      End
      Begin VB.Menu cCopiarLuces 
         Caption         =   "Copiar luces"
      End
      Begin VB.Menu cPegarLuces 
         Caption         =   "Pegar Luces"
      End
      Begin VB.Menu mnuBloquearS 
         Caption         =   "&Bloquear Selección"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuRealizarOperacion 
         Caption         =   "&Realizar Operación en Selección"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuDeshacerPegado 
         Caption         =   "Deshacer P&egado de Selección"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuLineEdicion0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeshacer 
         Caption         =   "&Deshacer"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuUtilizarDeshacer 
         Caption         =   "&Utilizar Deshacer"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuInfoMap 
         Caption         =   "&Información del Mapa"
      End
      Begin VB.Menu mnuLineEdicion1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInsertar 
         Caption         =   "&Insertar"
         Begin VB.Menu mnuInsertarTransladosAdyasentes 
            Caption         =   "&Translados a Mapas Adyasentes"
         End
         Begin VB.Menu mnuInsertarSuperficieAlAzar 
            Caption         =   "Superficie al &Azar"
         End
         Begin VB.Menu mnuInsertarSuperficieEnBordes 
            Caption         =   "Superficie en los &Bordes del Mapa"
         End
         Begin VB.Menu mnuInsertarSuperficieEnTodo 
            Caption         =   "Superficie en Todo el Mapa"
         End
         Begin VB.Menu mnuBloquearBordes 
            Caption         =   "Bloqueo en &Bordes del Mapa"
         End
         Begin VB.Menu mnuBloquearMapa 
            Caption         =   "Bloqueo en &Todo el Mapa"
         End
      End
      Begin VB.Menu mnuQuitar 
         Caption         =   "&Quitar"
         Begin VB.Menu mnuQuitarTranslados 
            Caption         =   "Todos los &Translados"
         End
         Begin VB.Menu mnuQuitarBloqueos 
            Caption         =   "Todos los &Bloqueos"
         End
         Begin VB.Menu mnuQuitarNPCs 
            Caption         =   "Todos los &NPC's"
         End
         Begin VB.Menu mnuQuitarNPCsHostiles 
            Caption         =   "Todos los NPC's &Hostiles"
         End
         Begin VB.Menu mnuQuitarObjetos 
            Caption         =   "Todos los &Objetos"
         End
         Begin VB.Menu mnuQuitarTriggers 
            Caption         =   "Todos los Tri&gger's"
         End
         Begin VB.Menu mnuQuitarSuperficieBordes 
            Caption         =   "Superficie de los B&ordes"
         End
         Begin VB.Menu mnuQuitarSuperficieDeCapa 
            Caption         =   "Superficie de la &Capa Seleccionada"
         End
         Begin VB.Menu mnuLineEdicion2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuQuitarTODO 
            Caption         =   "TODO"
         End
      End
      Begin VB.Menu mnuLineEdicion3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFunciones 
         Caption         =   "&Funciones"
         Begin VB.Menu mnuQuitarFunciones 
            Caption         =   "&Quitar Funciones"
            Shortcut        =   {F4}
         End
         Begin VB.Menu mnuAutoQuitarFunciones 
            Caption         =   "Auto-&Quitar Funciones"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuConfigAvanzada 
         Caption         =   "Configuracion A&vanzada de Superficie"
      End
      Begin VB.Menu mnuLineEdicion4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAutoCompletarSuperficies 
         Caption         =   "Auto-Completar &Superficies"
      End
      Begin VB.Menu mnuAutoCapturarSuperficie 
         Caption         =   "Auto-C&apturar información de la Superficie"
      End
      Begin VB.Menu mnuAutoCapturarTranslados 
         Caption         =   "Auto-&Capturar información de los Translados"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuAutoGuardarMapas 
         Caption         =   "Configuración de Auto-&Guardar Mapas"
      End
   End
   Begin VB.Menu MnuVer 
      Caption         =   "&Ver"
      Begin VB.Menu mnuCapas 
         Caption         =   "...&Capas"
         Begin VB.Menu mnuVerCapa1 
            Caption         =   "Capa &1 (Piso)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuVerCapa2 
            Caption         =   "Capa &2 (agua - sobre piso)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuVerCapa3 
            Caption         =   "Capa &3 (arboles, etc)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuVerCapa4 
            Caption         =   "Capa &4 (techos, etc)"
         End
         Begin VB.Menu MnuVerCapa5 
            Caption         =   "Capa &5 (Sobre layer 1 y 2, previo 3)"
            Checked         =   -1  'True
         End
         Begin VB.Menu MnuVerCapa9 
            Caption         =   "Capa &9 (sobre Layer3)"
         End
      End
      Begin VB.Menu mnuVerTranslados 
         Caption         =   "...&Translados"
      End
      Begin VB.Menu mnuVerBloqueos 
         Caption         =   "...&Bloqueos"
      End
      Begin VB.Menu mnuVerNPCs 
         Caption         =   "...&NPC's"
      End
      Begin VB.Menu mVerDecors 
         Caption         =   "...&Decors"
      End
      Begin VB.Menu mnuVerObjetos 
         Caption         =   "...&Objetos"
      End
      Begin VB.Menu mnuVerTriggers 
         Caption         =   "...Tri&gger's"
      End
      Begin VB.Menu mnuVerGrilla 
         Caption         =   "...Gri&lla"
      End
      Begin VB.Menu mnuLinMostrar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVerAutomatico 
         Caption         =   "Control &Automaticamente"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuPaneles 
      Caption         =   "&Paneles"
      Begin VB.Menu mnuSuperficie 
         Caption         =   "&Superficie"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuTranslados 
         Caption         =   "&Translados"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuBloquear 
         Caption         =   "&Bloquear"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuNPCs 
         Caption         =   "&NPC's"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuNPCsHostiles 
         Caption         =   "NPC's &Hostiles"
         Shortcut        =   {F9}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuObjetos 
         Caption         =   "&Objetos"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuTriggers 
         Caption         =   "Tri&gger's"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuPanelesLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQSuperficie 
         Caption         =   "Ocultar Superficie"
         Shortcut        =   +{F5}
      End
      Begin VB.Menu mnuQTranslados 
         Caption         =   "Ocultar Translados"
         Shortcut        =   +{F6}
      End
      Begin VB.Menu mnuQBloquear 
         Caption         =   "Ocultar Bloquear"
         Shortcut        =   +{F7}
      End
      Begin VB.Menu mnuQNPCs 
         Caption         =   "Ocultar NPC's"
         Shortcut        =   +{F8}
      End
      Begin VB.Menu mnuQNPCsHostiles 
         Caption         =   "Ocultar NPC's Hostiles"
         Shortcut        =   +{F9}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuQObjetos 
         Caption         =   "Ocultar Objetos"
         Shortcut        =   +{F11}
      End
      Begin VB.Menu mnuQTriggers 
         Caption         =   "Ocultar Trigger's"
         Shortcut        =   +{F12}
      End
      Begin VB.Menu mnuFuncionesLine1 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnuActualizarIndices 
         Caption         =   "&Actualizar Indices de..."
         Begin VB.Menu mnuActualizarSuperficies 
            Caption         =   "&Superficies"
         End
         Begin VB.Menu mnuActualizarNPCs 
            Caption         =   "&NPC's"
         End
         Begin VB.Menu mnuActualizarObjs 
            Caption         =   "&Objetos"
         End
         Begin VB.Menu mnuActualizarTriggers 
            Caption         =   "&Trigger's"
         End
         Begin VB.Menu mnuActualizarGraficos 
            Caption         =   "&Graficos"
         End
      End
      Begin VB.Menu mnuModoCaminata 
         Caption         =   "Modalidad &Caminata"
      End
      Begin VB.Menu mnuGRHaBMP 
         Caption         =   "Texture Maker"
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptimizar 
         Caption         =   "Optimi&zar Mapa"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuLine2 
         Caption         =   "Extras"
      End
      Begin VB.Menu mnuGuardarUltimaConfig 
         Caption         =   "&Guardar Ultima Configuración"
      End
   End
   Begin VB.Menu mnuObjSc 
      Caption         =   "mnuObjSc"
      Visible         =   0   'False
      Begin VB.Menu mnuConfigObjTrans 
         Caption         =   "&Utilizar como Objeto de Translados"
      End
   End
   Begin VB.Menu mele_Main 
      Caption         =   "E&lementos"
      Begin VB.Menu mele_decor 
         Caption         =   "Editar Seleccion"
      End
   End
End
Attribute VB_Name = "frmMain"
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
Public Statz As Boolean

Private Sub bI_Click()
    If bI.value Then 'index button
        lListado(0).Visible = False 'list texturas
        lListado(5).Visible = True  'list index
    Else
        lListado(5).Visible = False
        lListado(0).Visible = True
    End If
End Sub



Private Sub cALLC_Click() 'luces
If cALLC.value Then
    'Activamos
    If cHorizontal.value Then
        cHorizontal.value = 0
    ElseIf cVertical.value Then
        cVertical.value = 0
    ElseIf cUR.value Then
        cUR.value = 0
    ElseIf cBL.value Then
        cBL.value = 0
    ElseIf cBR.value Then
         cBR.value = 0
    ElseIf cUL.value Then
         cUL.value = 0
    ElseIf cCROSSUR.value Then
        cCROSSUR.value = 0
    ElseIf cCROSSUL.value Then
        cCROSSUL.value = 0
    ElseIf cNotUL.value Then
        cNotUL.value = 0
    ElseIf cNotUR.value Then
        cNotUR.value = 0
    ElseIf cNotBL.value Then
        cNotBL.value = 0
    ElseIf cNotBR.value Then
        cNotBR.value = 0
    End If
Else
End If
End Sub

Private Sub cBL_Click() 'luces
If cBL.value Then
    'Activamos
    If cHorizontal.value Then
        cHorizontal.value = 0
    ElseIf cVertical.value Then
        cVertical.value = 0
    ElseIf cUR.value Then
        cUR.value = 0
    ElseIf cUL.value Then
        cUL.value = 0
    ElseIf cBR.value Then
         cBR.value = 0
    ElseIf cALLC.value Then
         cALLC.value = 0
    ElseIf cCROSSUR.value Then
        cCROSSUR.value = 0
    ElseIf cCROSSUL.value Then
        cCROSSUL.value = 0
    ElseIf cNotUL.value Then
        cNotUL.value = 0
    ElseIf cNotUR.value Then
        cNotUR.value = 0
    ElseIf cNotBL.value Then
        cNotBL.value = 0
    ElseIf cNotBR.value Then
        cNotBR.value = 0
    End If
Else
End If
End Sub

Private Sub cBR_Click()
If cBR.value Then
    'Activamos
    If cHorizontal.value Then
        cHorizontal.value = 0
    ElseIf cVertical.value Then
        cVertical.value = 0
    ElseIf cUR.value Then
        cUR.value = 0
    ElseIf cBL.value Then
        cBL.value = 0
    ElseIf cUL.value Then
         cUL.value = 0
    ElseIf cALLC.value Then
         cALLC.value = 0
    ElseIf cCROSSUR.value Then
        cCROSSUR.value = 0
    ElseIf cCROSSUL.value Then
        cCROSSUL.value = 0
    ElseIf cNotUL.value Then
        cNotUL.value = 0
    ElseIf cNotUR.value Then
        cNotUR.value = 0
    ElseIf cNotBL.value Then
        cNotBL.value = 0
    ElseIf cNotBR.value Then
        cNotBR.value = 0
    End If
Else
End If
End Sub

Private Sub cCantFunc_Change(index As Integer) 'editar cantidad objetos
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
On Error Resume Next
    If Val(cCantFunc(index)) < 1 Then
      cCantFunc(index).Text = 1
    End If
    If Val(cCantFunc(index)) > 10000 Then
      cCantFunc(index).Text = 10000
    End If
End Sub

Private Sub cCapas_Change()
'*************************************************
'Author: ^[GS]^
'Last modified: 31/05/06
'*************************************************
    If Val(cCapas.Text) < 1 Then
      cCapas.Text = 1
    End If
    If Val(cCapas.Text) > 9 Then
      cCapas.Text = 9
    End If
    If Val(cCapas.Text) > 5 And Val(cCapas.Text) < 9 Then
        cCapas.Text = 5
    End If
    cCapaSel = Val(cCapas.Text)
    cCapas.Tag = vbNullString
End Sub

Private Sub cCapas_Click()
    If Val(cCapas.List(cCapas.ListIndex)) > 0 And Val(cCapas.List(cCapas.ListIndex)) <= 5 Then
        cCapaSel = Val(cCapas.List(cCapas.ListIndex))
    ElseIf Val(cCapas.List(cCapas.ListIndex)) = 9 Then
        cCapaSel = 9
    Else
        cCapas.ListIndex = 0
        cCapaSel = 1
    End If
    
End Sub

Private Sub cCapas_KeyPress(KeyAscii As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
If IsNumeric(Chr(KeyAscii)) = False Then KeyAscii = 0
End Sub
Private Sub cCopiarInterior()
Dim X As Long
Dim y As Long

Dim o As Long
Dim H As Long

Dim CX As Integer
Dim CY As Integer
o = SeleccionFX - SeleccionIX
H = SeleccionFY - SeleccionIY
SIx = o + 1
SIy = H + 1
ReDim SelInterior(1 To o + 1, 1 To H + 1)

CX = 0
For X = SeleccionIX To SeleccionFX
CX = CX + 1
CY = 0
    For y = SeleccionIY To SeleccionFY
      CY = CY + 1
        SelInterior(CX, CY) = MapData(X, y).InteriorVal
    Next y
Next X
End Sub
Private Sub cCopiarLuces_Click()
Dim y As Long
Dim X As Long
Dim o As Long
Dim H As Long
If Seleccionando Then

o = SeleccionFX - SeleccionIX
H = SeleccionFY - SeleccionIY
Erase SelLuz.TLP
ReDim SelLuz.TLP(1 To o + 1, 1 To H + 1)
SelLuz.Xx = o
SelLuz.xY = H

    For X = SeleccionIX To SeleccionFX
        For y = SeleccionIY To SeleccionFY
            SelLuz.TLP(X - SeleccionIX + 1, y - SeleccionIY + 1).light_value(0) = MapData(X, y).light_value(0)
            SelLuz.TLP(X - SeleccionIX + 1, y - SeleccionIY + 1).light_value(1) = MapData(X, y).light_value(1)
            SelLuz.TLP(X - SeleccionIX + 1, y - SeleccionIY + 1).light_value(2) = MapData(X, y).light_value(2)
            SelLuz.TLP(X - SeleccionIX + 1, y - SeleccionIY + 1).light_value(3) = MapData(X, y).light_value(3)
            
            SelLuz.TLP(X - SeleccionIX + 1, y - SeleccionIY + 1).Luz = MapData(X, y).Luz
            
            SelLuz.TLP(X - SeleccionIX + 1, y - SeleccionIY + 1).LV(0) = MapData(X, y).LV(0)
            SelLuz.TLP(X - SeleccionIX + 1, y - SeleccionIY + 1).LV(1) = MapData(X, y).LV(1)
            SelLuz.TLP(X - SeleccionIX + 1, y - SeleccionIY + 1).LV(2) = MapData(X, y).LV(2)
            SelLuz.TLP(X - SeleccionIX + 1, y - SeleccionIY + 1).LV(3) = MapData(X, y).LV(3)
            
            
                
        Next y
    Next X
    
End If
End Sub

Private Sub cCROSSUL_Click()
If cCROSSUL.value Then
    'Activamos
    If cHorizontal.value Then
        cHorizontal.value = 0
    ElseIf cVertical.value Then
        cVertical.value = 0
    ElseIf cUL.value Then
        cUL.value = 0
    ElseIf cBL.value Then
        cBL.value = 0
    ElseIf cBR.value Then
         cBR.value = 0
    ElseIf cALLC.value Then
         cALLC.value = 0
    ElseIf cCROSSUR.value Then
        cCROSSUR.value = 0
    ElseIf cUR.value Then
        cUR.value = 0
    ElseIf cNotUL.value Then
        cNotUL.value = 0
    ElseIf cNotUR.value Then
        cNotUR.value = 0
    ElseIf cNotBL.value Then
        cNotBL.value = 0
    ElseIf cNotBR.value Then
        cNotBR.value = 0
    End If
Else
End If
End Sub

Private Sub cCROSSUR_Click()
If cCROSSUR.value Then
    'Activamos
    If cHorizontal.value Then
        cHorizontal.value = 0
    ElseIf cVertical.value Then
        cVertical.value = 0
    ElseIf cUL.value Then
        cUL.value = 0
    ElseIf cBL.value Then
        cBL.value = 0
    ElseIf cBR.value Then
         cBR.value = 0
    ElseIf cALLC.value Then
         cALLC.value = 0
    ElseIf cUR.value Then
        cUR.value = 0
    ElseIf cCROSSUL.value Then
        cCROSSUL.value = 0
    ElseIf cNotUL.value Then
        cNotUL.value = 0
    ElseIf cNotUR.value Then
        cNotUR.value = 0
    ElseIf cNotBL.value Then
        cNotBL.value = 0
    ElseIf cNotBR.value Then
        cNotBR.value = 0
    End If
Else
End If
End Sub

Private Sub cFiltro_GotFocus(index As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
HotKeysAllow = False
End Sub

Private Sub cFiltro_KeyPress(index As Integer, KeyAscii As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
If KeyAscii = 13 Then
    Call Filtrar(index)
End If
End Sub

Private Sub cFiltro_LostFocus(index As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
HotKeysAllow = True
End Sub

Private Sub cGrh_KeyPress(KeyAscii As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

On Error GoTo Fallo
If KeyAscii = 13 Then
    Call fPreviewGrh(cGrh.Text)
    If frmMain.cGrh.ListCount > 5 Then
        frmMain.cGrh.RemoveItem 0
    End If
    frmMain.cGrh.AddItem frmMain.cGrh.Text
End If
Exit Sub
Fallo:
    cGrh.Text = 1

End Sub

Private Sub cGrill_Click()
            If frmMain.PreviewGrh.Visible = True Then
                Call modPaneles.VistaPreviaDeSup
            End If
End Sub

Private Sub chkGuardarInf_Click()
chkGuardarInf.Checked = Not chkGuardarInf.Checked
End Sub

Private Sub cHorizontal_Click()
If cHorizontal.value Then
    'Activamos
    If cVertical.value Then
        cVertical.value = 0
    ElseIf cUL.value Then
        cUL.value = 0
    ElseIf cUR.value Then
        cUR.value = 0
    ElseIf cBL.value Then
        cBL.value = 0
    ElseIf cBR.value Then
         cBR.value = 0
    ElseIf cALLC.value Then
         cALLC.value = 0
    ElseIf cCROSSUR.value Then
        cCROSSUR.value = 0
    ElseIf cCROSSUL.value Then
        cCROSSUL.value = 0
    ElseIf cNotUL.value Then
        cNotUL.value = 0
    ElseIf cNotUR.value Then
        cNotUR.value = 0
    ElseIf cNotBL.value Then
        cNotBL.value = 0
    ElseIf cNotBR.value Then
        cNotBR.value = 0
    End If
Else
End If
End Sub

Private Sub cInsertarBorde_Click()
    If cInsertarBorde.value Then
        cQuitarLuz.Enabled = False
        cInsertarLuz.Enabled = False
    Else
        cQuitarLuz.Enabled = True
        cInsertarLuz.Enabled = True
    End If
End Sub

Private Sub cInsertarFunc_Click(index As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
If cInsertarFunc(index).value = True Then
    cQuitarFunc(index).Enabled = False
    If index <> 2 Then cCantFunc(index).Enabled = False
    Call modPaneles.EstSelectPanel((index) + 3, True)
Else
    cQuitarFunc(index).Enabled = True
    If index <> 2 Then cCantFunc(index).Enabled = True
    Call modPaneles.EstSelectPanel((index) + 3, False)
End If
End Sub

Private Sub cInsertarLuz_Click()
    If cInsertarLuz.value Then
        cQuitarLuz.Enabled = False
        cInsertarBorde.Enabled = False
        
    Else
        cQuitarLuz.Enabled = True
        cInsertarBorde.Enabled = True
    End If
End Sub

Private Sub cInsertarParticula_Click()
    If cInsertarParticula.value Then
        cQuitarParticula.Enabled = False
        CmdInteriorI.Enabled = False
        CmdInteriorQ.Enabled = False
    Else
        cQuitarParticula.Enabled = True
        CmdInteriorI.Enabled = True
        CmdInteriorQ.Enabled = True
    End If
End Sub


Private Sub cInsertarTrans_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 22/05/06
'*************************************************
If cInsertarTrans.value = True Then
    cQuitarTrans.Enabled = False
    Call modPaneles.EstSelectPanel(1, True)
Else
    cQuitarTrans.Enabled = True
    Call modPaneles.EstSelectPanel(1, False)
End If
End Sub



Private Sub cInsertarTrigger_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
If cInsertarTrigger.value = True Then
    cQuitarTrigger.Enabled = False
    Call modPaneles.EstSelectPanel(6, True)
Else
    cQuitarTrigger.Enabled = True
    Call modPaneles.EstSelectPanel(6, False)
End If
End Sub

Private Sub cInterio_Click()
cCopiarInterior
End Sub

'Private Sub cmdFixTriggers_Click()
'Dim X As Long
'Dim Y As Long
'Dim newtrigger As Integer

'For X = 1 To 100

'For Y = 1 To 100
    
'    If MapData(X, Y).Trigger > 0 Then
'        newtrigger = newtrigger Xor (2 ^ (MapData(X, Y).Trigger - 1))
'    Else
'        newtrigger = 0
'    End If
'    MapData(X, Y).Trigger = newtrigger
'    newtrigger = 0
'Next Y


'Next X

'MsgBox "LISTO"
'End Sub



Private Sub CmdInteriorI_Click()
    If CmdInteriorI.value Then
        cInsertarParticula.Enabled = False
        CmdInteriorQ.Enabled = False
        cQuitarParticula.Enabled = False
    Else
        cInsertarParticula.Enabled = True
        CmdInteriorQ.Enabled = True
        cQuitarParticula.Enabled = True
    End If
End Sub

Private Sub CmdInteriorQ_Click()
    If CmdInteriorQ.value Then
        cInsertarParticula.Enabled = False
        CmdInteriorI.Enabled = False
        cQuitarParticula.Enabled = False
    Else
        cInsertarParticula.Enabled = True
        CmdInteriorI.Enabled = True
        cQuitarParticula.Enabled = True
    End If
End Sub



Private Sub cNotBL_Click()
If cNotBL.value Then
    'Activamos
    If cHorizontal.value Then
        cHorizontal.value = 0
    ElseIf cVertical.value Then
        cVertical.value = 0
    ElseIf cUL.value Then
        cUL.value = 0
    ElseIf cBL.value Then
        cBL.value = 0
    ElseIf cBR.value Then
         cBR.value = 0
    ElseIf cALLC.value Then
         cALLC.value = 0
    ElseIf cCROSSUR.value Then
        cCROSSUR.value = 0
    ElseIf cUR.value Then
        cUR.value = 0
    ElseIf cCROSSUL.value Then
        cCROSSUL.value = 0
    ElseIf cNotUR.value Then
        cNotUR.value = 0
    ElseIf cNotUL.value Then
        cNotUL.value = 0
    ElseIf cNotBR.value Then
        cNotBR.value = 0
    End If
Else
End If
End Sub

Private Sub cNotBR_Click()
If cNotBR.value Then
    'Activamos
    If cHorizontal.value Then
        cHorizontal.value = 0
    ElseIf cVertical.value Then
        cVertical.value = 0
    ElseIf cUL.value Then
        cUL.value = 0
    ElseIf cBL.value Then
        cBL.value = 0
    ElseIf cBR.value Then
         cBR.value = 0
    ElseIf cALLC.value Then
         cALLC.value = 0
    ElseIf cCROSSUR.value Then
        cCROSSUR.value = 0
    ElseIf cUR.value Then
        cUR.value = 0
    ElseIf cNotUL.value Then
        cNotUL.value = 0
    ElseIf cNotUR.value Then
        cNotUR.value = 0
    ElseIf cNotBL.value Then
        cNotBL.value = 0
    ElseIf cCROSSUL.value Then
        cCROSSUL.value = 0
    End If
Else
End If
End Sub

Private Sub cNotUL_Click()
If cNotUL.value Then
    'Activamos
    If cHorizontal.value Then
        cHorizontal.value = 0
    ElseIf cVertical.value Then
        cVertical.value = 0
    ElseIf cUL.value Then
        cUL.value = 0
    ElseIf cBL.value Then
        cBL.value = 0
    ElseIf cBR.value Then
         cBR.value = 0
    ElseIf cALLC.value Then
         cALLC.value = 0
    ElseIf cCROSSUR.value Then
        cCROSSUR.value = 0
    ElseIf cUR.value Then
        cUR.value = 0
    ElseIf cCROSSUL.value Then
        cCROSSUL.value = 0
    ElseIf cNotUR.value Then
        cNotUR.value = 0
    ElseIf cNotBL.value Then
        cNotBL.value = 0
    ElseIf cNotBR.value Then
        cNotBR.value = 0
    End If
Else
End If
End Sub

Private Sub cNotUR_Click()
If cNotUR.value Then
    'Activamos
    If cHorizontal.value Then
        cHorizontal.value = 0
    ElseIf cVertical.value Then
        cVertical.value = 0
    ElseIf cUL.value Then
        cUL.value = 0
    ElseIf cBL.value Then
        cBL.value = 0
    ElseIf cBR.value Then
         cBR.value = 0
    ElseIf cALLC.value Then
         cALLC.value = 0
    ElseIf cCROSSUR.value Then
        cCROSSUR.value = 0
    ElseIf cUR.value Then
        cUR.value = 0
    ElseIf cNotUL.value Then
        cNotUL.value = 0
    ElseIf cCROSSUL.value Then
        cCROSSUL.value = 0
    ElseIf cNotBL.value Then
        cNotBL.value = 0
    ElseIf cNotBR.value Then
        cNotBR.value = 0
    End If
Else
End If
End Sub



Private Sub COLOREXTRA_Click()
If COLOREXTRA.ListIndex = COLOREXTRA.ListCount - 1 Then
    frmMain.COLOR_CUSTOM_EXTRA.Visible = True
    Label12.Visible = True
Else
    frmMain.COLOR_CUSTOM_EXTRA.Visible = False
    Label12.Visible = False
End If
End Sub

Private Sub COLORSPOT_Click()
If COLORSPOT.ListIndex = COLORSPOT.ListCount - 1 Then
    frmMain.COLOR_CUSTOM_SPOT.Visible = True
    Label11.Visible = True
Else
    frmMain.COLOR_CUSTOM_SPOT.Visible = False
    Label11.Visible = False
End If

End Sub

Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    HoraLuz = Val(Combo1.Text)
    If HoraLuz < 0 Then HoraLuz = 0
    If HoraLuz > 24 Then HoraLuz = 24
    Combo1.Text = HoraLuz
    Recalcular_LUZ HoraLuz
End If
End Sub


Private Sub PegarInterior()

Dim XtilesSel As Integer
Dim YtilesSel As Integer
Dim i As Integer
Dim j As Integer

XtilesSel = UBound(SelInterior, 1)
YtilesSel = UBound(SelInterior, 2)

'si me excedo del limite del mapa, escribo solo lo necesario
If Mx + XtilesSel > 100 Then
    XtilesSel = 100 - Mx + 1
End If
If My + YtilesSel > 100 Then
    YtilesSel = 100 - My + 1
End If

modEdicion.Deshacer_Add "pegar interiores"

For i = 0 To XtilesSel - 1
    For j = 0 To YtilesSel - 1
        MapData(Mx + i, My + j).InteriorVal = SelInterior(i + 1, j + 1)
    Next j
Next i


'Dim X As Long
'Dim Y As Long
'Dim CX As Long
'Dim CY As Long
'CX = 0
'For X = SeleccionIX To SeleccionFX
'CX = CX + 1
'If CX <= SIx Then
'    CY = 0
'    For Y = SeleccionIY To SeleccionFY
'    CY = CY + 1
'    If CY <= SIy Then
'    MapData(X, Y).InteriorVal = SelInterior(CX, CY)
'    End If
'    Next Y
'End If
'Next X


End Sub
Private Sub cPegarLuces_Click()

Dim y As Long
Dim X As Long
Dim tX As Integer
Dim tY As Integer

tX = SelLuz.nX
tY = SelLuz.nY

If SelLuz.Xx = 0 Then Exit Sub

For X = tX To tX + SelLuz.Xx
For y = tY To tY + SelLuz.xY
    If X <= 100 And y <= 100 Then
    
        MapData(X, y).light_value(0) = SelLuz.TLP(X - tX + 1, y - tY + 1).light_value(0)
        MapData(X, y).light_value(1) = SelLuz.TLP(X - tX + 1, y - tY + 1).light_value(1)
        MapData(X, y).light_value(2) = SelLuz.TLP(X - tX + 1, y - tY + 1).light_value(2)
        MapData(X, y).light_value(3) = SelLuz.TLP(X - tX + 1, y - tY + 1).light_value(3)
        
        MapData(X, y).LV(0) = SelLuz.TLP(X - tX + 1, y - tY + 1).LV(0)
        MapData(X, y).LV(1) = SelLuz.TLP(X - tX + 1, y - tY + 1).LV(1)
        MapData(X, y).LV(2) = SelLuz.TLP(X - tX + 1, y - tY + 1).LV(2)
        MapData(X, y).LV(3) = SelLuz.TLP(X - tX + 1, y - tY + 1).LV(3)
        
         MapData(X, y).Luz = SelLuz.TLP(X - tX + 1, y - tY + 1).Luz
    End If
Next y
Next X
End Sub

Private Sub cQuitarLuz_Click()
    If cQuitarLuz.value Then
        cInsertarLuz.Enabled = False
        cInsertarBorde.Enabled = False
    Else
        cInsertarLuz.Enabled = True
        cInsertarBorde.Enabled = True
    End If
End Sub

Private Sub cQuitarParticula_Click()
    If cQuitarParticula.value Then
        cInsertarParticula.Enabled = False
        CmdInteriorI.Enabled = False
        CmdInteriorQ.Enabled = False
    Else
        cInsertarParticula.Enabled = True
        CmdInteriorI.Enabled = True
        CmdInteriorQ.Enabled = True
    End If
End Sub

Private Sub cUL_Click()
If cUL.value Then
    'Activamos
    If cHorizontal.value Then
        cHorizontal.value = 0
    ElseIf cVertical.value Then
        cVertical.value = 0
    ElseIf cUR.value Then
        cUR.value = 0
    ElseIf cBL.value Then
        cBL.value = 0
    ElseIf cBR.value Then
         cBR.value = 0
    ElseIf cALLC.value Then
         cALLC.value = 0
        ElseIf cCROSSUR.value Then
        cCROSSUR.value = 0
    ElseIf cCROSSUL.value Then
        cCROSSUL.value = 0
    ElseIf cNotUL.value Then
        cNotUL.value = 0
    ElseIf cNotUR.value Then
        cNotUR.value = 0
    ElseIf cNotBL.value Then
        cNotBL.value = 0
    ElseIf cNotBR.value Then
        cNotBR.value = 0
    End If
Else
End If
End Sub

Private Sub cUnionManual_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
cInsertarTrans.value = (cUnionManual.value = True)
Call cInsertarTrans_Click
End Sub

Private Sub cUR_Click()
If cUR.value Then
    'Activamos
    If cHorizontal.value Then
        cHorizontal.value = 0
    ElseIf cVertical.value Then
        cVertical.value = 0
    ElseIf cUL.value Then
        cUL.value = 0
    ElseIf cBL.value Then
        cBL.value = 0
    ElseIf cBR.value Then
         cBR.value = 0
    ElseIf cALLC.value Then
         cALLC.value = 0
    ElseIf cCROSSUR.value Then
        cCROSSUR.value = 0
    ElseIf cCROSSUL.value Then
        cCROSSUL.value = 0
    ElseIf cNotUL.value Then
        cNotUL.value = 0
    ElseIf cNotUR.value Then
        cNotUR.value = 0
    ElseIf cNotBL.value Then
        cNotBL.value = 0
    ElseIf cNotBR.value Then
        cNotBR.value = 0
    End If
Else
End If
End Sub

Private Sub cverBloqueos_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
mnuVerBloqueos.Checked = cVerBloqueos.value
End Sub

Private Sub cVertical_Click()
If cVertical.value Then
    'Activamos
    If cHorizontal.value Then
        cHorizontal.value = 0
    ElseIf cUL.value Then
        cUL.value = 0
    ElseIf cUR.value Then
        cUR.value = 0
    ElseIf cBL.value Then
        cBL.value = 0
    ElseIf cBR.value Then
         cBR.value = 0
    ElseIf cALLC.value Then
         cALLC.value = 0
        ElseIf cCROSSUR.value Then
        cCROSSUR.value = 0
    ElseIf cCROSSUL.value Then
        cCROSSUL.value = 0
    ElseIf cNotUL.value Then
        cNotUL.value = 0
    ElseIf cNotUR.value Then
        cNotUR.value = 0
    ElseIf cNotBL.value Then
        cNotBL.value = 0
    ElseIf cNotBR.value Then
        cNotBR.value = 0
    End If

Else
End If

End Sub

Private Sub cverTriggers_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
mnuVerTriggers.Checked = cVerTriggers.value
End Sub

Private Sub cNumFunc_KeyPress(index As Integer, KeyAscii As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
On Error Resume Next

If KeyAscii = 13 Then
    Dim Cont As String
    Cont = frmMain.cNumFunc(index).Text
    Call cNumFunc_LostFocus(index)
    If Cont <> frmMain.cNumFunc(index).Text Then Exit Sub
    If frmMain.cNumFunc(index).ListCount > 5 Then
        frmMain.cNumFunc(index).RemoveItem 0
    End If
    frmMain.cNumFunc(index).AddItem frmMain.cNumFunc(index).Text
    Exit Sub
ElseIf KeyAscii = 8 Then
    
ElseIf IsNumeric(Chr(KeyAscii)) = False Then
    KeyAscii = 0
    Exit Sub
End If

End Sub

Private Sub cNumFunc_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
On Error Resume Next
If cNumFunc(index).Text = vbNullString Then
    frmMain.cNumFunc(index).Text = IIf(index = 1, 500, 1)
End If
End Sub

Private Sub cNumFunc_LostFocus(index As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
On Error Resume Next
    If index = 0 Then
        If frmMain.cNumFunc(index).Text > 499 Or frmMain.cNumFunc(index).Text < 1 Then
            frmMain.cNumFunc(index).Text = 1
        End If
    ElseIf index = 1 Then
        If frmMain.cNumFunc(index).Text < 500 Or frmMain.cNumFunc(index).Text > 32000 Then
            frmMain.cNumFunc(index).Text = 500
        End If
    ElseIf index = 2 Then
        If frmMain.cNumFunc(index).Text < 1 Or frmMain.cNumFunc(index).Text > 32000 Then
            frmMain.cNumFunc(index).Text = 1
        End If
    End If
End Sub

Private Sub cInsertarBloqueo_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 29/05/06
'*************************************************
cInsertarBloqueo.Tag = vbNullString
If cInsertarBloqueo.value = True Then
    cQuitarBloqueo.Enabled = False
    Call modPaneles.EstSelectPanel(2, True)
Else
    cQuitarBloqueo.Enabled = True
    Call modPaneles.EstSelectPanel(2, False)
End If
End Sub

Private Sub cQuitarBloqueo_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
cInsertarBloqueo.Tag = vbNullString
If cQuitarBloqueo.value = True Then
    cInsertarBloqueo.Enabled = False
    Call modPaneles.EstSelectPanel(2, True)
Else
    cInsertarBloqueo.Enabled = True
    Call modPaneles.EstSelectPanel(2, False)
End If
End Sub

Private Sub cQuitarEnEstaCapa_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
If cQuitarEnEstaCapa.value = True Then
    lListado(0).Enabled = False
    cFiltro(0).Enabled = False
    cGrh.Enabled = False
    cSeleccionarSuperficie.Enabled = False
    cQuitarEnTodasLasCapas.Enabled = False
    Call modPaneles.EstSelectPanel(0, True)
Else
    lListado(0).Enabled = True
    cFiltro(0).Enabled = True
    cGrh.Enabled = True
    cSeleccionarSuperficie.Enabled = True
    cQuitarEnTodasLasCapas.Enabled = True
    Call modPaneles.EstSelectPanel(0, False)
End If
End Sub

Private Sub cQuitarEnTodasLasCapas_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
If cQuitarEnTodasLasCapas.value = True Then
    cCapas.Enabled = False
    lListado(0).Enabled = False
    cFiltro(0).Enabled = False
    cGrh.Enabled = False
    cSeleccionarSuperficie.Enabled = False
    cQuitarEnEstaCapa.Enabled = False
    Call modPaneles.EstSelectPanel(0, True)
Else
    cCapas.Enabled = True
    lListado(0).Enabled = True
    cFiltro(0).Enabled = True
    cGrh.Enabled = True
    cSeleccionarSuperficie.Enabled = True
    cQuitarEnEstaCapa.Enabled = True
    Call modPaneles.EstSelectPanel(0, False)
End If
End Sub


Private Sub cQuitarFunc_Click(index As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
If cQuitarFunc(index).value = True Then
    cInsertarFunc(index).Enabled = False
    cCantFunc(index).Enabled = False
    cNumFunc(index).Enabled = False
    cFiltro((index) + 1).Enabled = False
    lListado((index) + 1).Enabled = False
    Call modPaneles.EstSelectPanel((index) + 3, True)
Else
    cInsertarFunc(index).Enabled = True
    cCantFunc(index).Enabled = True
    cNumFunc(index).Enabled = True
    cFiltro((index) + 1).Enabled = True
    lListado((index) + 1).Enabled = True
    Call modPaneles.EstSelectPanel((index) + 3, False)
End If
End Sub

Private Sub cQuitarTrans_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
If cQuitarTrans.value = True Then
    cInsertarTransOBJ.Enabled = False
    cInsertarTrans.Enabled = False
    cUnionManual.Enabled = False
    cUnionAuto.Enabled = False
    tTMapa.Enabled = False
    tTX.Enabled = False
    tTY.Enabled = False
    mnuInsertarTransladosAdyasentes.Enabled = False
    Call modPaneles.EstSelectPanel(1, True)
Else
    tTMapa.Enabled = True
    tTX.Enabled = True
    tTY.Enabled = True
    cUnionAuto.Enabled = True
    cUnionManual.Enabled = True
    cInsertarTrans.Enabled = True
    cInsertarTransOBJ.Enabled = True
    mnuInsertarTransladosAdyasentes.Enabled = True
    Call modPaneles.EstSelectPanel(1, False)
End If
End Sub

Private Sub cQuitarTrigger_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
If cQuitarTrigger.value = True Then
    lListado(4).Enabled = False
    cInsertarTrigger.Enabled = False
    Call modPaneles.EstSelectPanel(6, True)
Else
    lListado(4).Enabled = True
    cInsertarTrigger.Enabled = True
    Call modPaneles.EstSelectPanel(6, False)
End If
End Sub

Private Sub cSeleccionarSuperficie_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
If cSeleccionarSuperficie.value = True Then
    cQuitarEnTodasLasCapas.Enabled = False
    cQuitarEnEstaCapa.Enabled = False
    Call modPaneles.EstSelectPanel(0, True)
    If frmMain.bI.value Then
        If frmMain.lListado(5).ListIndex >= 0 Then
            If Val(ReadField(1, frmMain.lListado(5).List(frmMain.lListado(5).ListIndex), Asc("-"))) > 0 Then
                SobreIndex = Val(ReadField(1, frmMain.lListado(5).List(frmMain.lListado(5).ListIndex), Asc("-")))
            Else
                SobreIndex = 0
            End If
        Else
            SobreIndex = 0
        End If
    Else
        If frmMain.lListado(0).ListIndex >= 0 And SelTexWe > 0 Then
            SobreIndex = DameIndexEnTexUL(SelTexWe)
        Else
            SobreIndex = 0
        End If
    End If
    
    
Else
    cQuitarEnTodasLasCapas.Enabled = True
    cQuitarEnEstaCapa.Enabled = True
    Call modPaneles.EstSelectPanel(0, False)
End If
End Sub

Private Sub cUnionAuto_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
frmUnionAdyacente.Show
End Sub

Private Sub decorb_Click()
If lListado(6).Visible Or lListado(3).Visible Then
    If decorb.value Then
        lListado(3).Visible = False
        lListado(6).Visible = True
        cInsertarFunc(2).Caption = "Insertar Decor"
        cQuitarFunc(2).Caption = "Quitar Decor"
        cNumFunc(2).Text = general_field_read(2, lListado(6).Text, Asc("#"))
        chkDecorBloq.Visible = True
    Else
        lListado(3).Visible = True
        lListado(6).Visible = False
        cInsertarFunc(2).Caption = "Insertar Obj"
        cQuitarFunc(2).Caption = "Quitar Obj"
        cNumFunc(2).Text = general_field_read(2, lListado(3).Text, Asc("#"))
                chkDecorBloq.Visible = False
    End If
ElseIf lListado(1).Visible Or lListado(7).Visible Then
    'NPC THING.
    If decorb.value Then
        lListado(7).Visible = True
        lListado(1).Visible = False
        cNumFunc(0).Text = general_field_read(2, lListado(7).Text, Asc("#"))
    Else
        lListado(7).Visible = False
        lListado(1).Visible = True
        cNumFunc(0).Text = general_field_read(2, lListado(1).Text, Asc("#"))
    End If
Else
    'TRIGGER.
    If decorb.value Then
        lListado(8).Visible = True
        lListado(4).Visible = False
            frmMain.cVerTriggers.Caption = "Mostrar Tipo-Terreno"
            frmMain.cInsertarTrigger.Caption = "Insertar Tipo-Terreno"
            frmMain.cQuitarTrigger.Caption = "Quitar Tipo-Terreno"
            frmMain.decorb.Caption = "Triggers"
    Else
        lListado(8).Visible = False
        lListado(4).Visible = True
            frmMain.cVerTriggers.Caption = "Mostrar Triggers"
            frmMain.cInsertarTrigger.Caption = "Insertar Triggers"
            frmMain.cQuitarTrigger.Caption = "Quitar Triggers"
            frmMain.decorb.Caption = "Tipo Terreno"
    End If
End If
End Sub

Private Sub Form_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Me.SetFocus

End Sub

Private Sub Form_Load()
    ALH = 29
    Statz = True
    
End Sub

Private Sub grdNuevoMapa_Click()
modMapIO.GuardarMapa
End Sub




Private Sub lLuces_Click()
If lLuces.ListIndex <> LuzSelecta Then
    LuzSelecta = lLuces.ListIndex
    tLuz.Text = lLuces.ListIndex
End If

End Sub



Private Sub lvButtons_H1_Click()

    
End Sub

Private Sub MainViewPic_DblClick()
'*************************************************
'Author: ^[GS]^
'Last modified: 28/05/06
'*************************************************

If Not MapaCargado Then Exit Sub

If SobreX > 0 And SobreY > 0 Then
    DobleClick Val(SobreX), Val(SobreY)
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 24/11/08
'*************************************************
' HotKeys
If HotKeysAllow = False Then Exit Sub

Select Case UCase(Chr(KeyAscii))
    Case "S" ' Activa/Desactiva Insertar Superficie
        cSeleccionarSuperficie.value = (cSeleccionarSuperficie.value = False)
        Call cSeleccionarSuperficie_Click
    Case "T" ' Activa/Desactiva Insertar Translados
        cInsertarTrans.value = (cInsertarTrans.value = False)
        Call cInsertarTrans_Click
    Case "B" ' Activa/Desactiva Insertar Bloqueos
        cInsertarBloqueo.value = (cInsertarBloqueo.value = False)
        Call cInsertarBloqueo_Click
    Case "N" ' Activa/Desactiva Insertar NPCs
        cInsertarFunc(0).value = (cInsertarFunc(0).value = False)
        Call cInsertarFunc_Click(0)
   ' Case "H" ' Activa/Desactiva Insertar NPCs Hostiles
   '     cInsertarFunc(1).value = (cInsertarFunc(1).value = False)
   '     Call cInsertarFunc_Click(1)
    Case "O" ' Activa/Desactiva Insertar Objetos
        cInsertarFunc(2).value = (cInsertarFunc(2).value = False)
        Call cInsertarFunc_Click(2)
    Case "G" ' Activa/Desactiva Insertar Triggers
        cInsertarTrigger.value = (cInsertarTrigger.value = False)
        Call cInsertarTrigger_Click
    Case "Q" ' Quitar Funciones
        Call mnuQuitarFunciones_Click
End Select
End Sub




Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    'If Seleccionando Then CopiarSeleccion

    
End Sub

Private Sub lListado_Click(index As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 29/05/06
'*************************************************
On Error Resume Next
If HotKeysAllow = False Then
    lListado(index).Tag = lListado(index).ListIndex
    Select Case index
        Case 0
            
            SelTexWe = Val(ReadField(2, lListado(index).List(lListado(index).ListIndex), Asc("[")))
            PutX = 0
            PutY = 0
            SelTexRecort = False
            SelTexFrame = 0
            AnalizeTexture SelTexWe
            If SelTexWe > 0 Then SobreIndex = DameIndexEnTexUL(SelTexWe)
            If frmMain.PreviewGrh.Visible = True Then
                Call modPaneles.VistaPreviaDeSup
            End If
        Case 1
            cNumFunc(0).Text = general_field_read(2, lListado(index).Text, Asc("#"))
        Case 2
            cNumFunc(1).Text = general_field_read(2, lListado(index).Text, Asc("#"))
        Case 3
            cNumFunc(2).Text = general_field_read(2, lListado(index).Text, Asc("#"))
            VistaPreviaIndex ObjData(Val(cNumFunc(2).Text)).grh_index
        Case 5
            VistaPreviaIndex
        Case 6
            cNumFunc(2).Text = general_field_read(2, lListado(index).Text, Asc("#"))
            VistaPreviaIndex Val(DecorData(Val(cNumFunc(2).Text)).DecorGrh(1))
        Case 7
            cNumFunc(0).Text = general_field_read(2, lListado(index).Text, Asc("#"))
    End Select
Else
    lListado(index).ListIndex = lListado(index).Tag
End If

End Sub
Public Sub VistaPreviaIndex(Optional ByVal index As Integer)

If Statz Then Exit Sub
If index = 0 Then
zCurrentIndex = Val(ReadField(1, lListado(5).List(lListado(5).ListIndex), Asc("-")))
Else
zCurrentIndex = index
End If

PreviewGrh.Cls
If NewIndexData(zCurrentIndex).OverWriteGrafico > 0 Then

    Dim s As RECT
    
    s.left = 0
    s.Right = 0
    s.Bottom = EstaticData(NewIndexData(zCurrentIndex).Estatic).H
    s.Right = EstaticData(NewIndexData(zCurrentIndex).Estatic).W
    
    modDXEngine.DibujareEnHwnd2 frmMain.PreviewGrh.hWnd, zCurrentIndex, s, 0, 0, True, True, 256, 256


End If



End Sub
Private Sub lListado_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
'*************************************************
'Author: ^[GS]^
'Last modified: 29/05/06
'*************************************************
If index = 3 And Button = 2 Then
    If lListado(3).ListIndex > -1 Then Me.PopupMenu mnuObjSc
End If
End Sub

Private Sub lListado_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
'*************************************************
'Author: ^[GS]^
'Last modified: 22/05/06
'*************************************************
On Error Resume Next
HotKeysAllow = False
End Sub







Private Sub mele_decor_Click()

If TipoSeleccionado = 0 Or ObjetoSeleccionado.X = 0 Or ObjetoSeleccionado.y = 0 Then
    MsgBox "No tienes ningun elemento seleccionado."
    Exit Sub
End If
If TipoSeleccionado = 1 Then
    frmEditorDecor.Show
    frmEditorDecor.Parse
    TipoSeleccionado = 0
ElseIf TipoSeleccionado = 2 Then
    FrmEditorNpc.Show
    FrmEditorNpc.Parse
    TipoSeleccionado = 0
End If
End Sub

'Private Sub mnuAbrirMapaI_Click()[MAPA INTERFACE VIEJO]
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
'Dialog.CancelError = True
'On Error GoTo ErrHandler

'DeseaGuardarMapa Dialog.FileName

'ObtenerNombreArchivo False

'If Len(Dialog.FileName) < 3 Then Exit Sub

'    If WalkMode = True Then
'        Call modGeneral.ToggleWalkMode
'    End If
'
'    Call modMapIO.NuevoMapa
'    modMapIO.MapaI_Cargar Dialog.FileName, MapData, False
'    DoEvents
'    mnuReAbrirMapa.Enabled = True
'    EngineRun = True
'    TIPOMAPAX = 0
'
'Exit Sub
'ErrHandler:
'End Sub



Private Sub mnuActualizarGraficos_Click()

DXPool.Texture_Remove_All
End Sub

Private Sub mnuActualizarSuperficies_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
'Call modIndices.CargarIndicesSuperficie
frmMain.lListado(0).Clear
modExtras.LoadTexWe
End Sub

Private Sub mnuAbrirMapa_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
Dim B As Byte
'*************************************************
On Error GoTo ErrHandler

Dialog.CancelError = True



DeseaGuardarMapa Dialog.FileName


ObtenerNombreArchivo False


If Len(Dialog.FileName) < 3 Then Exit Sub

    If WalkMode = True Then
        Call modGeneral.ToggleWalkMode
    End If

    Call modMapIO.NuevoMapa
    Dim QueAbro As Boolean
    Dim o As Integer
    Dim HayInf As Boolean
    o = FreeFile
    Open Dialog.FileName For Binary As #o
        If LOF(o) > 200000 Then
            QueAbro = True
        End If
    Close #o
    If UCase$(Right$(Dialog.FileName, 4)) = "TEMP" Then
        MapaTemporal = True
    Else
        MapaTemporal = False
    End If
    
    modMapIO.AbrirMapaComun Dialog.FileName
    
    DoEvents

    EngineRun = True

    TIPOMAPAX = 0
    
Exit Sub
ErrHandler:

MsgBox "ERROR EN ABRIR" & "_" & B & "_" & Err.Description
End Sub

Private Sub mnuacercade_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
frmAbout.Show
End Sub



Private Sub mnuActualizarNPCs_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call modIndices.CargarIndicesNPC
End Sub

Private Sub mnuActualizarObjs_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call modIndices.CargarIndicesOBJ
End Sub

Private Sub mnuActualizarTriggers_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call modIndices.CargarIndicesTriggers
End Sub

Private Sub mnuAutoCapturarTranslados_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 28/05/06
'*************************************************
mnuAutoCapturarTranslados.Checked = (mnuAutoCapturarTranslados.Checked = False)
End Sub

Private Sub mnuAutoCapturarSuperficie_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 28/05/06
'*************************************************
mnuAutoCapturarSuperficie.Checked = (mnuAutoCapturarSuperficie.Checked = False)

End Sub

Private Sub mnuAutoCompletarSuperficies_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
mnuAutoCompletarSuperficies.Checked = (mnuAutoCompletarSuperficies.Checked = False)

End Sub

Private Sub mnuAutoGuardarMapas_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
frmAutoGuardarMapa.Show
End Sub

Private Sub mnuAutoQuitarFunciones_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
mnuAutoQuitarFunciones.Checked = (mnuAutoQuitarFunciones.Checked = False)

End Sub

Private Sub mnuBloquear_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Dim i As Byte
For i = 0 To 6
    If i <> 2 Then
        frmMain.SelectPanel(i).value = False
        Call VerFuncion(i, False)
    End If
Next

modPaneles.VerFuncion 2, True
End Sub

Private Sub mnuBloquearBordes_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call modEdicion.Bloquear_Bordes
End Sub

Private Sub mnuBloquearMapa_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call modEdicion.Bloqueo_Todo(1)
End Sub

Private Sub mnuBloquearS_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 01/11/08
'*************************************************
Call modEdicion.Deshacer_Add("Bloquear Selección")
Call BlockearSeleccion
End Sub

Private Sub mnuConfigAvanzada_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
frmConfigSup.Show
End Sub

Private Sub mnuConfigObjTrans_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 29/05/06
'*************************************************
Cfg_TrOBJ = cNumFunc(2).Text
End Sub

Private Sub mnuCopiar_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 01/11/08
'*************************************************
Call CopiarSeleccionSinTraslados
End Sub

Private Sub mnuCopiarOld_Click()
Call CopiarSeleccion
End Sub

Private Sub mnuCortar_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 01/11/08
'*************************************************
Call modEdicion.Deshacer_Add("Cortar Selección")
Call CortarSeleccion
End Sub

Private Sub mnuDeshacer_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 15/10/06
'*************************************************
Call modEdicion.Deshacer_Recover
End Sub

Private Sub mnuDeshacerPegado_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 01/11/08
'*************************************************
Call modEdicion.Deshacer_Add("Deshacer Pegado de Selección")
Call DePegar
End Sub

Private Sub mnuGRHaBMP_Click()
Dim P As Long
If fTextureMaker.Visible Then Unload fTextureMaker

For P = 1 To NumTexWe
    fTextureMaker.Combo1.AddItem TexWE(P).Name & " [" & P & "]"
Next P
fTextureMaker.Show
End Sub

Private Sub mnuGuardarcomoBMP_Click()
'*************************************************
'Author: Salvito
'Last modified: 01/05/2008 - ^[GS]^
'*************************************************
        Render.Text1.Visible = True
        Render.Text2.Visible = True
        Render.Label1.Visible = True
        Render.Label2.Visible = True
        Render.Command1.Visible = True
        
        Render.Show

End Sub

Private Sub mnuGuardarcomoJPG_Click()
'*************************************************
'Author: Salvito
'Last modified: 01/05/2008 - ^[GS]^
'*************************************************
    Render.Text1.Visible = False
        Render.Text2.Visible = False
        Render.Label1.Visible = False
        Render.Label2.Visible = False
        Render.Command1.Visible = False
        
    Render.Show
    Render.Render
End Sub

Private Sub mnuGuardarMapa_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
modMapIO.GuardarMapa Dialog.FileName
End Sub

'Private Sub mnuGuardarMapaComo_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
'TIPOMAPA = 1
'modMapIO.GuardarMapa , 1
'End Sub

Private Sub mnuGuardarUltimaConfig_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 23/05/06
'*************************************************
mnuGuardarUltimaConfig.Checked = (mnuGuardarUltimaConfig.Checked = False)
End Sub

Private Sub mnuInfoMap_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
frmMapInfo.Show
frmMapInfo.Visible = True
End Sub

Private Sub mnuInformes_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
frmInformes.Show
End Sub



Private Sub mnuInsertarSuperficieAlAzar_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call modEdicion.Superficie_Azar
End Sub

Private Sub mnuInsertarSuperficieEnBordes_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call modEdicion.Superficie_Bordes
End Sub

Private Sub mnuInsertarSuperficieEnTodo_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call modEdicion.Superficie_Todo
End Sub

Private Sub mnuInsertarTransladosAdyasentes_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
frmUnionAdyacente.Show
End Sub

Private Sub mnuManual_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 24/11/08
'*************************************************
If LenB(Dir(App.PATH & "\manual\index.html", vbArchive)) <> 0 Then
    Call Shell("explorer " & App.PATH & "\manual\index.html")
    DoEvents
End If
End Sub

Private Sub mnuLine2_Click()
FrmExtras.Show

End Sub

Private Sub mnuModoCaminata_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 28/05/06
'*************************************************
ToggleWalkMode
End Sub

Private Sub mnuNPCs_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Dim i As Byte
For i = 0 To 6
    If i <> 3 Then
        frmMain.SelectPanel(i).value = False
        Call VerFuncion(i, False)
    End If
Next
modPaneles.VerFuncion 3, True
End Sub



'Private Sub mnuNPCsHostiles_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
'Dim i As Byte
'For i = 0 To 6
'    If i <> 4 Then
'        frmMain.SelectPanel(i).value = False
'        Call VerFuncion(i, False)
'    End If
'Next
'modPaneles.VerFuncion 4, True
'End Sub

Private Sub mnuNuevoMapa_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
On Error Resume Next
Dim LoopC As Integer

If MsgBox("¿Deseas crear un mapa grafico?", vbYesNo, "Tipo de Mapa") = vbYes Then
    TIPOMAPAX = 1
Else
    TIPOMAPAX = 0
End If

DeseaGuardarMapa Dialog.FileName

For LoopC = 0 To frmMain.MapPest.Count
    frmMain.MapPest(LoopC).Visible = False
Next

frmMain.Dialog.FileName = Empty

If WalkMode = True Then
    Call modGeneral.ToggleWalkMode
End If
modDXEngine.SPOTLIGHTS_LIMPIARTODOS

NuevoMapa.Show

Do Until NuevoOk

    DoEvents
    
Loop
NuevoOk = False

Call modMapIO.NuevoMapa






End Sub

Private Sub mnuObjetos_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Dim i As Byte
For i = 0 To 6
    If i <> 5 Then
        frmMain.SelectPanel(i).value = False
        Call VerFuncion(i, False)
    End If
Next
modPaneles.VerFuncion 5, True
End Sub


Private Sub mnuOptimizar_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 22/09/06
'*************************************************
frmOptimizar.Show
End Sub

Private Sub mnuPegar_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 01/11/08
'*************************************************
Call modEdicion.Deshacer_Add("Pegar Selección")
Call PegarSeleccion
End Sub

Private Sub mnuQBloquear_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
modPaneles.VerFuncion 2, False
End Sub

Private Sub mnuQNPCs_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
modPaneles.VerFuncion 3, False
End Sub

'Private Sub mnuQNPCsHostiles_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
'modPaneles.VerFuncion 4, False
'End Sub

Private Sub mnuQObjetos_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
modPaneles.VerFuncion 5, False
End Sub

Private Sub mnuQSuperficie_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
modPaneles.VerFuncion 0, False
End Sub

Private Sub mnuQTranslados_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
modPaneles.VerFuncion 1, False
End Sub

Private Sub mnuQTriggers_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
modPaneles.VerFuncion 6, False
End Sub


Private Sub mnuQuitarBloqueos_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call modEdicion.Bloqueo_Todo(0)
End Sub

Private Sub mnuQuitarFunciones_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
' Superficies
cSeleccionarSuperficie.value = False
Call cSeleccionarSuperficie_Click
cQuitarEnEstaCapa.value = False
Call cQuitarEnEstaCapa_Click
cQuitarEnTodasLasCapas.value = False
Call cQuitarEnTodasLasCapas_Click
' Translados
cQuitarTrans.value = False
Call cQuitarTrans_Click
cInsertarTrans.value = False
Call cInsertarTrans_Click
' Bloqueos
cQuitarBloqueo.value = False
Call cQuitarBloqueo_Click
cInsertarBloqueo.value = False
Call cInsertarBloqueo_Click
' Otras funciones
cInsertarFunc(0).value = False
Call cInsertarFunc_Click(0)
cInsertarFunc(1).value = False
Call cInsertarFunc_Click(1)
cInsertarFunc(2).value = False
Call cInsertarFunc_Click(2)
cQuitarFunc(0).value = False
Call cQuitarFunc_Click(0)
cQuitarFunc(1).value = False
Call cQuitarFunc_Click(1)
cQuitarFunc(2).value = False
Call cQuitarFunc_Click(2)
' Triggers
cInsertarTrigger.value = False
Call cInsertarTrigger_Click
cQuitarTrigger.value = False
Call cQuitarTrigger_Click


'Luces
cInsertarLuz.value = False
Call cInsertarLuz_Click

cQuitarLuz.value = False
Call cQuitarLuz_Click

'Particulas
cQuitarParticula.value = False
Call cQuitarParticula_Click

cInsertarParticula.value = False
Call cInsertarParticula_Click

CmdInteriorI.value = False
CmdInteriorI_Click
CmdInteriorQ.value = False
CmdInteriorQ_Click


End Sub

Private Sub mnuQuitarNPCs_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call modEdicion.Quitar_NPCs(False)
End Sub

'Private Sub mnuQuitarNPCsHostiles_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
'Call modEdicion.Quitar_NPCs(True)
'End Sub

Private Sub mnuQuitarObjetos_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call modEdicion.Quitar_Objetos
End Sub

Private Sub mnuQuitarSuperficieBordes_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call modEdicion.Quitar_Bordes
End Sub

Private Sub mnuQuitarSuperficieDeCapa_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
If cCapaSel > 0 And cCapaSel <= 5 Then
    Call modEdicion.Quitar_Capa(cCapaSel)
ElseIf cCapaSel = 9 Then
    Call modEdicion.Quitar_Capa(2)
End If
End Sub

Private Sub mnuQuitarTODO_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call modEdicion.Borrar_Mapa
End Sub

Private Sub mnuQuitarTranslados_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 16/10/06
'*************************************************
Call modEdicion.Quitar_Translados
End Sub

Private Sub mnuQuitarTriggers_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Call modEdicion.Quitar_Triggers
End Sub

'Private Sub mnuReAbrirMapa_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
'On Error GoTo ErrHandler
'    If General_File_Exist(Dialog.FileName, vbArchive) = False Then Exit Sub
'    If MapInfo.Changed = 1 Then
'        If MsgBox(MSGMod, vbExclamation + vbYesNo) = vbYes Then
'            modMapIO.GuardarMapa Dialog.FileName
'        End If
'    End If
'    Call modMapIO.NuevoMapa
'    modMapIO.AbrirMapa Dialog.FileName, MapData
'    DoEvents
'    mnuReAbrirMapa.Enabled = True
'    EngineRun = True
'Exit Sub
'ErrHandler:
'End Sub

Private Sub mnuRealizarOperacion_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 01/11/08
'*************************************************

UsingUndoSelection = False
Call modEdicion.Deshacer_Add("Realizar Operación en Selección")
UsingUndoSelection = True

Call AccionSeleccion
End Sub

Private Sub mnuSalir_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Unload Me
End Sub

Private Sub mnuSuperficie_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Dim i As Byte
For i = 0 To 6
    If i <> 0 Then
        frmMain.SelectPanel(i).value = False
        Call VerFuncion(i, False)
    End If
Next
modPaneles.VerFuncion 0, True
End Sub

Private Sub mnuTranslados_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Dim i As Byte
For i = 0 To 6
    If i <> 1 Then
        frmMain.SelectPanel(i).value = False
        Call VerFuncion(i, False)
    End If
Next
modPaneles.VerFuncion 1, True
End Sub

Private Sub mnuTriggers_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Dim i As Byte
For i = 0 To 6
    If i <> 6 Then
        frmMain.SelectPanel(i).value = False
        Call VerFuncion(i, False)
    End If
Next
modPaneles.VerFuncion 6, True
End Sub

Private Sub mnuUtilizarDeshacer_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 16/10/06
'*************************************************
mnuUtilizarDeshacer.Checked = (mnuUtilizarDeshacer.Checked = False)
End Sub


Private Sub mnuVerAutomatico_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
mnuVerAutomatico.Checked = (mnuVerAutomatico.Checked = False)
End Sub

Private Sub mnuVerBloqueos_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
cVerBloqueos.value = (cVerBloqueos.value = False)
mnuVerBloqueos.Checked = cVerBloqueos.value

End Sub

Private Sub mnuVerCapa1_Click()
mnuVerCapa1.Checked = (mnuVerCapa1.Checked = False)
End Sub

Private Sub mnuVerCapa2_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
mnuVerCapa2.Checked = (mnuVerCapa2.Checked = False)
End Sub

Private Sub mnuVerCapa3_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
mnuVerCapa3.Checked = (mnuVerCapa3.Checked = False)
End Sub

Private Sub mnuVerCapa4_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
mnuVerCapa4.Checked = (mnuVerCapa4.Checked = False)
End Sub


Private Sub MnuVerCapa5_Click()
    If VerCapa5 Then
    MnuVerCapa5.Checked = False
    VerCapa5 = MnuVerCapa5.Checked
    Else
    MnuVerCapa5.Checked = True
    VerCapa5 = True
    End If
End Sub

Private Sub MnuVerCapa9_Click()
MnuVerCapa9.Checked = (MnuVerCapa9.Checked = False)
End Sub

Private Sub mnuVerGrilla_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 24/11/08
'*************************************************
VerGrilla = (VerGrilla = False)
mnuVerGrilla.Checked = VerGrilla
End Sub

Private Sub mnuVerNPCs_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 26/05/06
'*************************************************
mnuVerNPCs.Checked = (mnuVerNPCs.Checked = False)

End Sub

Private Sub mnuVerObjetos_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 26/05/06
'*************************************************
mnuVerObjetos.Checked = (mnuVerObjetos.Checked = False)

End Sub

Private Sub mnuVerTranslados_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 26/05/06
'*************************************************
mnuVerTranslados.Checked = (mnuVerTranslados.Checked = False)

End Sub

Private Sub mnuVerTriggers_Click()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
cVerTriggers.value = (cVerTriggers.value = False)
mnuVerTriggers.Checked = cVerTriggers.value
End Sub





Private Sub mVerDecors_Click()
    mVerDecors.Checked = Not mVerDecors.Checked
    VerDecors = mVerDecors.Checked
End Sub

Private Sub OpMpGr_Click()
On Error GoTo ErrHandler

Dialog.CancelError = True



DeseaGuardarMapa Dialog.FileName


ObtenerNombreArchivo False


If Len(Dialog.FileName) < 3 Then Exit Sub

    If WalkMode = True Then
        Call modGeneral.ToggleWalkMode
    End If

    Call modMapIO.NuevoMapa
    modMapIO.AbrirMapaGrafico Dialog.FileName

    DoEvents
    EngineRun = True

    TIPOMAPAX = 1
    
Exit Sub
ErrHandler:

MsgBox "ERROR EN ABRIR" & "_" & B & "_" & Err.Description


End Sub

Private Sub picRadar_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
'*************************************************
'Author: ^[GS]^
'Last modified: 29/05/06
'*************************************************
If X < MinXBorder Then X = 11
If X > MaxXBorder Then X = 89
If y < MinYBorder Then y = 10
If y > MaxYBorder Then y = 92

UserPos.X = X
UserPos.y = y
bRefreshRadar = True
End Sub

Private Sub picRadar_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
'*************************************************
'Author: ^[GS]^
'Last modified: 28/05/06
'*************************************************
MiRadarX = X
MiRadarY = y
End Sub



Private Sub MainViewPic_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06 - GS
'Last modified: 20/11/07 - Loopzer
'*************************************************

Dim tX As Integer
Dim tY As Integer

If Not MapaCargado Then Exit Sub

ConvertCPtoTP 0, 0, X, y, tX, tY
Mx = tX
My = tY

If tY < 1 Or tY > 100 Then Exit Sub
If tX < 1 Or tX > 100 Then Exit Sub

'If Shift = 1 And Button = 2 Then PegarSeleccion tX, tY: Exit Sub
If Shift = 1 And Button = 1 Then
    Seleccionando = True
    SeleccionIX = tX '+ UserPos.X
    SeleccionIY = tY '+ UserPos.Y
Else
    ClickEdit Button, tX, tY
End If
SelLuz.nX = tX
SelLuz.nY = tY


End Sub


Private Sub MainViewpic_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06 - GS
'*************************************************

Dim tX As Integer
Dim tY As Integer

'Make sure map is loaded
If Not MapaCargado Then Exit Sub
HotKeysAllow = True

ConvertCPtoTP 0, 0, X, y, tX, tY
Mx = tX
My = tY

Label13.Caption = "X: " & tX & " - Y: " & tY
If tX < 10 Or tY < 10 Or tX > 90 Or tY > 90 Then
    Label13.ForeColor = vbRed
Else
    Label13.ForeColor = vbWhite
End If
 If Shift = 1 And Button = 1 Then
    Seleccionando = True
    SeleccionFX = tX '+ TileX
    SeleccionFY = tY '+ TileY
Else
    ClickEdit Button, tX, tY
End If
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 24/11/08
'*************************************************

' Guardar configuración
WriteVar inipath & "WorldEditor.ini", "CONFIGURACION", "GuardarConfig", IIf(frmMain.mnuGuardarUltimaConfig.Checked = True, "1", "0")
If frmMain.mnuGuardarUltimaConfig.Checked = True Then
    WriteVar inipath & "WorldEditor.ini", "PATH", "UltimoMapa", Dialog.FileName
    WriteVar inipath & "WorldEditor.ini", "MOSTRAR", "ControlAutomatico", IIf(frmMain.mnuVerAutomatico.Checked = True, "1", "0")
    WriteVar inipath & "WorldEditor.ini", "MOSTRAR", "Capa2", IIf(frmMain.mnuVerCapa2.Checked = True, "1", "0")
    WriteVar inipath & "WorldEditor.ini", "MOSTRAR", "Capa3", IIf(frmMain.mnuVerCapa3.Checked = True, "1", "0")
    WriteVar inipath & "WorldEditor.ini", "MOSTRAR", "Capa4", IIf(frmMain.mnuVerCapa4.Checked = True, "1", "0")
    WriteVar inipath & "WorldEditor.ini", "MOSTRAR", "Capa9", IIf(frmMain.mnuVerCapa4.Checked = True, "1", "0")
    WriteVar inipath & "WorldEditor.ini", "MOSTRAR", "Translados", IIf(frmMain.mnuVerTranslados.Checked = True, "1", "0")
    WriteVar inipath & "WorldEditor.ini", "MOSTRAR", "Objetos", IIf(frmMain.mnuVerObjetos.Checked = True, "1", "0")
    WriteVar inipath & "WorldEditor.ini", "MOSTRAR", "NPCs", IIf(frmMain.mnuVerNPCs.Checked = True, "1", "0")
    WriteVar inipath & "WorldEditor.ini", "MOSTRAR", "Triggers", IIf(frmMain.mnuVerTriggers.Checked = True, "1", "0")
    WriteVar inipath & "WorldEditor.ini", "MOSTRAR", "Grilla", IIf(frmMain.mnuVerGrilla.Checked = True, "1", "0")
    WriteVar inipath & "WorldEditor.ini", "MOSTRAR", "Decors", IIf(frmMain.mVerDecors.Checked = True, "1", "0")
    
    WriteVar inipath & "WorldEditor.ini", "MOSTRAR", "Bloqueos", IIf(frmMain.mnuVerBloqueos.Checked = True, "1", "0")
    WriteVar inipath & "WorldEditor.ini", "MOSTRAR", "LastPos", UserPos.X & "-" & UserPos.y
    WriteVar inipath & "WorldEditor.ini", "CONFIGURACION", "UtilizarDeshacer", IIf(frmMain.mnuUtilizarDeshacer.Checked = True, "1", "0")
    WriteVar inipath & "WorldEditor.ini", "CONFIGURACION", "AutoCapturarTrans", IIf(frmMain.mnuAutoCapturarTranslados.Checked = True, "1", "0")
    WriteVar inipath & "WorldEditor.ini", "CONFIGURACION", "AutoCapturarSup", IIf(frmMain.mnuAutoCapturarSuperficie.Checked = True, "1", "0")
    WriteVar inipath & "WorldEditor.ini", "CONFIGURACION", "ObjTranslado", Val(Cfg_TrOBJ)
End If

'Allow MainLoop to close program
If prgRun = True Then
    prgRun = False
    Cancel = 1
End If

End Sub



Private Sub pInter_Click()
PegarInterior
End Sub

Private Sub PONERSPOT_Click()
If PONERSPOT Then
    QUITARSPOT.value = False
End If
End Sub

Private Sub PreviewGrh_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim kx As Integer
Dim ky As Integer
If SelTexWe = 0 Then Exit Sub
kx = X / Screen.TwipsPerPixelX
ky = y / Screen.TwipsPerPixelY
If Button = vbLeftButton Then



Dim P As Long


    If TexWE(SelTexWe).NumIndex > 0 Then
    
        For P = 1 To TexWE(SelTexWe).NumIndex
            With EstaticData(NewIndexData(TexWE(SelTexWe).index(P).Num).Estatic)
                
                If kx >= TexWE(SelTexWe).index(P).X And kx <= TexWE(SelTexWe).index(P).X + .W And ky >= TexWE(SelTexWe).index(P).y And ky <= TexWE(SelTexWe).index(P).y + .H Then
                
                    SelTexFrame = P
                    Exit For
                End If
            End With
        Next P
        If P > TexWE(SelTexWe).NumIndex Then SelTexFrame = 0
    End If

Else

SelTexFrame = 0

End If

VistaPreviaDeSup
End Sub

Private Sub PreviewGrh_Paint()
Call modPaneles.VistaPreviaDeSup
End Sub

Private Sub QUITARSPOT_Click()
If QUITARSPOT Then
    PONERSPOT.value = False
End If
End Sub

Private Sub SelectPanel_Click(index As Integer)
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
Dim i As Byte
For i = 0 To 10
    If i <> index Then
        SelectPanel(i).value = False
        Call VerFuncion(i, False)
    End If
Next
If mnuAutoQuitarFunciones.Checked = True Then Call mnuQuitarFunciones_Click
Call VerFuncion(index, SelectPanel(index).value)
End Sub




Private Sub SPOTEDITAR_Click()


If frmMain.PONERSPOT.value Then frmMain.PONERSPOT.value = False
If frmMain.QUITARSPOT.value Then frmMain.QUITARSPOT.value = False


If LUZ_SELECTA > 0 Then
    If LUZ_SELECTA <= UBound(SPOT_LIGHTS) Then
        With MapData(SPOT_LIGHTS(LUZ_SELECTA).Mx, SPOT_LIGHTS(LUZ_SELECTA).My).SPOTLIGHT
            
            .OffsetX = Val(frmMain.SPOT_OFFSETX.Text)
            .OffsetY = Val(frmMain.SPOT_OFFSETY.Text)
            
            .INTENSITY = Val(frmMain.SPOT_INTENSIDAD.Text)
            
            .SPOT_TIPO = frmMain.SPOT_ANIM.ListIndex
            
            .SPOT_COLOR_BASE = frmMain.COLORSPOT.ListIndex + 1
            
            If .SPOT_COLOR_BASE = frmMain.COLORSPOT.ListCount Then
            .Color = Val(frmMain.COLOR_CUSTOM_SPOT.Text)
            
            End If
            
            .SPOT_COLOR_EXTRA = frmMain.COLOREXTRA.ListIndex
            
            If .SPOT_COLOR_EXTRA = frmMain.COLOREXTRA.ListCount - 1 Then
            .COLOR_EXTRA = Val(frmMain.COLOR_CUSTOM_EXTRA.Text)
            End If
            
            If .SPOT_TIPO = 0 Then
                .Grafico = Val(frmMain.GRAFICO_SPOT.Text)
            End If
            .EXTRA_GRAFICO = Val(frmMain.GRAFICO_SPOT_COLOR.Text)
            
            
        End With
        
        With SPOT_LIGHTS(LUZ_SELECTA)
        

            .OffsetX = Val(frmMain.SPOT_OFFSETX.Text)
            .OffsetY = Val(frmMain.SPOT_OFFSETY.Text)
            
            .INTENSITY = Val(frmMain.SPOT_INTENSIDAD.Text)
            
            .SPOT_TIPO = frmMain.SPOT_ANIM.ListIndex
            
            .SPOT_COLOR_BASE = frmMain.COLORSPOT.ListIndex + 1
            
            .SPOT_COLOR_EXTRA = frmMain.COLOREXTRA.ListIndex
            If .SPOT_COLOR_BASE = frmMain.COLORSPOT.ListCount Then
            .Color = Val(frmMain.COLOR_CUSTOM_SPOT.Text)
            End If
            
            
            If .SPOT_COLOR_EXTRA = frmMain.COLOREXTRA.ListCount - 1 Then
            .COLOR_EXTRA = Val(frmMain.COLOR_CUSTOM_EXTRA.Text)
            End If
            
            If .SPOT_TIPO = 0 Then
                .Grafico = Val(frmMain.GRAFICO_SPOT.Text)
            End If
            .EXTRA_GRAFICO = Val(frmMain.GRAFICO_SPOT_COLOR.Text)
        
        End With
        
        
    End If

Else

    MsgBox "debes seleccionar una luz con click derecho"
End If

End Sub

Private Sub TimAutoGuardarMapa_Timer()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************
If mnuAutoGuardarMapas.Checked = True Then
    bAutoGuardarMapaCount = bAutoGuardarMapaCount + 1
    If bAutoGuardarMapaCount >= bAutoGuardarMapa Then
        If MapInfo.Changed = 1 Then ' Solo guardo si el mapa esta modificado
            modMapIO.GuardarMapa Dialog.FileName
        End If
        bAutoGuardarMapaCount = 0
    End If
End If
End Sub


Public Sub ObtenerNombreArchivo(ByVal Guardar As Boolean)
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************

With Dialog
    .Filter = "Mapas de Argentum Online (*.map;*.maptemp)|*.map;*.maptemp"
    If Guardar Then
            .DialogTitle = "Guardar"
            .DefaultExt = ".txt"
            .FileName = vbNullString
            .flags = cdlOFNPathMustExist
            .ShowSave
    Else
        .DialogTitle = "Cargar"
        .FileName = vbNullString
        .flags = cdlOFNFileMustExist

        .ShowOpen

    End If
End With
End Sub

Private Sub tLuz_Change()
If Val(tLuz.Text) <> LuzSelecta Then
    If Val(tLuz.Text) <= lLuces.ListCount - 1 Then
        LuzSelecta = Val(tLuz.Text)
        lLuces.ListIndex = Val(tLuz.Text)

    Else
        tLuz.Text = LuzSelecta
    End If
End If
End Sub

Private Sub txtNumSurface_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If Val(frmMain.txtNumSurface) > 0 Then
    
        DibujarGEnPic frmMain.picsur, Val(frmMain.txtNumSurface), 0, 0
        
    
    
    End If
    End If
End Sub

Private Sub txtNumSurface_LostFocus()
    If Val(frmMain.txtNumSurface) > 0 Then
        If Val(frmMain.txtNumSurface) <= 5499 And Val(frmMain.txtNumSurface) >= 4000 Then
        DibujarGEnPic frmMain.picsur, Val(frmMain.txtNumSurface), 0, 0
        End If
    
    
    End If
End Sub

Private Sub VistaStat_Click()
    If Statz Then
        Statz = False
        VistaStat.Caption = "Stats"
        StatTxt.Visible = False
        PreviewGrh.Visible = True
    Else
        Statz = True
        VistaStat.Caption = "Vista previa"
        PreviewGrh.Visible = False
        StatTxt.Visible = True
    End If
End Sub
