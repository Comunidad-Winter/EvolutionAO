VERSION 5.00
Begin VB.Form frmGuildBrief 
   BorderStyle     =   0  'None
   ClientHeight    =   7650
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   7620
   ClipControls    =   0   'False
   ControlBox      =   0   'False
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   510
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   508
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "&H8000000A&"
   Begin VB.TextBox Desc 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   720
      Left            =   600
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   19
      Top             =   6195
      Width           =   6330
   End
   Begin VB.Image imgSolicitarIngreso 
      Height          =   375
      Left            =   5880
      MouseIcon       =   "frmGuildBrief.frx":0000
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   7170
      Width           =   1455
   End
   Begin VB.Image imgDeclararGuerra 
      Height          =   375
      Left            =   4200
      MouseIcon       =   "frmGuildBrief.frx":0152
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   7170
      Width           =   1575
   End
   Begin VB.Image imgOfrecerAlianza 
      Height          =   375
      Left            =   2640
      MouseIcon       =   "frmGuildBrief.frx":02A4
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   7200
      Width           =   1455
   End
   Begin VB.Image imgOfrecerPaz 
      Height          =   375
      Left            =   1440
      MouseIcon       =   "frmGuildBrief.frx":03F6
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   7200
      Width           =   1095
   End
   Begin VB.Image imgCerrar 
      Height          =   360
      Left            =   120
      MouseIcon       =   "frmGuildBrief.frx":0548
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   7170
      Width           =   1215
   End
   Begin VB.Label Codex 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Index           =   0
      Left            =   675
      TabIndex        =   18
      Top             =   3660
      Width           =   6225
   End
   Begin VB.Label Codex 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Index           =   1
      Left            =   675
      TabIndex        =   17
      Top             =   3900
      Width           =   6225
   End
   Begin VB.Label Codex 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Index           =   2
      Left            =   675
      TabIndex        =   16
      Top             =   4140
      Width           =   6225
   End
   Begin VB.Label Codex 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Index           =   3
      Left            =   675
      TabIndex        =   15
      Top             =   4380
      Width           =   6225
   End
   Begin VB.Label Codex 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Index           =   4
      Left            =   675
      TabIndex        =   14
      Top             =   4620
      Width           =   6225
   End
   Begin VB.Label Codex 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Index           =   5
      Left            =   675
      TabIndex        =   13
      Top             =   4860
      Width           =   6225
   End
   Begin VB.Label Codex 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Index           =   6
      Left            =   675
      TabIndex        =   12
      Top             =   5100
      Width           =   6225
   End
   Begin VB.Label Codex 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Index           =   7
      Left            =   675
      TabIndex        =   11
      Top             =   5340
      Width           =   6225
   End
   Begin VB.Label nombre 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Left            =   1500
      TabIndex        =   10
      Top             =   690
      Width           =   4695
   End
   Begin VB.Label fundador 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Left            =   1620
      TabIndex        =   9
      Top             =   1020
      Width           =   2775
   End
   Begin VB.Label creacion 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Left            =   5985
      TabIndex        =   8
      Top             =   1005
      Width           =   1455
   End
   Begin VB.Label lider 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Left            =   1245
      TabIndex        =   7
      Top             =   1350
      Width           =   3135
   End
   Begin VB.Label web 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Left            =   1545
      TabIndex        =   6
      Top             =   1710
      Width           =   2655
   End
   Begin VB.Label Miembros 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Left            =   5325
      TabIndex        =   5
      Top             =   1695
      Width           =   1935
   End
   Begin VB.Label eleccion 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Left            =   5355
      TabIndex        =   4
      Top             =   1365
      Width           =   1815
   End
   Begin VB.Label lblAlineacion 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Left            =   5400
      TabIndex        =   3
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Enemigos 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Left            =   2220
      TabIndex        =   2
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label Aliados 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Left            =   2010
      TabIndex        =   1
      Top             =   2370
      Width           =   1575
   End
   Begin VB.Label antifaccion 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Left            =   2640
      TabIndex        =   0
      Top             =   2715
      Width           =   2415
   End
End
Attribute VB_Name = "frmGuildBrief"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Evolution Online 0.11.6
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Evolution Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'

Option Explicit

Private clsFormulario As clsFormMovementManager

Private cBotonGuerra As clsGraphicalButton
Private cBotonAlianza As clsGraphicalButton
Private cBotonPaz As clsGraphicalButton
Private cBotonSolicitarIngreso As clsGraphicalButton
Private cBotonCerrar As clsGraphicalButton

Public LastPressed As clsGraphicalButton

Public EsLeader As Boolean

Private Sub Desc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastPressed.ToggleToNormal
End Sub

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me

    Me.Picture = LoadPicture(App.path & "\graficos\VentanaDetallesClan.jpg")

    Call LoadButtons

End Sub

Private Sub LoadButtons()
    Dim GrhPath As String

    GrhPath = DirGraficos

    Set cBotonGuerra = New clsGraphicalButton
    Set cBotonAlianza = New clsGraphicalButton
    Set cBotonPaz = New clsGraphicalButton
    Set cBotonSolicitarIngreso = New clsGraphicalButton
    Set cBotonCerrar = New clsGraphicalButton

    Set LastPressed = New clsGraphicalButton

    If Not EsLeader Then
        imgDeclararGuerra.Visible = False
        imgOfrecerAlianza.Visible = False
        imgOfrecerPaz.Visible = False
    End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastPressed.ToggleToNormal
End Sub

Private Sub imgCerrar_Click()
    Unload Me
End Sub

Private Sub imgDeclararGuerra_Click()
    Call WriteGuildDeclareWar(Nombre.Caption)
    Unload Me
End Sub

Private Sub imgOfrecerAlianza_Click()
    frmCommet.Nombre = Nombre.Caption
    frmCommet.T = Tipo.ALIANZA
    Call frmCommet.Show(vbModal, frmGuildBrief)
End Sub

Private Sub imgOfrecerPaz_Click()
    frmCommet.Nombre = Nombre.Caption
    frmCommet.T = Tipo.PAZ
    Call frmCommet.Show(vbModal, frmGuildBrief)
End Sub

Private Sub imgSolicitarIngreso_Click()
    Call frmGuildSol.RecieveSolicitud(Nombre.Caption)
    Call frmGuildSol.Show(vbModal, frmGuildBrief)
End Sub
