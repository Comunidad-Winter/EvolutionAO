VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOpciones 
   BackColor       =   &H8000000A&
   BorderStyle     =   0  'None
   ClientHeight    =   7185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4830
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
   Icon            =   "frmOpciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   479
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   322
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      ForeColor       =   &H0080C0FF&
      Height          =   615
      ItemData        =   "frmOpciones.frx":0152
      Left            =   1800
      List            =   "frmOpciones.frx":015F
      TabIndex        =   4
      Top             =   4380
      Width           =   975
   End
   Begin VB.TextBox txtCantMensajes 
      Alignment       =   2  'Center
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
      Height          =   240
      Left            =   2310
      MaxLength       =   1
      TabIndex        =   3
      Text            =   "5"
      Top             =   2430
      Width           =   255
   End
   Begin VB.TextBox txtLevel 
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
      Height          =   225
      Left            =   3495
      MaxLength       =   2
      TabIndex        =   2
      Text            =   "40"
      Top             =   4125
      Width           =   255
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   255
      Index           =   0
      Left            =   1455
      TabIndex        =   0
      Top             =   1140
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Max             =   100
      TickStyle       =   3
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   255
      Index           =   1
      Left            =   1455
      TabIndex        =   1
      Top             =   1410
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      LargeChange     =   10
      Max             =   100
      TickStyle       =   3
   End
   Begin VB.Label btnDiscord 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   2640
      MouseIcon       =   "frmOpciones.frx":0185
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   6600
      Width           =   1815
   End
   Begin VB.Image imgChkDesactivarFragShooter 
      Height          =   225
      Left            =   435
      Top             =   4740
      Width           =   210
   End
   Begin VB.Image imgChkAlMorir 
      Height          =   225
      Left            =   435
      Top             =   4425
      Width           =   210
   End
   Begin VB.Image imgChkRequiredLvl 
      Height          =   225
      Left            =   435
      Top             =   4110
      Width           =   210
   End
   Begin VB.Image imgChkNoMostrarNews 
      Height          =   225
      Left            =   2475
      Top             =   3315
      Width           =   210
   End
   Begin VB.Image imgChkMostrarNews 
      Height          =   225
      Left            =   435
      Top             =   3315
      Width           =   210
   End
   Begin VB.Image imgChkPantalla 
      Height          =   225
      Left            =   1950
      Top             =   2430
      Width           =   210
   End
   Begin VB.Image imgChkConsola 
      Height          =   225
      Left            =   435
      Top             =   2430
      Width           =   210
   End
   Begin VB.Image imgChkEfectosSonido 
      Height          =   225
      Left            =   435
      Top             =   1680
      Width           =   210
   End
   Begin VB.Image imgChkSonidos 
      Height          =   225
      Left            =   435
      Top             =   1395
      Width           =   210
   End
   Begin VB.Image imgChkMusica 
      Height          =   225
      Left            =   435
      Top             =   1125
      Width           =   210
   End
   Begin VB.Image imgTutorial 
      Height          =   330
      Left            =   2520
      MouseIcon       =   "frmOpciones.frx":02D7
      MousePointer    =   99  'Custom
      Top             =   5880
      Width           =   1935
   End
   Begin VB.Image imgSoporte 
      Height          =   330
      Left            =   960
      MouseIcon       =   "frmOpciones.frx":0429
      MousePointer    =   99  'Custom
      Top             =   360
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Image imgRadio 
      Height          =   330
      Left            =   1680
      MouseIcon       =   "frmOpciones.frx":057B
      MousePointer    =   99  'Custom
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Image btnWeb 
      Height          =   330
      Left            =   240
      MouseIcon       =   "frmOpciones.frx":06CD
      MousePointer    =   99  'Custom
      Top             =   6600
      Width           =   2160
   End
   Begin VB.Image imgCambiarPasswd 
      Height          =   330
      Left            =   240
      MouseIcon       =   "frmOpciones.frx":081F
      MousePointer    =   99  'Custom
      Top             =   240
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Image imgMapa 
      Height          =   330
      Left            =   360
      MouseIcon       =   "frmOpciones.frx":0971
      MousePointer    =   99  'Custom
      Top             =   5880
      Width           =   2040
   End
   Begin VB.Image imgMsgPersonalizado 
      Height          =   330
      Left            =   2520
      MouseIcon       =   "frmOpciones.frx":0AC3
      MousePointer    =   99  'Custom
      Top             =   5280
      Width           =   1935
   End
   Begin VB.Image imgConfigTeclas 
      Height          =   330
      Left            =   360
      MouseIcon       =   "frmOpciones.frx":0C15
      MousePointer    =   99  'Custom
      Top             =   5280
      Width           =   2040
   End
   Begin VB.Image imgSalir 
      Height          =   390
      Left            =   4320
      MouseIcon       =   "frmOpciones.frx":0D67
      MousePointer    =   99  'Custom
      Top             =   120
      Width           =   435
   End
End
Attribute VB_Name = "frmOpciones"
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

Private RutaConfig As String
Private clsFormulario As clsFormMovementManager

Private cBotonConfigTeclas As clsGraphicalButton
Private cBotonMsgPersonalizado As clsGraphicalButton
Private cBotonMapa As clsGraphicalButton
Private cBotonCambiarPasswd As clsGraphicalButton
Private cBotonManual As clsGraphicalButton
Private cBotonRadio As clsGraphicalButton
Private cBotonSoporte As clsGraphicalButton
Private cBotonTutorial As clsGraphicalButton
Private cBotonSalir As clsGraphicalButton

Public LastPressed As clsGraphicalButton

Private picCheckBox As Picture

Private bMusicActivated As Boolean
Private bSoundActivated As Boolean
Private bSoundEffectsActivated As Boolean

Private loading As Boolean

Private Sub btnManual_Click()
End Sub

Private Sub btnDiscord_Click()
    Call ShellExecute(0, "Open", "https://discord.gg/gtwRYZx", vbNullString, App.path, SW_SHOWNORMAL)
End Sub

Private Sub btnWeb_Click()
    Call ShellExecute(0, "Open", "https://www.evolutionao.com.uy", vbNullString, App.path, SW_SHOWNORMAL)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastPressed.ToggleToNormal
End Sub

Private Sub imgCambiarPasswd_Click()
    Call frmNewPassword.Show(vbModal, Me)
End Sub

Private Sub imgChkAlMorir_Click()
    ClientSetup.bDie = Not ClientSetup.bDie

    If ClientSetup.bDie Then
        imgChkAlMorir.Picture = picCheckBox
    Else
        Set imgChkAlMorir.Picture = Nothing
    End If
End Sub

Private Sub imgChkDesactivarFragShooter_Click()
    ClientSetup.bActive = Not ClientSetup.bActive

    If ClientSetup.bActive Then
        Set imgChkDesactivarFragShooter.Picture = Nothing
    Else
        imgChkDesactivarFragShooter.Picture = picCheckBox
    End If
End Sub

Private Sub imgChkRequiredLvl_Click()
    ClientSetup.bKill = Not ClientSetup.bKill

    If ClientSetup.bKill Then
        imgChkRequiredLvl.Picture = picCheckBox
    Else
        Set imgChkRequiredLvl.Picture = Nothing
    End If
End Sub

Private Sub imgTutorial_Click()
    Call frmManual.Show
End Sub

Private Sub List1_Click()

    Select Case List1.ListIndex
        Case 0
            FpsLimiter = 16
        Case 1
            FpsLimiter = 10
        Case 2
            FpsLimiter = 0
    End Select

    Call WriteVar(RutaConfig, "INIT", "bFps", FpsLimiter)

End Sub

Private Sub txtCantMensajes_Change()

    txtCantMensajes.Text = Val(txtCantMensajes.Text)

    If txtCantMensajes.Text > 0 Then
        DialogosClanes.CantidadDialogos = txtCantMensajes.Text
    Else
        txtCantMensajes.Text = 5
    End If

End Sub

Private Sub txtLevel_Change()
    If Not IsNumeric(txtLevel) Then txtLevel = 0
    txtLevel = Trim$(txtLevel)
    ClientSetup.byMurderedLevel = CByte(txtLevel)
End Sub

Private Sub imgChkConsola_Click()
    DialogosClanes.Activo = False

    imgChkConsola.Picture = picCheckBox
    Set imgChkPantalla.Picture = Nothing
End Sub

Private Sub imgChkEfectosSonido_Click()

    If loading Then Exit Sub

    Call Audio.PlayWave(SND_CLICK)

    bSoundEffectsActivated = Not bSoundEffectsActivated

    Audio.SoundEffectsActivated = bSoundEffectsActivated

    If bSoundEffectsActivated Then
        imgChkEfectosSonido.Picture = picCheckBox
    Else
        Set imgChkEfectosSonido.Picture = Nothing
    End If

End Sub

Private Sub imgChkMostrarNews_Click()
    ClientSetup.bGuildNews = True

    imgChkMostrarNews.Picture = picCheckBox
    Set imgChkNoMostrarNews.Picture = Nothing
End Sub

Private Sub imgChkMusica_Click()

    If loading Then Exit Sub

    Call Audio.PlayWave(SND_CLICK)

    bMusicActivated = Not bMusicActivated

    If Not bMusicActivated Then
        Audio.MusicActivated = False
        Slider1(0).Enabled = False
        Set imgChkMusica.Picture = Nothing
    Else
        If Not Audio.MusicActivated Then  'Prevent the music from reloading
            Audio.MusicActivated = True
            Slider1(0).Enabled = True
            Slider1(0).Value = Audio.MusicVolume
        End If

        imgChkMusica.Picture = picCheckBox
    End If

End Sub

Private Sub imgChkSonidos_Click()

    If loading Then Exit Sub

    Call Audio.PlayWave(SND_CLICK)

    bSoundActivated = Not bSoundActivated

    If Not bSoundActivated Then
        Audio.SoundActivated = False
        RainBufferIndex = 0
        frmMain.IsPlaying = PlayLoop.plNone
        Slider1(1).Enabled = False

        Set imgChkSonidos.Picture = Nothing
    Else
        Audio.SoundActivated = True
        Slider1(1).Enabled = True
        Slider1(1).Value = Audio.SoundVolume

        imgChkSonidos.Picture = picCheckBox
    End If
End Sub

Private Sub imgChkNoMostrarNews_Click()
    ClientSetup.bGuildNews = False

    imgChkNoMostrarNews.Picture = picCheckBox
    Set imgChkMostrarNews.Picture = Nothing
End Sub

Private Sub imgChkPantalla_Click()
    DialogosClanes.Activo = True

    imgChkPantalla.Picture = picCheckBox
    Set imgChkConsola.Picture = Nothing
End Sub

Private Sub imgConfigTeclas_Click()
    If Not loading Then _
       Call Audio.PlayWave(SND_CLICK)
    Call frmCustomKeys.Show(vbModal, Me)
End Sub

Private Sub imgManual_Click()
    
End Sub

Private Sub imgMapa_Click()
    Call frmMapa.Show(vbModal, Me)
End Sub

Private Sub imgMsgPersonalizado_Click()
    Call frmMessageTxt.Show(vbModeless, Me)
End Sub

Private Sub imgRadio_Click()
    If Not loading Then _
       Call Audio.PlayWave(SND_CLICK)

    Call ShellExecute(0, "Open", "https://www.evolutionao.com/radio.html", "", App.path, SW_SHOWNORMAL)
End Sub

Private Sub imgSalir_Click()
    Unload Me
    frmMain.SetFocus
End Sub

Private Sub imgSoporte_Click()

    If Not loading Then _
       Call Audio.PlayWave(SND_CLICK)

    Call ShellExecute(0, "Open", "https://www.evolutionao.com/Contacto.html", "", App.path, SW_SHOWNORMAL)
End Sub

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me

    Me.Picture = LoadPicture(App.path & "\graficos\VentanaOpciones.jpg")
    LoadButtons

    loading = True      'Prevent sounds when setting check's values
    LoadUserConfig
    loading = False     'Enable sounds when setting check's values
End Sub

Private Sub LoadButtons()
    Dim GrhPath As String

    GrhPath = DirGraficos

    Set cBotonConfigTeclas = New clsGraphicalButton
    Set cBotonMsgPersonalizado = New clsGraphicalButton
    Set cBotonMapa = New clsGraphicalButton
    Set cBotonCambiarPasswd = New clsGraphicalButton
    Set cBotonManual = New clsGraphicalButton
    Set cBotonRadio = New clsGraphicalButton
    Set cBotonSoporte = New clsGraphicalButton
    Set cBotonSalir = New clsGraphicalButton

    Set LastPressed = New clsGraphicalButton

    Set picCheckBox = LoadPicture(GrhPath & "CheckBoxOpciones.jpg")
End Sub

Private Sub LoadUserConfig()

    ' Load music config
    bMusicActivated = Audio.MusicActivated
    Slider1(0).Enabled = bMusicActivated

    If bMusicActivated Then
        imgChkMusica.Picture = picCheckBox

        Slider1(0).Value = Audio.MusicVolume
    End If


    ' Load Sound config
    bSoundActivated = Audio.SoundActivated
    Slider1(1).Enabled = bSoundActivated

    If bSoundActivated Then
        imgChkSonidos.Picture = picCheckBox

        Slider1(1).Value = Audio.SoundVolume
    End If

    RutaConfig = App.path & "\INIT\Configuracion.dat"

    ' Load Sound Effects config
    bSoundEffectsActivated = Audio.SoundEffectsActivated
    If bSoundEffectsActivated Then imgChkEfectosSonido.Picture = picCheckBox

    txtCantMensajes.Text = CStr(DialogosClanes.CantidadDialogos)

    If DialogosClanes.Activo Then
        imgChkPantalla.Picture = picCheckBox
    Else
        imgChkConsola.Picture = picCheckBox
    End If

    If ClientSetup.bGuildNews Then
        imgChkMostrarNews.Picture = picCheckBox
    Else
        imgChkNoMostrarNews.Picture = picCheckBox
    End If

    If ClientSetup.bKill Then imgChkRequiredLvl.Picture = picCheckBox
    If ClientSetup.bDie Then imgChkAlMorir.Picture = picCheckBox
    If Not ClientSetup.bActive Then imgChkDesactivarFragShooter.Picture = picCheckBox

    txtLevel = ClientSetup.byMurderedLevel
End Sub

Private Sub Slider1_Change(Index As Integer)
    Select Case Index
        Case 0
            Audio.MusicVolume = Slider1(0).Value
        Case 1
            Audio.SoundVolume = Slider1(1).Value
    End Select
End Sub

Private Sub Slider1_Scroll(Index As Integer)
    Select Case Index
        Case 0
            Audio.MusicVolume = Slider1(0).Value
        Case 1
            Audio.SoundVolume = Slider1(1).Value
    End Select
End Sub
