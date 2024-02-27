VERSION 5.00
Begin VB.Form frmMenuSound 
   BorderStyle     =   0  'None
   ClientHeight    =   3000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3000
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox imgChkMusica 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   600
      MouseIcon       =   "frmMenu.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   1660
      Width           =   180
   End
   Begin VB.CheckBox imgChkSonidos 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   600
      MouseIcon       =   "frmMenu.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   910
      Width           =   180
   End
   Begin VB.Image ImgGuardar 
      Height          =   465
      Left            =   3480
      MouseIcon       =   "frmMenu.frx":02A4
      MousePointer    =   99  'Custom
      Top             =   2280
      Width           =   1890
   End
   Begin VB.Image ImgCerrar 
      Height          =   465
      Left            =   600
      MouseIcon       =   "frmMenu.frx":03F6
      MousePointer    =   99  'Custom
      Top             =   2280
      Width           =   1890
   End
End
Attribute VB_Name = "frmMenuSound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Loaded As Boolean
Private RutaConfig As String
Private bMusicActivated As Byte
Private bSoundActivated As Byte

Private Sub Form_Load()

    Loaded = False

    Me.Picture = LoadPicture(App.path & "\Graficos\VentanaConfEfectSonidos.jpg")
    RutaConfig = App.path & "\INIT\Configuracion.dat"

    bMusicActivated = Val(GetVar(RutaConfig, "INIT", "bMusicActivated"))
    bSoundActivated = Val(GetVar(RutaConfig, "INIT", "bSoundActivated"))

    imgChkMusica.Value = bMusicActivated
    imgChkSonidos.Value = bSoundActivated

    Loaded = True

End Sub

Private Sub imgCerrar_Click()
    Unload Me
End Sub

Private Sub imgChkMusica_Click()

    If Not Loaded Then Exit Sub

    Call Audio.PlayWave(SND_CLICK)
    bMusicActivated = IIf(bMusicActivated, 0, 1)
    imgChkMusica.Value = bMusicActivated

End Sub

Private Sub imgChkSonidos_Click()

    If Not Loaded Then Exit Sub

    Call Audio.PlayWave(SND_CLICK)
    bSoundActivated = IIf(bSoundActivated, 0, 1)
    imgChkSonidos.Value = bSoundActivated

End Sub

Private Sub imgGuardar_Click()
    Call WriteVar(RutaConfig, "INIT", "bMusicActivated", bMusicActivated)
    Call WriteVar(RutaConfig, "INIT", "bSoundActivated", bSoundActivated)
    Unload Me
End Sub
