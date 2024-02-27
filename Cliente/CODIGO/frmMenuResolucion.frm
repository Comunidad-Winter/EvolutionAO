VERSION 5.00
Begin VB.Form frmMenuResolucion 
   BorderStyle     =   0  'None
   ClientHeight    =   3000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2970.297
   ScaleMode       =   0  'User
   ScaleWidth      =   5940.594
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox imgChkReso 
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
      Left            =   480
      MouseIcon       =   "frmMenuResolucion.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   1320
      Width           =   180
   End
   Begin VB.Image ImgCerrar 
      Height          =   465
      Left            =   600
      MouseIcon       =   "frmMenuResolucion.frx":0152
      MousePointer    =   99  'Custom
      Top             =   2280
      Width           =   1890
   End
   Begin VB.Image ImgGuardar 
      Height          =   465
      Left            =   3480
      MouseIcon       =   "frmMenuResolucion.frx":02A4
      MousePointer    =   99  'Custom
      Top             =   2280
      Width           =   1890
   End
End
Attribute VB_Name = "frmMenuResolucion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Loaded As Boolean
Private RutaConfig As String
Private bResolucion As Byte

Private Sub Form_Load()

    Loaded = False

    Me.Picture = LoadPicture(App.path & "\Graficos\VentanaConfPantallaCompleta.jpg")
    RutaConfig = App.path & "\INIT\Configuracion.dat"

    bResolucion = Val(GetVar(RutaConfig, "INIT", "bResolucion"))
    imgChkReso.Value = bResolucion

    Loaded = True

End Sub

Private Sub imgCerrar_Click()
    Unload Me
End Sub

Private Sub imgChkReso_Click()

    If Not Loaded Then Exit Sub

    Call Audio.PlayWave(SND_CLICK)
    bResolucion = IIf(bResolucion, 0, 1)
    imgChkReso.Value = bResolucion

End Sub

Private Sub imgGuardar_Click()
    Call WriteVar(RutaConfig, "INIT", "bResolucion", bResolucion)
    Unload Me
End Sub
