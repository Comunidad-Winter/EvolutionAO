VERSION 5.00
Begin VB.Form FrmRanking 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4890
   ClientLeft      =   0
   ClientTop       =   15
   ClientWidth     =   2985
   LinkTopic       =   "Form1"
   ScaleHeight     =   4890
   ScaleWidth      =   2985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2850
      TabIndex        =   0
      Top             =   20
      Width           =   135
   End
   Begin VB.Image ImgNivel 
      Height          =   405
      Left            =   840
      Picture         =   "frmRanking.frx":0000
      Top             =   720
      Width           =   1350
   End
   Begin VB.Image ImgOro 
      Height          =   405
      Left            =   840
      Picture         =   "frmRanking.frx":552B
      Top             =   2520
      Width           =   1350
   End
   Begin VB.Image ImgReto 
      Height          =   405
      Left            =   840
      Picture         =   "frmRanking.frx":AB74
      Top             =   1920
      Width           =   1350
   End
   Begin VB.Image ImgCiudadanos 
      Height          =   405
      Left            =   840
      Picture         =   "frmRanking.frx":100A7
      Top             =   3720
      Width           =   1350
   End
   Begin VB.Image ImgFrags 
      Height          =   405
      Left            =   840
      Picture         =   "frmRanking.frx":15908
      Top             =   1320
      Width           =   1350
   End
   Begin VB.Image ImgCriminales 
      Height          =   405
      Left            =   840
      Picture         =   "frmRanking.frx":1AE72
      Top             =   3120
      Width           =   1350
   End
End
Attribute VB_Name = "FrmRanking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager

Public LastButtonPressed As clsGraphicalButton

Private Sub Form_Load()

    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    
    'Me.Picture = LoadPicture(App.path & "\graficos\VentanaInvocar.jpg")

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub ImgCriminales_Click()
    Call Audio.PlayWave(SND_CLICK)
    Call WriteSolicitarRanking(TopCriminales)
    Unload Me
End Sub

Private Sub ImgFrags_Click()
    Call Audio.PlayWave(SND_CLICK)
    Call WriteSolicitarRanking(TopFrags)
    Unload Me
End Sub

Private Sub ImgNivel_Click()
    Call Audio.PlayWave(SND_CLICK)
    Call WriteSolicitarRanking(TopLevel)
    Unload Me
End Sub

Private Sub ImgCiudadanos_Click()
    Call Audio.PlayWave(SND_CLICK)
    Call WriteSolicitarRanking(TopCiudadanos)
    Unload Me
End Sub

Private Sub ImgReto_Click()
    Call Audio.PlayWave(SND_CLICK)
    Call WriteSolicitarRanking(TopRetos)
    Unload Me
End Sub

Private Sub ImgOro_Click()
    Call Audio.PlayWave(SND_CLICK)
    Call WriteSolicitarRanking(TopOro)
    Unload Me
End Sub

Private Sub Label1_Click()
    Call Audio.PlayWave(SND_CLICK)
    Unload Me
    frmMain.SetFocus
End Sub
