VERSION 5.00
Begin VB.Form frmMenuRanking 
   BackColor       =   &H80000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4065
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   4065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label btnCerrar 
      BackStyle       =   0  'Transparent
      Height          =   220
      Left            =   3700
      MouseIcon       =   "lblMenuRanking.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   240
      Width           =   170
   End
   Begin VB.Label lblFragsCiu 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   960
      MouseIcon       =   "lblMenuRanking.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   4800
      Width           =   2205
   End
   Begin VB.Label lblFragsCri 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   930
      MouseIcon       =   "lblMenuRanking.frx":02A4
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   3990
      Width           =   2205
   End
   Begin VB.Label lvlRetos 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   960
      MouseIcon       =   "lblMenuRanking.frx":03F6
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   2520
      Width           =   2205
   End
   Begin VB.Label lblFrags 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   960
      MouseIcon       =   "lblMenuRanking.frx":0548
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   3240
      Width           =   2085
   End
   Begin VB.Label lblOro 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   960
      MouseIcon       =   "lblMenuRanking.frx":069A
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   1800
      Width           =   2085
   End
   Begin VB.Label lblNivel 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   840
      MouseIcon       =   "lblMenuRanking.frx":07EC
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   960
      Width           =   2205
   End
End
Attribute VB_Name = "frmMenuRanking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager

Private Sub btnCerrar_Click()
    Unload Me
    frmMain.SetFocus
End Sub

Private Sub Form_Load()

    On Error Resume Next

    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
    Me.Picture = LoadPicture(App.path & "\Graficos\VentanaMenuRanking.jpg")
    
End Sub

Private Sub lblFrags_Click()
    Call Audio.PlayWave(SND_CLICK)
    Call WriteSolicitarRanking(TopFrags)
    Unload Me
End Sub

Private Sub lblFragsCiu_Click()
    Call Audio.PlayWave(SND_CLICK)
    Call WriteSolicitarRanking(TopCriminales)
    Unload Me
End Sub

Private Sub lblFragsCri_Click()
    Call Audio.PlayWave(SND_CLICK)
    Call WriteSolicitarRanking(TopCiudadanos)
    Unload Me
End Sub

Private Sub lblNivel_Click()
    Call Audio.PlayWave(SND_CLICK)
    Call WriteSolicitarRanking(TopLevel)
    Unload Me
End Sub

Private Sub lblOro_Click()
    Call Audio.PlayWave(SND_CLICK)
    Call WriteSolicitarRanking(TopOro)
    Unload Me
End Sub

Private Sub lvlRetos_Click()
    Call Audio.PlayWave(SND_CLICK)
    Call WriteSolicitarRanking(TopRetos)
    Unload Me
End Sub
