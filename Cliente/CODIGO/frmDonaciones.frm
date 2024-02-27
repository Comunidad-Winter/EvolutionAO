VERSION 5.00
Begin VB.Form frmDonaciones 
   BackColor       =   &H80000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   ScaleHeight     =   500
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label btnInfoTYC 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   600
      MouseIcon       =   "frmDonaciones.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   480
      Width           =   495
   End
   Begin VB.Label lblCerrar 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   9480
      MouseIcon       =   "frmDonaciones.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   360
      Width           =   375
   End
End
Attribute VB_Name = "frmDonaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ClassMove As New clsFormMovementManager
Private clsFormulario As clsFormMovementManager

Private Sub btnInfoTYC_Click()
    frmDonacionesTYC.Show
End Sub

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
End Sub

Private Sub lblCerrar_Click()
    Unload Me
    frmMain.SetFocus
End Sub
