VERSION 5.00
Begin VB.Form frmDonacionesTYC 
   BackColor       =   &H80000000&
   BorderStyle     =   0  'None
   ClientHeight    =   7500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   10500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label btnIrWeb 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   4320
      MouseIcon       =   "frmDonacionesTYC.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   6600
      Width           =   1815
   End
   Begin VB.Label btnCerrar 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   9480
      MouseIcon       =   "frmDonacionesTYC.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   360
      Width           =   375
   End
End
Attribute VB_Name = "frmDonacionesTYC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager

Private Sub btnCerrar_Click()
    frmDonaciones.Show
    Unload Me
End Sub

Private Sub btnIrWeb_Click()
    Call ShellExecute(0, "Open", "https://www.evolutionao.com.uy/Donaciones.html", "", App.path, SW_SHOWNORMAL)
End Sub

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me
End Sub
