VERSION 5.00
Begin VB.Form frmMapa 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   7650
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11760
   ForeColor       =   &H8000000B&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMapa.frx":0000
   ScaleHeight     =   7650
   ScaleWidth      =   11760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image imgToogleMap 
      Height          =   495
      Index           =   1
      Left            =   9600
      MousePointer    =   99  'Custom
      Top             =   3600
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgToogleMap 
      Height          =   495
      Index           =   0
      Left            =   1560
      MousePointer    =   99  'Custom
      Top             =   3840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image imgCerrar 
      Height          =   375
      Left            =   11400
      MouseIcon       =   "frmMapa.frx":655AE
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "frmMapa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager

Private Sub Form_Load()

    On Error Resume Next
    
    Me.Picture = LoadPicture(DirGraficos & "Mapa1.jpg")
    
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me

End Sub

Private Sub imgCerrar_Click()
    Unload Me
    frmMain.SetFocus
End Sub
