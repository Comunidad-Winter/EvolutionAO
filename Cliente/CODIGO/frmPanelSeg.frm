VERSION 5.00
Begin VB.Form frmPanelSeg 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "              Panel"
   ClientHeight    =   915
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   915
   ScaleWidth      =   2175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblVida 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "305/305"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1725
   End
   Begin VB.Label lblMana 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "2050/2052"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   510
      Width           =   1725
   End
   Begin VB.Image ImgMana 
      Height          =   225
      Left            =   195
      Picture         =   "frmPanelSeg.frx":0000
      Top             =   480
      Width           =   1785
   End
   Begin VB.Image ImgVida 
      Height          =   225
      Left            =   195
      Picture         =   "frmPanelSeg.frx":05E5
      Top             =   225
      Width           =   1785
   End
End
Attribute VB_Name = "frmPanelSeg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager
Public LastButtonPressed As clsGraphicalButton

Private Sub Command1_Click()
    Call WriteSeguimiento("1")
    Unload Me
End Sub

Private Sub Form_Load()

    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me

    Set LastButtonPressed = New clsGraphicalButton

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastButtonPressed.ToggleToNormal
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call WriteSeguimiento("1")
    Unload Me
End Sub
