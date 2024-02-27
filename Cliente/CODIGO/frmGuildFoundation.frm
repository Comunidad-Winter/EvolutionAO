VERSION 5.00
Begin VB.Form frmGuildFoundation 
   BorderStyle     =   0  'None
   Caption         =   "Creaci�n de un Clan"
   ClientHeight    =   3840
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   4050
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   256
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtClanName 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   1815
      Width           =   3345
   End
   Begin VB.TextBox txtWeb 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   2850
      Width           =   3345
   End
   Begin VB.Image imgSiguiente 
      Height          =   375
      Left            =   2280
      MouseIcon       =   "frmGuildFoundation.frx":0000
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   3360
      Width           =   1560
   End
   Begin VB.Image imgCancelar 
      Height          =   375
      Left            =   150
      MouseIcon       =   "frmGuildFoundation.frx":0152
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   3330
      Width           =   1665
   End
End
Attribute VB_Name = "frmGuildFoundation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Evolution Online 0.11.6
'
'Copyright (C) 2002 M�rquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Mat�as Fernando Peque�o
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
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'

Option Explicit
Private clsFormulario As clsFormMovementManager

Private cBotonSiguiente As clsGraphicalButton
Private cBotonCancelar As clsGraphicalButton

Public LastPressed As clsGraphicalButton

Private Sub Form_Deactivate()
    Me.SetFocus
End Sub

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me

    Me.Picture = LoadPicture(App.path & "\graficos\VentanaNombreClan.jpg")

    Call LoadButtons

    If Len(txtClanName.Text) <= 30 Then
        If Not AsciiValidos(txtClanName) Then
            MsgBox "Nombre invalido."
            Exit Sub
        End If
    Else
        MsgBox "Nombre demasiado extenso."
        Exit Sub
    End If

End Sub

Private Sub LoadButtons()
    Dim GrhPath As String

    GrhPath = DirGraficos

    Set cBotonSiguiente = New clsGraphicalButton
    Set cBotonCancelar = New clsGraphicalButton

    Set LastPressed = New clsGraphicalButton
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastPressed.ToggleToNormal
End Sub

Private Sub imgCancelar_Click()
    Unload Me
End Sub

Private Sub imgSiguiente_Click()
    ClanName = txtClanName.Text
    Site = txtWeb.Text
    Unload Me
    frmGuildDetails.Show , frmMain
End Sub

Private Sub txtWeb_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastPressed.ToggleToNormal
End Sub