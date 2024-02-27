VERSION 5.00
Begin VB.Form frmGuildDetails 
   BorderStyle     =   0  'None
   Caption         =   "Detalles del Clan"
   ClientHeight    =   6810
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   6840
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   454
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   456
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtDesc 
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
      Height          =   1110
      Left            =   675
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   660
      Width           =   5415
   End
   Begin VB.TextBox txtCodex1 
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
      Height          =   195
      Index           =   0
      Left            =   750
      TabIndex        =   1
      Top             =   3297
      Width           =   5325
   End
   Begin VB.TextBox txtCodex1 
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
      Height          =   195
      Index           =   1
      Left            =   750
      TabIndex        =   2
      Top             =   3660
      Width           =   5325
   End
   Begin VB.TextBox txtCodex1 
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
      Height          =   195
      Index           =   2
      Left            =   750
      TabIndex        =   3
      Top             =   4035
      Width           =   5325
   End
   Begin VB.TextBox txtCodex1 
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
      Height          =   195
      Index           =   3
      Left            =   750
      TabIndex        =   4
      Top             =   4410
      Width           =   5325
   End
   Begin VB.TextBox txtCodex1 
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
      Height          =   195
      Index           =   4
      Left            =   750
      TabIndex        =   5
      Top             =   4785
      Width           =   5325
   End
   Begin VB.TextBox txtCodex1 
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
      Height          =   195
      Index           =   5
      Left            =   750
      TabIndex        =   6
      Top             =   5160
      Width           =   5325
   End
   Begin VB.TextBox txtCodex1 
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
      Height          =   195
      Index           =   6
      Left            =   750
      TabIndex        =   7
      Top             =   5535
      Width           =   5325
   End
   Begin VB.TextBox txtCodex1 
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
      Height          =   195
      Index           =   7
      Left            =   750
      TabIndex        =   8
      Top             =   5910
      Width           =   5325
   End
   Begin VB.Image imgConfirmar 
      Height          =   405
      Left            =   4800
      MouseIcon       =   "frmGuildDetails.frx":0000
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   6240
      Width           =   1605
   End
   Begin VB.Image imgSalir 
      Height          =   405
      Left            =   360
      MouseIcon       =   "frmGuildDetails.frx":0152
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   6240
      Width           =   1680
   End
End
Attribute VB_Name = "frmGuildDetails"
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

Private clsFormulario As clsFormMovementManager

Private cBotonConfirmar As clsGraphicalButton
Private cBotonSalir As clsGraphicalButton

Public LastPressed As clsGraphicalButton

Private Const MAX_DESC_LENGTH As Integer = 520
Private Const MAX_CODEX_LENGTH As Integer = 100

Private Sub Form_Load()
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me

    Me.Picture = LoadPicture(App.path & "\graficos\VentanaCodex.jpg")

    Call LoadButtons
End Sub

Private Sub LoadButtons()
    Dim GrhPath As String

    GrhPath = DirGraficos

    Set cBotonConfirmar = New clsGraphicalButton
    Set cBotonSalir = New clsGraphicalButton

    Set LastPressed = New clsGraphicalButton
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastPressed.ToggleToNormal
End Sub

Private Sub imgConfirmar_Click()
    Dim fdesc As String
    Dim Codex() As String
    Dim k As Byte
    Dim Cont As Byte

    fdesc = Replace(txtDesc, vbCrLf, "º", , , vbBinaryCompare)


    Cont = 0
    For k = 0 To txtCodex1.UBound
        If LenB(txtCodex1(k).Text) <> 0 Then Cont = Cont + 1
    Next k

    If Cont < 4 Then
        MsgBox "Debes definir al menos cuatro mandamientos."
        Exit Sub
    End If

    ReDim Codex(txtCodex1.UBound) As String
    For k = 0 To txtCodex1.UBound
        Codex(k) = txtCodex1(k)
    Next k

    If CreandoClan Then
        Call WriteCreateNewGuild(fdesc, ClanName, Site, Codex)
    Else
        Call WriteClanCodexUpdate(fdesc, Codex)
    End If

    CreandoClan = False
    Unload Me
End Sub

Private Sub imgSalir_Click()
    Unload Me
End Sub

Private Sub txtCodex1_Change(Index As Integer)
    If Len(txtCodex1.Item(Index).Text) > MAX_CODEX_LENGTH Then _
       txtCodex1.Item(Index).Text = Left$(txtCodex1.Item(Index).Text, MAX_CODEX_LENGTH)
End Sub

Private Sub txtCodex1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastPressed.ToggleToNormal
End Sub

Private Sub txtDesc_Change()
    If Len(txtDesc.Text) > MAX_DESC_LENGTH Then _
       txtDesc.Text = Left$(txtDesc.Text, MAX_DESC_LENGTH)
End Sub
