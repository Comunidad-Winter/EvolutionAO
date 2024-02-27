VERSION 5.00
Begin VB.Form frmGuildMember 
   BackColor       =   &H8000000B&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5640
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   ScaleHeight     =   376
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   399
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstMiembros 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   2175
      Left            =   3315
      TabIndex        =   3
      Top             =   870
      Width           =   2190
   End
   Begin VB.ListBox lstClanes 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   2175
      Left            =   390
      TabIndex        =   2
      Top             =   870
      Width           =   2190
   End
   Begin VB.TextBox txtSearch 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   210
      Left            =   300
      TabIndex        =   1
      Top             =   3664
      Width           =   2400
   End
   Begin VB.Label lblCantMiembros 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "Miriam CLM"
         Size            =   9
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   195
      Left            =   4665
      TabIndex        =   0
      Top             =   3540
      Width           =   360
   End
   Begin VB.Image imgCerrar 
      Height          =   495
      Left            =   3240
      MouseIcon       =   "frmGuildMember.frx":0000
      MousePointer    =   99  'Custom
      Top             =   4920
      Width           =   2295
   End
   Begin VB.Image imgNoticias 
      Height          =   495
      Left            =   390
      MouseIcon       =   "frmGuildMember.frx":0152
      MousePointer    =   99  'Custom
      Top             =   4920
      Width           =   2175
   End
   Begin VB.Image imgDetalles 
      Height          =   375
      Left            =   390
      MouseIcon       =   "frmGuildMember.frx":02A4
      MousePointer    =   99  'Custom
      Top             =   4200
      Width           =   2175
   End
End
Attribute VB_Name = "frmGuildMember"
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

Private cBotonNoticias As clsGraphicalButton
Private cBotonDetalles As clsGraphicalButton
Private cBotonCerrar As clsGraphicalButton

Public LastPressed As clsGraphicalButton

Private Sub Form_Load()

    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me

    Me.Picture = LoadPicture(DirGraficos & "VentanaMiembroClan.jpg")

    Call LoadButtons

End Sub

Private Sub LoadButtons()
    Dim GrhPath As String

    GrhPath = DirGraficos

    Set cBotonNoticias = New clsGraphicalButton
    Set cBotonDetalles = New clsGraphicalButton
    Set cBotonCerrar = New clsGraphicalButton

    Set LastPressed = New clsGraphicalButton

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastPressed.ToggleToNormal
End Sub

Private Sub imgCerrar_Click()
    Unload Me
End Sub

Private Sub imgDetalles_Click()
    If lstClanes.ListIndex = -1 Then Exit Sub

    frmGuildBrief.EsLeader = False

    Call WriteGuildRequestDetails(lstClanes.List(lstClanes.ListIndex))
End Sub

Private Sub imgNoticias_Click()
    Call WriteShowGuildNews
End Sub

Private Sub txtSearch_Change()
    Call FiltrarListaClanes(txtSearch.Text)
End Sub

Private Sub txtSearch_GotFocus()
    With txtSearch
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Public Sub FiltrarListaClanes(ByRef sCompare As String)

    Dim lIndex As Long

    If UBound(GuildNames) <> 0 Then
        With lstClanes
            'Limpio la lista
            .Clear

            .Visible = False

            ' Recorro los arrays
            For lIndex = 0 To UBound(GuildNames)
                ' Si coincide con los patrones
                If InStr(1, UCase$(GuildNames(lIndex)), UCase$(sCompare)) Then
                    ' Lo agrego a la lista
                    .AddItem GuildNames(lIndex)
                End If
            Next lIndex

            .Visible = True
        End With
    End If

End Sub

