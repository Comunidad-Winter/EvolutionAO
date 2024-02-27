VERSION 5.00
Begin VB.Form frmViajes 
   BorderStyle     =   0  'None
   ClientHeight    =   6750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   ScaleHeight     =   6750
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox ListViajes 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "PMingLiU-ExtB"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   3630
      Left            =   1070
      TabIndex        =   0
      Top             =   1440
      Width           =   4680
   End
   Begin VB.Label LblNivel 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "47"
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
      Height          =   195
      Left            =   1550
      TabIndex        =   3
      Top             =   5680
      Width           =   255
   End
   Begin VB.Label LblTiempo 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "60"
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
      Height          =   195
      Left            =   6250
      TabIndex        =   2
      Top             =   5690
      Width           =   255
   End
   Begin VB.Image ImgAceptarr 
      Height          =   495
      Left            =   4800
      MouseIcon       =   "frmViajes.frx":0000
      MousePointer    =   99  'Custom
      Top             =   6120
      Width           =   1815
   End
   Begin VB.Image ImgCancelar 
      Height          =   495
      Left            =   240
      MouseIcon       =   "frmViajes.frx":0152
      MousePointer    =   99  'Custom
      Top             =   6120
      Width           =   1815
   End
   Begin VB.Label LblPrecio 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   360
      Left            =   2475
      TabIndex        =   1
      Top             =   6350
      Width           =   1860
   End
End
Attribute VB_Name = "frmViajes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    Me.Picture = LoadPicture(DirGraficos & "VentanaPasajes.JPG")

    Dim LoopC As Long

    For LoopC = 1 To MAX_CIUDADES
        Call ListViajes.AddItem(Viajes(LoopC).Nombre)
    Next LoopC
    
End Sub

Private Sub ImgAceptarr_Click()

    Dim Index As Integer
    Index = ListViajes.ListIndex + 1

    If Index <> 0 Then
        Call WriteViajar(Viajes(Index).ID)
        Unload Me
    Else
        Call MsgBox("Selecciona un item de la lista")
    End If

End Sub

Private Sub imgCancelar_Click()
    MAX_CIUDADES = 0
    Erase Viajes()
    Unload Me
End Sub

Private Sub ListViajes_Click()

    Dim Index As Integer, Precio As Long

    Index = ListViajes.ListIndex + 1
    Precio = Viajes(Index).Precio

    If Precio <> 0 Then
        LblPrecio.Caption = "$" & Format$(Precio, "##,##")
    Else
        LblPrecio.Caption = "¡Gratis!"
    End If

    LblTiempo.Caption = Viajes(Index).Tiempo
    lblNivel.Caption = Viajes(Index).Nivel

End Sub
