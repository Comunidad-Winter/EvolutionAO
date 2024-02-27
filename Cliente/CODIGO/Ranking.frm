VERSION 5.00
Begin VB.Form frmRanking 
   BackColor       =   &H8000000B&
   BorderStyle     =   0  'None
   ClientHeight    =   7500
   ClientLeft      =   6825
   ClientTop       =   1050
   ClientWidth     =   6000
   DrawMode        =   12  'Nop
   LinkTopic       =   "Ranking"
   ScaleHeight     =   500
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblValor 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Index           =   9
      Left            =   4515
      TabIndex        =   18
      Top             =   6360
      Width           =   1605
   End
   Begin VB.Label lblNombre 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      Caption         =   "Soy un trabajador"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Index           =   9
      Left            =   3120
      TabIndex        =   17
      Top             =   6390
      Width           =   1605
   End
   Begin VB.Label lblValor 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Index           =   6
      Left            =   1740
      TabIndex        =   16
      Top             =   6390
      Width           =   1605
   End
   Begin VB.Label lblNombre 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      Caption         =   "Soy un trabajador"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Index           =   6
      Left            =   270
      TabIndex        =   15
      Top             =   6390
      Width           =   1605
   End
   Begin VB.Label lblValor 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   255
      Index           =   3
      Left            =   3375
      TabIndex        =   14
      Top             =   3900
      Width           =   1605
   End
   Begin VB.Label lblNombre 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      Caption         =   "Soy un trabajador"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   255
      Index           =   3
      Left            =   1920
      TabIndex        =   13
      Top             =   3900
      Width           =   1605
   End
   Begin VB.Label lblValor 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Index           =   7
      Left            =   4515
      TabIndex        =   12
      Top             =   4830
      Width           =   1605
   End
   Begin VB.Label lblValor 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Index           =   8
      Left            =   4515
      TabIndex        =   11
      Top             =   5625
      Width           =   1605
   End
   Begin VB.Label lblNombre 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      Caption         =   "Soy un trabajador"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Index           =   8
      Left            =   3120
      TabIndex        =   10
      Top             =   5625
      Width           =   1605
   End
   Begin VB.Label lblValor 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Index           =   5
      Left            =   1740
      TabIndex        =   9
      Top             =   5625
      Width           =   1605
   End
   Begin VB.Label lblNombre 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      Caption         =   "Soy un trabajador"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Index           =   5
      Left            =   270
      TabIndex        =   8
      Top             =   5625
      Width           =   1605
   End
   Begin VB.Label lblValor 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   2
      Left            =   3375
      TabIndex        =   7
      Top             =   2955
      Width           =   1605
   End
   Begin VB.Label lblNombre 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      Caption         =   "Soy un trabajador"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   2
      Left            =   1920
      TabIndex        =   6
      Top             =   2955
      Width           =   1605
   End
   Begin VB.Label lblNombre 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      Caption         =   "Soy un trabajador"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Index           =   7
      Left            =   3120
      TabIndex        =   5
      Top             =   4830
      Width           =   1605
   End
   Begin VB.Label lblValor 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Index           =   4
      Left            =   1740
      TabIndex        =   4
      Top             =   4830
      Width           =   1605
   End
   Begin VB.Label lblNombre 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      Caption         =   "Soy un trabajador"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Index           =   4
      Left            =   270
      TabIndex        =   3
      Top             =   4830
      Width           =   1605
   End
   Begin VB.Label lblValor 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   1
      Left            =   3375
      TabIndex        =   2
      Top             =   1980
      Width           =   1605
   End
   Begin VB.Label lblNombre 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      Caption         =   "Soy un trabajador"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   1
      Top             =   1980
      Width           =   1605
   End
   Begin VB.Label lblCerrarRanking 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   2160
      MouseIcon       =   "Ranking.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   6960
      Width           =   1695
   End
End
Attribute VB_Name = "frmRanking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsFormulario As clsFormMovementManager

Private Sub Form_Load()

    On Error Resume Next

    Select Case TopPressed
        Case TopFrags
            Me.Picture = LoadPicture(DirGraficos & "RankingFrags.jpg")
        Case TopOro
            Me.Picture = LoadPicture(DirGraficos & "RankingOro.jpg")
        Case TopLevel
            Me.Picture = LoadPicture(DirGraficos & "RankingNiveles.jpg")
        Case TopRetos
            Me.Picture = LoadPicture(DirGraficos & "RankingRetos.jpg")
        Case TopCriminales
            Me.Picture = LoadPicture(DirGraficos & "RankingFragsCriminales.jpg")
        Case TopCiudadanos
            Me.Picture = LoadPicture(DirGraficos & "RankingFragsCiudadanos.jpg")
    End Select
    
    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me

End Sub

Private Sub lblCerrarRanking_Click()
    Call frmMenuRanking.Show
    Unload Me
End Sub

