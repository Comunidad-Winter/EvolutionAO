VERSION 5.00
Begin VB.Form frmManual 
   BackColor       =   &H8000000C&
   BorderStyle     =   0  'None
   ClientHeight    =   7305
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11250
   DrawMode        =   14  'Copy Pen
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   11250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label btnTrabajo 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Height          =   300
      Index           =   3
      Left            =   180
      MouseIcon       =   "frmManual.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   6480
      Width           =   2700
   End
   Begin VB.Label btnTorneo 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Height          =   300
      Index           =   2
      Left            =   180
      MouseIcon       =   "frmManual.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   6000
      Width           =   2700
   End
   Begin VB.Label btnComandos 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Height          =   300
      Index           =   1
      Left            =   180
      MouseIcon       =   "frmManual.frx":02A4
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   5520
      Width           =   2700
   End
   Begin VB.Label btnHogar 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Height          =   300
      Left            =   180
      MouseIcon       =   "frmManual.frx":03F6
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   5040
      Width           =   2700
   End
   Begin VB.Label btnParty 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Height          =   300
      Left            =   180
      MouseIcon       =   "frmManual.frx":0548
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   3105
      Width           =   2700
   End
   Begin VB.Label btnClan 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Height          =   300
      Left            =   180
      MouseIcon       =   "frmManual.frx":069A
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   2625
      Width           =   2700
   End
   Begin VB.Label btnRunas 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Height          =   300
      Left            =   180
      MouseIcon       =   "frmManual.frx":07EC
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   4080
      Width           =   2700
   End
   Begin VB.Label btnInicio 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Height          =   300
      Left            =   180
      MouseIcon       =   "frmManual.frx":093E
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   1185
      Width           =   2700
   End
   Begin VB.Label btnCopas 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Height          =   300
      Index           =   0
      Left            =   180
      MouseIcon       =   "frmManual.frx":0A90
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   4560
      Width           =   2700
   End
   Begin VB.Label btnOro 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Height          =   300
      Left            =   180
      MouseIcon       =   "frmManual.frx":0BE2
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   2145
      Width           =   2700
   End
   Begin VB.Label btnFaccion 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Height          =   300
      Left            =   180
      MouseIcon       =   "frmManual.frx":0D34
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   3585
      Width           =   2700
   End
   Begin VB.Label btnNivel 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Height          =   300
      Index           =   0
      Left            =   180
      MouseIcon       =   "frmManual.frx":0E86
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   1665
      Width           =   2700
   End
   Begin VB.Label btnCerrar 
      BackStyle       =   0  'Transparent
      Height          =   465
      Left            =   10440
      MouseIcon       =   "frmManual.frx":0FD8
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "frmManual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Definimos que nuestro formulario se puede mover.
Private clsFormulario As clsFormMovementManager
Private imgManual() As Picture

Private Sub btnCerrar_Click()
    Unload Me
End Sub

Private Sub btnClan_Click()
    Me.Picture = imgManual(6)
End Sub

Private Sub btnComandos_Click(Index As Integer)
    Me.Picture = imgManual(10)
End Sub

Private Sub btnCopas_Click(Index As Integer)
    Me.Picture = imgManual(4)
End Sub

Private Sub btnHogar_Click()
    Me.Picture = imgManual(8)
End Sub

Private Sub btnInicio_Click()
    'Seteamos la imagen del forndo del formulario (Volver inicio)
    Me.Picture = imgManual(0)
End Sub

Private Sub btnNivel_Click(Index As Integer)
    'Seteamos la imagen del forndo del formulario (nivel 15-24)
    Me.Picture = imgManual(1)
End Sub

Private Sub btnFaccion_Click()
    'Seteamos la imagen del forndo del formulario (nivel 25-34)
    Me.Picture = imgManual(2)
End Sub

Private Sub btnOro_Click()
    'Seteamos la imagen del forndo del formulario (nivel 35-42)
    Me.Picture = imgManual(3)
End Sub

Private Sub btnParty_Click()
    Me.Picture = imgManual(7)
End Sub

Private Sub btnRunas_Click()
    Me.Picture = imgManual(5)
End Sub

Private Sub btnTorneo_Click(Index As Integer)
    Me.Picture = imgManual(9)
End Sub

Private Sub btnTrabajo_Click(Index As Integer)
    Me.Picture = imgManual(11)
End Sub

Private Sub Form_Load()

    On Error GoTo Error

    'Permitimos que el formulario MANUAL se pueda mover.
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me

    Dim Ruta As String
    Ruta = DirGraficos & "manual"

    Dim LoopC As Long

    Do While FileExist(Ruta & LoopC & ".jpg", vbNormal)
        ReDim Preserve imgManual(0 To LoopC) As Picture
        Set imgManual(LoopC) = LoadPicture(Ruta & LoopC & ".jpg")
        LoopC = LoopC + 1
    Loop

    'Imagen de fondo
    Me.Picture = imgManual(0)

    Exit Sub
Error:
    MsgBox Err.Description, vbInformation, "Error: " & Err.Number
    Unload Me
End Sub
