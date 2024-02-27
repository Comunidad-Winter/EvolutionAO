VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Begin VB.Form frmLauncher 
   BackColor       =   &H8000000A&
   BorderStyle     =   0  'None
   Caption         =   "Evolution AO"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9000
   BeginProperty Font 
      Name            =   "Miriam Mono CLM"
      Size            =   8.25
      Charset         =   177
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "frmLauncher.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9000
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstRedditPosts 
      Appearance      =   0  'Flat
      BackColor       =   &H80000012&
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   1830
      Left            =   750
      TabIndex        =   0
      Top             =   6250
      Width           =   3120
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   480
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Image ImgRedes 
      Height          =   375
      Index           =   3
      Left            =   6120
      MouseIcon       =   "frmLauncher.frx":000C
      MousePointer    =   99  'Custom
      ToolTipText     =   "Discord"
      Top             =   7790
      Width           =   495
   End
   Begin VB.Image ImgRedes 
      Height          =   375
      Index           =   2
      Left            =   5280
      MouseIcon       =   "frmLauncher.frx":015E
      MousePointer    =   99  'Custom
      ToolTipText     =   "YouTube"
      Top             =   7790
      Width           =   495
   End
   Begin VB.Image ImgRedes 
      Height          =   375
      Index           =   1
      Left            =   7680
      MouseIcon       =   "frmLauncher.frx":02B0
      MousePointer    =   99  'Custom
      ToolTipText     =   "Facebook"
      Top             =   7790
      Width           =   495
   End
   Begin VB.Image ImgRedes 
      Height          =   375
      Index           =   0
      Left            =   6840
      MouseIcon       =   "frmLauncher.frx":0402
      MousePointer    =   99  'Custom
      ToolTipText     =   "Sitio Web"
      Top             =   7785
      Width           =   495
   End
   Begin VB.Label lblEstado 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "ONLINE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   5140
      TabIndex        =   1
      Top             =   6340
      Width           =   3135
   End
   Begin VB.Image ImgMinimizar 
      Height          =   345
      Left            =   7440
      MouseIcon       =   "frmLauncher.frx":0554
      MousePointer    =   99  'Custom
      Top             =   240
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Image ImgCerrar 
      Height          =   465
      Left            =   8160
      MouseIcon       =   "frmLauncher.frx":06A6
      MousePointer    =   99  'Custom
      Top             =   360
      Width           =   435
   End
   Begin VB.Image ImgMenu 
      Height          =   495
      Index           =   2
      Left            =   600
      MouseIcon       =   "frmLauncher.frx":07F8
      MousePointer    =   99  'Custom
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Image ImgMenu 
      Height          =   495
      Index           =   1
      Left            =   5880
      MouseIcon       =   "frmLauncher.frx":094A
      MousePointer    =   99  'Custom
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Image ImgMenu 
      Height          =   1095
      Index           =   0
      Left            =   3600
      MouseIcon       =   "frmLauncher.frx":0A9C
      MousePointer    =   99  'Custom
      ToolTipText     =   "Jugar"
      Top             =   2640
      Width           =   1695
   End
End
Attribute VB_Name = "frmLauncher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    'morFuego.FileName = App.path & "\Graficos\fuego.gif"
    'morFuego.GifLoop = True
    
    Me.Picture = LoadPicture(App.path & "\Graficos\VentanaLauncher.JPG")

    Call Winsock1.Connect(CurServerIp, CurServerPort)
    lstRedditPosts.Clear

    Dim Ruta As String, LoopC As Long
    Ruta = App.path & "\INIT\Noticias.dat"

    For LoopC = 1 To Val(GetVar(Ruta, "INIT", "MAX_NOTICIAS"))
        lstRedditPosts.AddItem (CStr(GetVar(Ruta, "INIT", LoopC)))
    Next LoopC

    FpsLimiter = Val(GetVar(App.path & "\INIT\Configuracion.dat", "INIT", "bFps"))

End Sub

Private Sub imgCerrar_Click()
    Unload Me
End Sub

Private Sub ImgMenu_Click(Index As Integer)

    Select Case Index
        Case 0
            Unload Me
            Call Main
        Case 1
            Call frmMenuSound.Show(vbModal, frmLauncher)
        Case 2
            Call frmMenuResolucion.Show(vbModal, frmLauncher)
    End Select

End Sub

Private Sub ImgMinimizar_Click()
    WindowState = 1
End Sub

Private Sub ImgRedes_Click(Index As Integer)

    Dim URL As String

    Select Case Index
        Case 0
            URL = "https://www.evolutionao.com.uy"
        Case 1
            URL = "https://www.facebook.com/groups/EvolutionAO"
        Case 2
            URL = "https://www.youtube.com/EvoAO"
        Case 3
            URL = "https://discord.gg/gtwRYZx"
    End Select

    Call ShellExecute(0, "Open", URL, vbNullString, App.path, SW_SHOWNORMAL)

End Sub

Private Sub Winsock1_Connect()
    lblEstado.ForeColor = vbGreen
    lblEstado.Caption = "ONLINE"

    Call Winsock1.Close
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    lblEstado.ForeColor = vbRed
    lblEstado.Caption = "OFFLINE"
End Sub
