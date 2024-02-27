VERSION 5.00
Begin VB.Form frmConnect 
   BackColor       =   &H80000000&
   BorderStyle     =   0  'None
   Caption         =   "Argentum Online"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00000040&
   BeginProperty Font 
      Name            =   "Matura MT Script Capitals"
      Size            =   23.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConnect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CheckBox chkRecordar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5175
      MaskColor       =   &H00404040&
      TabIndex        =   4
      Top             =   4800
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.TextBox txtPasswd 
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
      Height          =   255
      IMEMode         =   3  'DISABLE
      Left            =   5100
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   4470
      Width           =   1860
   End
   Begin VB.TextBox txtNombre 
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
      Height          =   255
      Left            =   5100
      TabIndex        =   0
      Top             =   3135
      Width           =   1860
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Recordar contraseña"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   5445
      TabIndex        =   5
      Top             =   4785
      Width           =   1455
   End
   Begin VB.Image web 
      Height          =   495
      Left            =   5160
      MouseIcon       =   "frmConnect.frx":000C
      MousePointer    =   99  'Custom
      Top             =   6300
      Width           =   1725
   End
   Begin VB.Image ImgRecuperarCuenta 
      Height          =   405
      Left            =   5040
      MouseIcon       =   "frmConnect.frx":015E
      MousePointer    =   99  'Custom
      Top             =   7800
      Width           =   1935
   End
   Begin VB.Image ImgCrearCuenta 
      Height          =   495
      Left            =   5160
      MouseIcon       =   "frmConnect.frx":02B0
      MousePointer    =   99  'Custom
      Top             =   5700
      Width           =   1725
   End
   Begin VB.Label txtNotificaciones 
      BackStyle       =   0  'Transparent
      Caption         =   "Primer servidor de Uruguay"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   1
      Left            =   9360
      TabIndex        =   3
      Top             =   8160
      Visible         =   0   'False
      Width           =   2280
   End
   Begin VB.Label txtNotificaciones 
      BackStyle       =   0  'Transparent
      Caption         =   "©EvolutionAo, Copyright 2021"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   0
      Left            =   9360
      TabIndex        =   2
      Top             =   8400
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Image imgConectarse 
      Height          =   465
      Left            =   5175
      MouseIcon       =   "frmConnect.frx":0402
      MousePointer    =   99  'Custom
      Top             =   5040
      Width           =   1725
   End
   Begin VB.Image imgSalir 
      Height          =   405
      Left            =   11040
      MouseIcon       =   "frmConnect.frx":0554
      MousePointer    =   99  'Custom
      Top             =   600
      Width           =   345
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub chkRecordar_Click()
    Call Audio.PlayWave(SND_CLICK)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then prgRun = False
End Sub

Private Sub Form_Load()
    If Not Connected Then
        EngineRun = False
        Me.Picture = LoadPicture(App.path & "\graficos\VentanaConectar.jpg")
    End If
End Sub

Private Sub imgConectarse_Click()

    If frmMain.Socket1.Connected Then
        frmMain.Socket1.Disconnect
        frmMain.Socket1.Cleanup
        DoEvents
    End If

    AccountName = txtNombre.Text
    AccountPassword = txtPasswd.Text

    If CheckUserData Then
        EstadoLogin = E_MODO.Normal

        frmMain.Socket1.HostName = CurServerIp
        frmMain.Socket1.RemotePort = CurServerPort
        frmMain.Socket1.Connect

        If chkRecordar.Value Then
            Call SaveRecu(txtNombre.Text, txtPasswd.Text)
        End If
    End If

End Sub

Private Sub ImgCrearCuenta_Click()

    EstadoLogin = E_MODO.CrearCuenta

    If frmMain.Socket1.Connected Then
        frmMain.Socket1.Disconnect
        frmMain.Socket1.Cleanup
        DoEvents
    End If

    frmMain.Socket1.HostName = CurServerIp
    frmMain.Socket1.RemotePort = CurServerPort
    frmMain.Socket1.Connect

End Sub

Private Sub ImgRecuperarCuenta_Click()

    EstadoLogin = E_MODO.RecuperarCuenta

    If frmMain.Socket1.Connected Then
        frmMain.Socket1.Disconnect
        frmMain.Socket1.Cleanup
        DoEvents
    End If

    frmMain.Socket1.HostName = CurServerIp
    frmMain.Socket1.RemotePort = CurServerPort
    frmMain.Socket1.Connect

End Sub

Private Sub imgSalir_Click()
    prgRun = False
End Sub

Private Sub txtNombre_Change()

    If LenB(txtNombre.Text) < 1 Then
        txtPasswd.Text = vbNullString
        Exit Sub
    End If

    If Len(txtNombre.Text) > 4 Then
        Dim ID As Byte
        ID = NickExiste(txtNombre.Text)

        If ID <> 0 Then txtPasswd.Text = Recu(ID).Password
    End If

End Sub

Private Sub txtPasswd_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then imgConectarse_Click
End Sub

Private Sub web_Click()
    Call ShellExecute(0, "Open", "https://www.evolutionao.com.uy", vbNullString, App.path, SW_SHOWNORMAL)
End Sub
