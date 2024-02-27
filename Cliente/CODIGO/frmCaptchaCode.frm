VERSION 5.00
Begin VB.Form frmCaptchaCode 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   ClientHeight    =   2415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3090
   LinkTopic       =   "Form1"
   ScaleHeight     =   2415
   ScaleWidth      =   3090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pCaptcha 
      Appearance      =   0  'Flat
      BackColor       =   &H00101010&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   600
      ScaleHeight     =   345
      ScaleWidth      =   1905
      TabIndex        =   1
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox txtCaptcha 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000012&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      IMEMode         =   3  'DISABLE
      Left            =   600
      MaxLength       =   4
      TabIndex        =   0
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Image ImgAceptarr 
      Height          =   495
      Left            =   720
      MouseIcon       =   "frmCaptchaCode.frx":0000
      MousePointer    =   99  'Custom
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label lCodigo 
      BackStyle       =   0  'Transparent
      Caption         =   "Ingrese el código aquí:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Código de confirmación"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Aceptar"
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   1320
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
End
Attribute VB_Name = "frmCaptchaCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    
    Dim LoopC As Long
    
    For LoopC = 0 To 3
        CaptchaCode(LoopC) = RandomNumber(1, 9)
    Next LoopC
    
End Sub

Private Sub ImgAceptarr_Click()
    
    If LCase$(txtCaptcha.Text) <> (Chr$(CaptchaCode(0)) & Chr$(CaptchaCode(1)) & Chr$(CaptchaCode(2)) & Chr$(CaptchaCode(3))) Then
        MsgBox "Los códigos de confirmación no coinciden. Vuelva a ingresarlo."
        txtCaptcha.Text = vbNullString
        txtCaptcha.SetFocus
        Exit Sub
    End If
    
End Sub

Private Sub pCaptcha_Click()
    Call UpdateCaptcha
End Sub

Private Sub pCaptcha_Paint()
    Call UpdateCaptcha
End Sub

Private Sub Form_Resize()
    Call UpdateCaptcha
End Sub

Public Sub UpdateCaptcha()

    Dim LoopC As Long
    
    For LoopC = 0 To 3
        CaptchaCode(LoopC) = RandomNumber(1, 9)
    Next LoopC

    pCaptcha.Cls
    pCaptcha.Line (RandomNumber(1, 30), RandomNumber(1, pCaptcha.ScaleHeight))-(RandomNumber(pCaptcha.ScaleWidth - 30, pCaptcha.ScaleHeight), RandomNumber(1, pCaptcha.ScaleHeight)), RGB(RandomNumber(60, 255), RandomNumber(60, 255), RandomNumber(40, 200))
    pCaptcha.Line (RandomNumber(pCaptcha.ScaleWidth, 30), RandomNumber(1, pCaptcha.ScaleHeight))-(RandomNumber(pCaptcha.ScaleHeight - 10, pCaptcha.ScaleHeight), RandomNumber(1, pCaptcha.ScaleHeight)), RGB(RandomNumber(60, 255), RandomNumber(60, 255), RandomNumber(40, 200))
    pCaptcha.Line (RandomNumber(1, 30), RandomNumber(1, pCaptcha.ScaleWidth))-(RandomNumber(pCaptcha.ScaleWidth - 30, pCaptcha.ScaleHeight), RandomNumber(1, pCaptcha.ScaleHeight)), RGB(RandomNumber(60, 255), RandomNumber(60, 255), RandomNumber(40, 200))
    pCaptcha.Line (RandomNumber(pCaptcha.ScaleWidth, 30), RandomNumber(1, pCaptcha.ScaleHeight))-(RandomNumber(pCaptcha.ScaleWidth - 20, pCaptcha.ScaleWidth), RandomNumber(1, pCaptcha.ScaleHeight)), RGB(RandomNumber(60, 255), RandomNumber(60, 255), RandomNumber(40, 255))
    pCaptcha.Line (RandomNumber(1, 30), RandomNumber(1, pCaptcha.ScaleWidth))-(RandomNumber(pCaptcha.ScaleWidth - 30, pCaptcha.ScaleHeight), RandomNumber(1, pCaptcha.ScaleHeight)), RGB(RandomNumber(60, 255), RandomNumber(60, 255), RandomNumber(40, 255))
    pCaptcha.CurrentX = (pCaptcha.ScaleWidth / 2) - RandomNumber(300, 400)
    pCaptcha.CurrentY = (pCaptcha.ScaleHeight / 2) - RandomNumber(140, 170)
    pCaptcha.ForeColor = RGB(RandomNumber(60, 255), RandomNumber(60, 255), RandomNumber(60, 255))
    pCaptcha.Print Chr(CaptchaCode(0))
    pCaptcha.CurrentX = (pCaptcha.ScaleWidth / 2) - RandomNumber(-60, 100)
    pCaptcha.CurrentY = (pCaptcha.ScaleHeight / 2) - RandomNumber(140, 170)
    pCaptcha.ForeColor = RGB(RandomNumber(60, 255), RandomNumber(60, 255), RandomNumber(60, 255))
    pCaptcha.Print Chr(CaptchaCode(2))
    pCaptcha.CurrentX = (pCaptcha.ScaleWidth / 2) - RandomNumber(-100, -200)
    pCaptcha.CurrentY = (pCaptcha.ScaleHeight / 2) - RandomNumber(140, 170)
    pCaptcha.ForeColor = RGB(RandomNumber(60, 255), RandomNumber(60, 255), RandomNumber(60, 255))
    pCaptcha.Print Chr(CaptchaCode(3))
    pCaptcha.CurrentX = (pCaptcha.ScaleWidth / 2) - RandomNumber(150, 200)
    pCaptcha.CurrentY = (pCaptcha.ScaleHeight / 2) - RandomNumber(150, 170)
    pCaptcha.ForeColor = RGB(RandomNumber(60, 255), RandomNumber(60, 255), RandomNumber(60, 255))
    pCaptcha.Print Chr(CaptchaCode(1))
    pCaptcha.Line (RandomNumber(pCaptcha.ScaleWidth, 30), RandomNumber(1, pCaptcha.ScaleHeight))-(RandomNumber(pCaptcha.ScaleHeight - 30, pCaptcha.ScaleHeight), RandomNumber(1, pCaptcha.ScaleHeight)), RGB(RandomNumber(60, 255), RandomNumber(60, 255), RandomNumber(40, 255))
    pCaptcha.Line (RandomNumber(1, 30), RandomNumber(1, pCaptcha.ScaleHeight))-(RandomNumber(pCaptcha.ScaleWidth - 30, pCaptcha.ScaleHeight), RandomNumber(1, pCaptcha.ScaleHeight)), RGB(RandomNumber(60, 255), RandomNumber(60, 255), RandomNumber(40, 255))
    pCaptcha.Line (RandomNumber(pCaptcha.ScaleWidth, 30), RandomNumber(1, pCaptcha.ScaleHeight))-(RandomNumber(pCaptcha.ScaleHeight, pCaptcha.ScaleHeight), RandomNumber(1, pCaptcha.ScaleHeight)), RGB(RandomNumber(60, 255), RandomNumber(60, 255), RandomNumber(40, 255))

End Sub
