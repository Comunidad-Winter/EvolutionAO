VERSION 5.00
Begin VB.Form frmCrearCuenta 
   BackColor       =   &H80000000&
   BorderStyle     =   0  'None
   Caption         =   "Crear Cuenta"
   ClientHeight    =   3000
   ClientLeft      =   5115
   ClientTop       =   4125
   ClientWidth     =   6000
   Icon            =   "frmCrearCuenta.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3000
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCuentaPin 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   250
      Left            =   2540
      MaxLength       =   20
      MousePointer    =   3  'I-Beam
      TabIndex        =   4
      Top             =   2480
      Width           =   900
   End
   Begin VB.TextBox txtCuentaNombre 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   250
      Left            =   772
      MaxLength       =   20
      MousePointer    =   3  'I-Beam
      TabIndex        =   0
      Top             =   787
      Width           =   2100
   End
   Begin VB.TextBox txtCuentaRepite 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   250
      IMEMode         =   3  'DISABLE
      Left            =   3218
      MaxLength       =   50
      MousePointer    =   3  'I-Beam
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1890
      Width           =   2100
   End
   Begin VB.TextBox txtCuentaPassword 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   250
      IMEMode         =   3  'DISABLE
      Left            =   772
      MaxLength       =   50
      MousePointer    =   3  'I-Beam
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1890
      Width           =   2100
   End
   Begin VB.TextBox txtCuentaEmail 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   250
      Left            =   3218
      MaxLength       =   30
      MousePointer    =   3  'I-Beam
      TabIndex        =   1
      Top             =   780
      Width           =   2100
   End
   Begin VB.Image ImgSalir 
      Height          =   375
      Left            =   720
      MouseIcon       =   "frmCrearCuenta.frx":000C
      MousePointer    =   99  'Custom
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Image imgCrearCuenta 
      Height          =   375
      Left            =   3720
      MouseIcon       =   "frmCrearCuenta.frx":015E
      MousePointer    =   99  'Custom
      Top             =   2400
      Width           =   1575
   End
End
Attribute VB_Name = "frmCrearCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public LastButtonPressed As clsGraphicalButton

Private Sub Form_Load()

    On Error Resume Next

    Me.Picture = LoadPicture(DirGraficos & "VentanaCrearCuenta.jpg")

    txtCuentaNombre.Text = vbNullString
    txtCuentaPassword.Text = vbNullString
    txtCuentaEmail.Text = vbNullString
    txtCuentaRepite.Text = vbNullString

End Sub

Private Sub ImgCrearCuenta_Click()

    If Not IsFormValid Then Exit Sub

    AccountName = txtCuentaNombre.Text
    AccountPassword = txtCuentaPassword.Text
    AccountEmail = txtCuentaEmail.Text
    AccountPin = txtCuentaPin.Text

    Call Login

End Sub

Private Sub imgSalir_Click()
    Unload Me
End Sub

Private Function IsFormValid() As Boolean

    If LenB(txtCuentaNombre.Text) = 0 Then
        Call MsgBox("Ingrese un nombre.")
        Exit Function
    End If
    
    If Len(txtCuentaNombre.Text) > 20 Then
        Call MsgBox("El nombre debe tener máximo 10 carácteres.")
        Exit Function
    End If
    
    If Len(txtCuentaPassword.Text) < 5 Then
        Call MsgBox("La contraseña debe tener almenos 5 carácteres.")
        Exit Function
    End If

    If LenB(txtCuentaEmail.Text) = 0 Then
        Call MsgBox("Ingrese un email.")
        Exit Function
    End If
    
    If Len(txtCuentaEmail.Text) > 30 Then
        Call MsgBox("El email debe tener menos de 30 carácteres.")
        Exit Function
    End If

    If Len(txtCuentaPin.Text) < 5 Then
        Call MsgBox("El pin debe tener almenos 5 carácteres.")
        Exit Function
    End If

    If Len(txtCuentaPin.Text) > 20 Then
        Call MsgBox("El pin debe tener menos de 20 carácteres.")
        Exit Function
    End If
    
    If Not CheckMailString(txtCuentaEmail.Text) Then
        Call MsgBox("Ingrese un email valido.")
        Exit Function
    End If
    
    If txtCuentaPassword.Text <> txtCuentaRepite.Text Then
        Call MsgBox("Las contraseñas no coinciden.")
        Exit Function
    End If

    Dim LoopC As Long
    Dim CharAscii As Integer

    For LoopC = 1 To Len(txtCuentaNombre.Text)
        CharAscii = Asc(mid$(txtCuentaNombre.Text, LoopC, 1))
        If Not LegalCharacter(CharAscii) Then
            Call MsgBox("Nombre inválido. El caractér " & Chr$(CharAscii) & " no está permitido.")
            Exit Function
        End If
    Next LoopC

    For LoopC = 1 To Len(txtCuentaPassword.Text)
        CharAscii = Asc(mid$(txtCuentaPassword.Text, LoopC, 1))
        If Not LegalCharacter(CharAscii) Then
            Call MsgBox("Password inválido. El caractér " & Chr$(CharAscii) & " no está permitido.")
            Exit Function
        End If
    Next LoopC

    IsFormValid = True

End Function

Private Sub txtCuentaNombre_Change()
    txtCuentaNombre.Text = LTrim(txtCuentaNombre.Text)
End Sub

Private Sub txtCuentaNombre_KeyPress(KeyAscii As Integer)
    If (KeyAscii >= 48) And (KeyAscii <= 57) Then
        KeyAscii = 0
    End If
End Sub
