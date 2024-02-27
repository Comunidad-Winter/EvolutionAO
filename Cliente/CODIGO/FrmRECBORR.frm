VERSION 5.00
Begin VB.Form FrmRECBORR 
   BorderStyle     =   0  'None
   Caption         =   "Panel de Usuario"
   ClientHeight    =   3000
   ClientLeft      =   0
   ClientTop       =   45
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox DATO3 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   285
      Left            =   2470
      TabIndex        =   2
      Top             =   2420
      Width           =   1050
   End
   Begin VB.TextBox DATO1 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   285
      Left            =   1100
      TabIndex        =   1
      Top             =   1035
      Width           =   3890
   End
   Begin VB.TextBox DATO2 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   285
      Left            =   1070
      TabIndex        =   0
      Top             =   1770
      Width           =   3890
   End
   Begin VB.Image Image 
      Height          =   255
      Left            =   480
      MouseIcon       =   "FrmRECBORR.frx":0000
      MousePointer    =   99  'Custom
      Top             =   2520
      Width           =   1560
   End
   Begin VB.Image imgACCION 
      Height          =   255
      Left            =   3840
      MouseIcon       =   "FrmRECBORR.frx":0152
      MousePointer    =   99  'Custom
      Top             =   2520
      Width           =   1695
   End
End
Attribute VB_Name = "FrmRECBORR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
    DATO1.SetFocus
End Sub

Private Sub Form_Load()

    Select Case EstadoLogin
            'Case E_MODO.BorrarCuenta
            'Me.Caption = "Borrar Cuenta"
            'Me.Picture = LoadPicture(DirGraficos & "VentanaBorrar.jpg")
        Case E_MODO.RecuperarCuenta
            Me.Caption = "Recuperar Cuenta"
            Me.Picture = LoadPicture(DirGraficos & "VentanaRecuperar.jpg")
    End Select

End Sub

Private Sub Image_Click()
    Unload Me
End Sub

Private Sub imgACCION_Click()

    AccountName = DATO1
    AccountPin = DATO3

    If Len(AccountName) < 2 Then
        Call MsgBox("Ingrese un nombre.")
        Exit Sub
    End If
    
    If Len(AccountPin) < 2 Then
        Call MsgBox("Ingrese un pin.")
        Exit Sub
    End If

    Select Case EstadoLogin

            'Case E_MODO.BorrarCuenta

            'UserPassword = DATO2

            'If Len(UserPassword) < 2 Then
            'MsgBox ("Ingrese un password.")
            'Exit Sub
            'End If

        Case E_MODO.RecuperarCuenta

            AccountEmail = DATO2

            If Len(AccountEmail) < 3 Then
                Call MsgBox("Ingrese un E-Mail.")
                Exit Sub
            End If

            If Not CheckMailString(AccountEmail) Then
                Call MsgBox("Ingrese un email valido.")
                Exit Sub
            End If

        Case Else

            Call MsgBox("Ocurrio un error En el proceso. Reintentelo.")
            Unload Me
            Exit Sub

    End Select

    Call Login
    Unload Me

End Sub
