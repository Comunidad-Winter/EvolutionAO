VERSION 5.00
Begin VB.Form frmRetos1vs1 
   BackColor       =   &H8000000B&
   BorderStyle     =   0  'None
   ClientHeight    =   4695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton Option2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   1385
      MaskColor       =   &H0000FFFF&
      MouseIcon       =   "frmRetos1vs1.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   1380
      Width           =   210
   End
   Begin VB.TextBox tOponente 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   250
      Index           =   1
      Left            =   2650
      MaxLength       =   30
      TabIndex        =   10
      Top             =   2670
      Width           =   1440
   End
   Begin VB.TextBox tOponenteTeam 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   250
      Left            =   535
      MaxLength       =   30
      TabIndex        =   9
      Top             =   2670
      Width           =   1440
   End
   Begin VB.TextBox tTeam 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   250
      Left            =   1300
      MaxLength       =   30
      TabIndex        =   8
      Top             =   2030
      Width           =   2090
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1390
      MouseIcon       =   "frmRetos1vs1.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   280
      Value           =   -1  'True
      Width           =   210
   End
   Begin VB.TextBox txtPoints 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H0080C0FF&
      Height          =   200
      Left            =   3070
      MaxLength       =   4
      TabIndex        =   5
      Text            =   "0"
      Top             =   3765
      Width           =   950
   End
   Begin VB.TextBox tOro 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H0080C0FF&
      Height          =   200
      Left            =   970
      MaxLength       =   7
      TabIndex        =   4
      Text            =   "20000"
      Top             =   3765
      Width           =   950
   End
   Begin VB.CheckBox chkDrop 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      Caption         =   "Drop"
      ForeColor       =   &H80000008&
      Height          =   215
      Left            =   490
      MaskColor       =   &H00000000&
      TabIndex        =   0
      Top             =   3120
      Width           =   215
   End
   Begin VB.TextBox tOponente 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   250
      Index           =   0
      Left            =   1300
      MaxLength       =   30
      TabIndex        =   3
      Top             =   890
      Width           =   2090
   End
   Begin VB.Image imgPorItem 
      Height          =   210
      Left            =   480
      MouseIcon       =   "frmRetos1vs1.frx":02A4
      MousePointer    =   99  'Custom
      Top             =   3120
      Width           =   255
   End
   Begin VB.Label cmdUnload 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   4320
      MouseIcon       =   "frmRetos1vs1.frx":03F6
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   0
      Width           =   375
   End
   Begin VB.Label cmdSend 
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   1440
      MouseIcon       =   "frmRetos1vs1.frx":0548
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   4200
      Width           =   1815
   End
End
Attribute VB_Name = "frmRetos1vs1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ClassMove As New clsFormMovementManager

Private Sub cmdSend_Click()

    If (Val(tOro.Text) < 1) Then
        Call ShowConsoleMsg("Cantidad de oro inválida.")
        Exit Sub
    End If

    If Not IsNumeric(tOro.Text) Then
        Call ShowConsoleMsg("Oro inválido.")
        Exit Sub
    End If

    If (Val(txtPoints.Text) < 0) Then
        Call ShowConsoleMsg("Cantidad de Canjes inválida.")
        Exit Sub
    End If

    If Option1.Value <> 0 Then

        If (tOponente(0).Text = vbNullString) Then
            Call ShowConsoleMsg("Nombre del oponente inválido.")
            Exit Sub
        End If

        Call WriteOtherSendReto(tOponente(0).Text, Val(tOro.Text), chkDrop.Value, Val(txtPoints.Text))

    Else

        If (tTeam.Text = vbNullString) Or (IsNumeric(tTeam.Text)) Then
            Call ShowConsoleMsg("Nombre del compañero inválido.")
            Exit Sub
        End If

        If (tOponente(1).Text = vbNullString) Or (IsNumeric(tOponente(1).Text)) Then
            Call ShowConsoleMsg("Nombre del oponente inválido.")
            Exit Sub
        End If

        If (tOponenteTeam.Text = vbNullString) Or (IsNumeric(tOponenteTeam.Text)) Then
            Call ShowConsoleMsg("Nombre del compañero del oponente inválido.")
            Exit Sub
        End If

        Call WriteSendReto(tTeam.Text, tOponente(1).Text, tOponenteTeam.Text, chkDrop.Value, Val(tOro.Text), Val(txtPoints.Text), 0)    ' CBool(chkRespawn.Value))

    End If

    Unload Me
    frmMain.SetFocus

End Sub

Private Sub cmdUnload_Click()

    Unload Me
    frmMain.SetFocus

End Sub

Private Sub Form_Load()

    On Error Resume Next

    Set ClassMove = New clsFormMovementManager
    Call ClassMove.Initialize(Me)
    Me.Picture = LoadPicture(App.path & "\Graficos\VentanaRetos.jpg")

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set ClassMove = Nothing

End Sub

Private Sub imgPorItem_Click()

    If chkDrop.Value = 0 Then
        chkDrop.Value = 1
    Else
        chkDrop.Value = 0
    End If

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Option1_Click()

    If Option2.Value Then
        Option2.Value = 0
    End If

    tOponente(1).Enabled = False
    tOponente(0).Enabled = True
    tOponenteTeam.Enabled = False
    tTeam.Enabled = False

End Sub

Private Sub Option2_Click()

    If Option1.Value Then
        Option1.Value = 0
    End If

    tOponente(0).Enabled = False
    tOponente(1).Enabled = True
    tOponenteTeam.Enabled = True
    tTeam.Enabled = True

End Sub

Private Sub tOponente_Change(Index As Integer)
    tOponente(Index).Text = LTrim$(tOponente(Index).Text)
End Sub

Private Sub tOponenteTeam_Change()
    tOponenteTeam.Text = LTrim$(tOponenteTeam.Text)
End Sub

Private Sub tTeam_Change()
    tTeam.Text = LTrim$(tTeam.Text)
End Sub
