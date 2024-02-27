VERSION 5.00
Begin VB.Form frmPanelAccount 
   BackColor       =   &H80000011&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9000
   ClientLeft      =   -120
   ClientTop       =   -1155
   ClientWidth     =   12000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H0080C0FF&
   Icon            =   "frmPanelAccount.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   770.878
   ScaleMode       =   0  'User
   ScaleWidth      =   808.081
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000005&
      Height          =   285
      Left            =   5160
      MaxLength       =   20
      TabIndex        =   24
      Top             =   480
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Timer AnimacionTimer 
      Interval        =   100
      Left            =   360
      Top             =   360
   End
   Begin VB.PictureBox picChar 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1215
      Index           =   7
      Left            =   9815
      MouseIcon       =   "frmPanelAccount.frx":000C
      MousePointer    =   99  'Custom
      ScaleHeight     =   81
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   73
      TabIndex        =   22
      Top             =   3841
      Width           =   1095
   End
   Begin VB.PictureBox picChar 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1215
      Index           =   6
      Left            =   6935
      MouseIcon       =   "frmPanelAccount.frx":015E
      MousePointer    =   99  'Custom
      ScaleHeight     =   81
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   73
      TabIndex        =   21
      Top             =   3841
      Width           =   1095
   End
   Begin VB.PictureBox picChar 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1215
      Index           =   5
      Left            =   4002
      MouseIcon       =   "frmPanelAccount.frx":02B0
      MousePointer    =   99  'Custom
      ScaleHeight     =   81
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   73
      TabIndex        =   20
      Top             =   3841
      Width           =   1095
   End
   Begin VB.PictureBox picChar 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1215
      Index           =   4
      Left            =   1114
      MouseIcon       =   "frmPanelAccount.frx":0402
      MousePointer    =   99  'Custom
      ScaleHeight     =   81
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   73
      TabIndex        =   19
      Top             =   3841
      Width           =   1095
   End
   Begin VB.PictureBox picChar 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1215
      Index           =   3
      Left            =   9815
      MouseIcon       =   "frmPanelAccount.frx":0554
      MousePointer    =   99  'Custom
      ScaleHeight     =   81
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   73
      TabIndex        =   18
      Top             =   1270
      Width           =   1095
   End
   Begin VB.PictureBox picChar 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1215
      Index           =   2
      Left            =   6935
      MouseIcon       =   "frmPanelAccount.frx":06A6
      MousePointer    =   99  'Custom
      ScaleHeight     =   81
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   73
      TabIndex        =   17
      Top             =   1270
      Width           =   1095
   End
   Begin VB.PictureBox picChar 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1215
      Index           =   1
      Left            =   4002
      MouseIcon       =   "frmPanelAccount.frx":07F8
      MousePointer    =   99  'Custom
      ScaleHeight     =   81
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   73
      TabIndex        =   16
      Top             =   1270
      Width           =   1095
   End
   Begin VB.PictureBox picChar 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1215
      Index           =   0
      Left            =   1114
      MouseIcon       =   "frmPanelAccount.frx":094A
      MousePointer    =   99  'Custom
      ScaleHeight     =   81
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   73
      TabIndex        =   15
      Top             =   1270
      Width           =   1095
   End
   Begin VB.Label lblCrearPersonaje 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   4680
      MouseIcon       =   "frmPanelAccount.frx":0A9C
      MousePointer    =   99  'Custom
      TabIndex        =   23
      Top             =   6360
      Width           =   2775
   End
   Begin VB.Image imgCambiarPassword 
      Height          =   495
      Left            =   4560
      MouseIcon       =   "frmPanelAccount.frx":0BEE
      MousePointer    =   99  'Custom
      Top             =   7080
      Width           =   2895
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje 1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   195
      Index           =   1
      Left            =   1129
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   2895
      Width           =   1035
   End
   Begin VB.Image imgSalir 
      Height          =   495
      Left            =   4560
      MouseIcon       =   "frmPanelAccount.frx":0D40
      MousePointer    =   99  'Custom
      Top             =   7800
      Width           =   2895
   End
   Begin VB.Label lblCharData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   5
      Left            =   8235
      TabIndex        =   14
      Top             =   7650
      Width           =   45
   End
   Begin VB.Label lblCharData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   4
      Left            =   8235
      TabIndex        =   13
      Top             =   7125
      Width           =   45
   End
   Begin VB.Label lblCharData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   3
      Left            =   8235
      TabIndex        =   12
      Top             =   6600
      Width           =   45
   End
   Begin VB.Label lblCharData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   1
      Left            =   1215
      TabIndex        =   11
      Top             =   7650
      Width           =   45
   End
   Begin VB.Label lblCharData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   2
      Left            =   1215
      TabIndex        =   10
      Top             =   7125
      Width           =   45
   End
   Begin VB.Label lblCharData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   0
      Left            =   1215
      TabIndex        =   9
      Top             =   6600
      Width           =   45
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje 8"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   195
      Index           =   8
      Left            =   9860
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   5464
      Width           =   1035
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje 7"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   195
      Index           =   7
      Left            =   6979
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   5464
      Width           =   1035
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje 6"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   195
      Index           =   6
      Left            =   4039
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   5464
      Width           =   1035
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje 5"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   195
      Index           =   5
      Left            =   1129
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   5464
      Width           =   1035
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje 4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   195
      Index           =   4
      Left            =   9860
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   2895
      Width           =   1035
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje 3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   195
      Index           =   3
      Left            =   6979
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   2895
      Width           =   1035
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Personaje 2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   195
      Index           =   2
      Left            =   4039
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   2895
      Width           =   1035
   End
   Begin VB.Label lblAccData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cuenta"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   375
      Index           =   0
      Left            =   3891
      TabIndex        =   5
      Top             =   5880
      Visible         =   0   'False
      Width           =   4185
   End
   Begin VB.Menu mBorrar 
      Caption         =   "Borrar Personaje"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu bBorrar 
         Caption         =   "¿Borrar personaje?"
      End
   End
End
Attribute VB_Name = "frmPanelAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private LastNameClicked As Byte

Private LastSelected As Byte
Private Frame_Counter As Byte
Private AUX_Rect As RECT

Private Sub AnimacionTimer_Timer()

    If Seleccionado <> 0 Then
        Frame_Counter = Frame_Counter + 1
        If Frame_Counter > 254 Then Frame_Counter = 1
        Call DrawPJ(Seleccionado)
    End If

End Sub

Private Sub Form_Click()
    
    If LastNameClicked > 0 Then
        If UCase$(txtName.Text) = UCase$(lblAccData(LastNameClicked).Caption) Then
            cPJ(LastNameClicked).Nombre = txtName.Text
            lblAccData(LastNameClicked).Caption = cPJ(LastNameClicked).Nombre
            txtName.Text = vbNullString
            txtName.Visible = False
            LastNameClicked = 0
        End If
    End If
    
End Sub

Private Sub Form_Load()

    Unload frmConnect
    Me.Picture = LoadPicture(DirGraficos & "VentanaCuenta.jpg")

    Dim i As Long

    Me.Icon = frmMain.Icon

    For i = 1 To 8
        Me.lblAccData(i).Caption = vbNullString
    Next i

    Me.lblAccData(0).Caption = AccountName

    With AUX_Rect
        .Left = 0
        .Top = 0
        .Right = 50
        .Bottom = 65
    End With

End Sub

Private Sub Image5_Click()

    If LenB(lblAccData(Seleccionado + 1).Caption) <> 0 Then
        UserName = lblAccData(Seleccionado + 1).Caption
        Call WriteLoginExistingChar
    End If

End Sub

Private Sub lblName_Click(Index As Integer)
    Seleccionado = Index
End Sub

Private Sub imgCrearPersonaje_Click()

    If NumberOfCharacters > 9 Then
        Call MsgBox("No puedes crear mas de 10 personajes.")
        Exit Sub
    End If

    Dim LoopC As Long

    For LoopC = 1 To 10
        If LenB(lblAccData(LoopC).Caption) = 0 Then
            frmCrearPersonaje.Show
            Exit Sub
        End If
    Next LoopC

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Seleccionado <> 0 Then
        Seleccionado = 0

        If LastSelected <> Seleccionado Then
            Frame_Counter = 1
            Call DrawPJ(LastSelected)
            lblAccData(LastSelected).FontUnderline = False
        End If
    End If

End Sub

Private Sub imgCambiarPassword_Click()
    frmNewPasswordAccount.Show
End Sub

Private Sub imgSalir_Click()
    Unload Me
    frmConnect.Show
End Sub

Private Sub lblAccData_Click(Index As Integer)

    If LenB(lblAccData(Index).Caption) <> 0 Then

        If Index > 0 Then
            LastNameClicked = Index
            txtName.MaxLength = Len(lblAccData(Index).Caption)
            txtName.Text = lblAccData(Index).Caption
            txtName.SelStart = txtName.MaxLength

            txtName.Top = lblAccData(Index).Top + 20
            txtName.Left = lblAccData(Index).Left - 40

            txtName.Visible = True
            txtName.SetFocus
        End If
    End If

End Sub

Private Sub lblCrearPersonaje_Click()
    frmCrearPersonaje.Show
End Sub

Private Sub picChar_Click(Index As Integer)

    Seleccionado = Index + 1

    If Seleccionado > NumberOfCharacters Then
        Seleccionado = 0
        Exit Sub
    End If

    With cPJ(Seleccionado)
        If LenB(.Nombre) <> 0 Then
            lblCharData(0) = "Nombre: " & .Nombre
            lblCharData(1) = "Clase: " & ListaClases(.Class)
            lblCharData(2) = "Raza: " & ListaRazas(.Race)
            lblCharData(3) = "Nivel: " & .Level
            lblCharData(4) = "Oro: " & .Gold
            lblCharData(5) = "Mapa: " & .Map
        Else
            lblCharData(0) = vbNullString
            lblCharData(1) = vbNullString
            lblCharData(2) = vbNullString
            lblCharData(3) = vbNullString
            lblCharData(4) = vbNullString
            lblCharData(5) = vbNullString
        End If
    End With

End Sub

Private Sub picChar_DblClick(Index As Integer)

    Seleccionado = Index + 1

    If Seleccionado > NumberOfCharacters Then
        Seleccionado = 0
        frmCrearPersonaje.Show
        Exit Sub
    End If

    If LenB(cPJ(Seleccionado).Nombre) <> 0 Then
        UserName = cPJ(Seleccionado).Nombre
        Call WriteLoginExistingChar
    Else
        Unload Me
        frmCrearPersonaje.Show
    End If

End Sub

Public Sub DrawPJ(ByVal Index As Byte)

    If Index < 1 Then Exit Sub
    If Index > NumberOfCharacters Then Exit Sub

    picChar(Index - 1).Cls

    If LenB(cPJ(Index).Nombre) <> 0 Then

        Dim Head_OffSet As Byte
        Call Clear_Rect_Account(AUX_Rect)

        With cPJ(Index)

            If .Body <> 0 Then

                If .Race <> eRaza.Gnomo And .Race <> eRaza.Enano Then
                    Head_OffSet = 12
                Else
                    Head_OffSet = 22
                End If

                If .Body <> 8 Then
                    Call DrawGrh(BodyData(.Body).Walk(3), 25, 20, Index, 1)
                Else
                    Call DrawGrh(BodyData(.Body).Walk(3), 20, 28, Index, 1)
                End If

                If .Head <> 0 Then
                    If .Head <> 500 Then
                        Call DrawGrh(HeadData(.Head).Head(3), 29, Head_OffSet, Index)
                    Else
                        Call DrawGrh(HeadData(.Head).Head(3), 26, 17, Index)
                    End If
                End If

                If .Helmet <> 0 Then
                    If .Helmet <> 3 And .Helmet <> 4 And .Helmet <> 7 Then
                        If .Helmet <> 21 Then
                            Call DrawGrh(CascoAnimData(.Helmet).Head(3), 29, Head_OffSet, Index)
                        Else
                            Call DrawGrh(CascoAnimData(.Helmet).Head(3), 28, Head_OffSet - 4, Index)
                        End If
                    Else
                        Call DrawGrh(CascoAnimData(.Helmet).Head(3), 29, 0, Index)
                    End If
                End If

                If .Weapon <> 0 Then
                    Call DrawGrh(WeaponAnimData(.Weapon).WeaponWalk(3), 25, 20, Index, 1)
                End If

                If .Shield <> 0 Then
                    Call DrawGrh(ShieldAnimData(.Shield).ShieldWalk(3), 25, 20, Index, 1)
                End If

            End If

        End With

    End If

    picChar(Index - 1).Refresh

End Sub

Private Sub DrawGrh(Grh As Grh, ByVal X As Byte, ByVal Y As Byte, ByVal Slot As Byte, Optional Animacion As Byte = 0)

    If Grh.GrhIndex < 1 Then Exit Sub

    Dim GrhIndex As Integer

    If Seleccionado <> Slot Then
        GrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    Else
        If Animacion <> 1 Then
            GrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
        Else
            If Frame_Counter > GrhData(Grh.GrhIndex).NumFrames Then Frame_Counter = 1
            GrhIndex = GrhData(Grh.GrhIndex).Frames(Frame_Counter)
        End If
    End If

    If GrhIndex < 1 Then Exit Sub

    Dim N_Rect As RECT

    With N_Rect
        .Left = GrhData(GrhIndex).sX
        .Top = GrhData(GrhIndex).sY
        .Right = .Left + GrhData(GrhIndex).pixelWidth
        .Bottom = .Top + GrhData(GrhIndex).pixelHeight
    End With

    Call Draw_Account_Grh(Slot, X, Y, GrhData(GrhIndex).FileNum, N_Rect, AUX_Rect)

End Sub

Private Sub picChar_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbRightButton Then
        Seleccionado = Index + 1

        If Seleccionado > NumberOfCharacters Then
            Seleccionado = 0
            Exit Sub
        End If

        Call PopUpMenu(mBorrar)
    End If

End Sub

Private Sub bBorrar_Click()

    If Seleccionado <> 0 Then

        UserName = cPJ(Seleccionado).Nombre

        If LenB(UserName) <> 0 Then
            Dim tStr As String
            tStr = InputBox("Escriba el email de la cuenta.", "Borrar personaje")

            If Not CheckMailString(tStr) Then
                Call MsgBox("Ingrese un email valido.")
                Exit Sub
            End If
            
            If LenB(tStr) < 1 Then Exit Sub
            AccountEmail = tStr
            
            tStr = InputBox("Escriba el pin de la cuenta.", "Borrar personaje")
            
            If LenB(tStr) < 1 Or Len(tStr) > 20 Then
                Call MsgBox("Ingrese un pin valido.")
                Exit Sub
            End If
            
            AccountPin = tStr

            If MsgBox(" ¿Estás realmente seguro que deseas eliminar el personaje " & UserName & "?", vbYesNo, "Atencion!") = vbYes Then
                Call WriteDeleteCharAccount
            End If
        End If

    End If

End Sub

Private Sub picChar_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Seleccionado = Index + 1

    If Seleccionado > NumberOfCharacters Then
        Seleccionado = 0
        Exit Sub
    End If

    If LastSelected <> Seleccionado Then
        Frame_Counter = 1
        Call DrawPJ(LastSelected)
        lblAccData(LastSelected).FontUnderline = False
    End If

    lblAccData(Seleccionado).FontUnderline = True
    LastSelected = Seleccionado

End Sub

Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)

    If LastNameClicked < 1 Then Exit Sub
    If KeyCode = 16 Then Exit Sub

    If KeyCode <> 8 Then
        If KeyCode <> 13 Then
            If Len(txtName.Text) >= txtName.MaxLength Then Exit Sub
        Else
            If KeyCode = 13 Then
                If UCase$(txtName.Text) = UCase$(lblAccData(LastNameClicked).Caption) Then
                    cPJ(LastNameClicked).Nombre = txtName.Text
                    lblAccData(LastNameClicked).Caption = cPJ(LastNameClicked).Nombre
                    txtName.Text = vbNullString
                    txtName.Visible = False
                    LastNameClicked = 0
                End If
                Exit Sub
            End If
        End If
    End If

End Sub
