VERSION 5.00
Begin VB.Form frmRetos2vs2 
   BorderStyle     =   0  'None
   ClientHeight    =   5100
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox tPoints 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000005&
      Height          =   340
      Left            =   2280
      MaxLength       =   3
      TabIndex        =   8
      Text            =   "0"
      Top             =   3480
      Width           =   1155
   End
   Begin VB.TextBox tOro 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000005&
      Height          =   340
      Left            =   360
      MaxLength       =   7
      TabIndex        =   7
      Text            =   "123123"
      Top             =   3480
      Width           =   1035
   End
   Begin VB.TextBox tOponente 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000005&
      Height          =   340
      Left            =   360
      MaxLength       =   30
      TabIndex        =   6
      Top             =   2670
      Width           =   3075
   End
   Begin VB.TextBox tOponenteTeam 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000005&
      Height          =   340
      Left            =   360
      MaxLength       =   30
      TabIndex        =   5
      Top             =   1920
      Width           =   3075
   End
   Begin VB.CheckBox chkRespawn 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      Caption         =   "Drop"
      ForeColor       =   &H80000008&
      Height          =   215
      Left            =   2040
      MaskColor       =   &H00000000&
      TabIndex        =   1
      Top             =   4070
      Width           =   215
   End
   Begin VB.CheckBox chkDrop 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      Caption         =   "Drop"
      ForeColor       =   &H80000008&
      Height          =   215
      Left            =   260
      MaskColor       =   &H00000000&
      TabIndex        =   0
      Top             =   4070
      Width           =   215
   End
   Begin VB.TextBox tTeam 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000005&
      Height          =   340
      Left            =   360
      MaxLength       =   30
      TabIndex        =   4
      Top             =   1135
      Width           =   3075
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   2280
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   600
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label cmdSend 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   4560
      Width           =   1575
   End
End
Attribute VB_Name = "frmRetos2vs2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private classMove As New clsFormMovementManager

Private Sub Form_Load()

    Set classMove = New clsFormMovementManager

    Call classMove.Initialize(Me)

    Me.Picture = LoadPicture(App.path & "\Graficos\Retos2vs2.jpg")

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set classMove = Nothing

End Sub

Private Sub Image1_Click()

    If chkDrop.Value = 0 Then
        chkDrop.Value = 1
    Else
        chkDrop.Value = 0
    End If

End Sub

Private Sub Image2_Click()

    If chkRespawn.Value = 0 Then
        chkRespawn.Value = 1
    Else
        chkRespawn.Value = 0
    End If

End Sub

Private Sub Label1_Click()

    Unload Me

    frmMain.SetFocus

End Sub

Private Sub tOponente_Change()
    tOponente.Text = LTrim$(tOponente.Text)
End Sub

Private Sub tOponenteTeam_Change()
    tOponenteTeam.Text = LTrim$(tOponenteTeam.Text)
End Sub

Private Sub tTeam_Change()
    tTeam.Text = LTrim$(tTeam.Text)
End Sub
