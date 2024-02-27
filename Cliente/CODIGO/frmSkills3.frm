VERSION 5.00
Begin VB.Form frmSkills3 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9000
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   450
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image imgNavegacion 
      Height          =   300
      Left            =   4605
      Top             =   4110
      Width           =   1440
   End
   Begin VB.Image imgCombateSinArmas 
      Height          =   345
      Left            =   4605
      Top             =   3735
      Width           =   2100
   End
   Begin VB.Image imgCombateDistancia 
      Height          =   345
      Left            =   4605
      Top             =   3345
      Width           =   2280
   End
   Begin VB.Image imgDomar 
      Height          =   345
      Left            =   4605
      Top             =   2970
      Width           =   1845
   End
   Begin VB.Image imgLiderazgo 
      Height          =   330
      Left            =   4605
      Top             =   2580
      Width           =   1215
   End
   Begin VB.Image imgHerreria 
      Height          =   345
      Left            =   4605
      Top             =   2205
      Width           =   1065
   End
   Begin VB.Image imgCarpinteria 
      Height          =   360
      Left            =   4605
      Top             =   1830
      Width           =   1365
   End
   Begin VB.Image imgMineria 
      Height          =   360
      Left            =   4605
      Top             =   1440
      Width           =   1005
   End
   Begin VB.Image imgPesca 
      Height          =   330
      Left            =   4605
      Top             =   1110
      Width           =   780
   End
   Begin VB.Image imgEscudos 
      Height          =   270
      Left            =   4605
      Top             =   720
      Width           =   2340
   End
   Begin VB.Image imgComercio 
      Height          =   330
      Left            =   360
      Top             =   4050
      Width           =   1350
   End
   Begin VB.Image imgTalar 
      Height          =   270
      Left            =   360
      Top             =   3690
      Width           =   885
   End
   Begin VB.Image imgSupervivencia 
      Height          =   330
      Left            =   360
      Top             =   3300
      Width           =   1800
   End
   Begin VB.Image imgOcultarse 
      Height          =   270
      Left            =   375
      Top             =   2970
      Width           =   1275
   End
   Begin VB.Image imgApunialar 
      Height          =   360
      Left            =   375
      Top             =   2550
      Width           =   1230
   End
   Begin VB.Image imgMeditar 
      Height          =   270
      Left            =   375
      Top             =   2175
      Width           =   1065
   End
   Begin VB.Image imgCombateArmas 
      Height          =   315
      Left            =   375
      Top             =   1800
      Width           =   2475
   End
   Begin VB.Image imgEvasion 
      Height          =   330
      Left            =   375
      Top             =   1425
      Width           =   2505
   End
   Begin VB.Image imgRobar 
      Height          =   300
      Left            =   375
      Top             =   1065
      Width           =   870
   End
   Begin VB.Image imgMagia 
      Height          =   330
      Left            =   375
      Top             =   675
      Width           =   870
   End
   Begin VB.Label lblHelp 
      BackStyle       =   0  'Transparent
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
      Height          =   1215
      Left            =   600
      TabIndex        =   21
      Top             =   4740
      Width           =   7815
   End
   Begin VB.Image imgCancelar 
      Height          =   360
      Left            =   360
      MouseIcon       =   "frmSkills3.frx":0000
      MousePointer    =   99  'Custom
      Top             =   6240
      Width           =   1575
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   1
      Left            =   3465
      TabIndex        =   20
      Top             =   750
      Width           =   405
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   2
      Left            =   3465
      TabIndex        =   19
      Top             =   1125
      Width           =   405
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   3
      Left            =   3465
      TabIndex        =   18
      Top             =   1500
      Width           =   405
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   4
      Left            =   3465
      TabIndex        =   17
      Top             =   1875
      Width           =   405
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   5
      Left            =   3465
      TabIndex        =   16
      Top             =   2250
      Width           =   405
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   6
      Left            =   3465
      TabIndex        =   15
      Top             =   2625
      Width           =   405
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   7
      Left            =   3465
      TabIndex        =   14
      Top             =   3000
      Width           =   405
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   8
      Left            =   3465
      TabIndex        =   13
      Top             =   3375
      Width           =   405
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   9
      Left            =   3465
      TabIndex        =   12
      Top             =   3750
      Width           =   405
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Index           =   10
      Left            =   3465
      TabIndex        =   11
      Top             =   4125
      Width           =   405
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   11
      Left            =   7905
      TabIndex        =   10
      Top             =   750
      Width           =   405
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   12
      Left            =   7905
      TabIndex        =   9
      Top             =   1125
      Width           =   405
   End
   Begin VB.Image imgMas1 
      Height          =   300
      Left            =   3900
      MouseIcon       =   "frmSkills3.frx":0152
      MousePointer    =   99  'Custom
      Top             =   720
      Width           =   300
   End
   Begin VB.Image imgMas2 
      Height          =   300
      Left            =   3900
      MouseIcon       =   "frmSkills3.frx":02A4
      MousePointer    =   99  'Custom
      Top             =   1095
      Width           =   300
   End
   Begin VB.Image imgMenos2 
      Height          =   300
      Left            =   3135
      MouseIcon       =   "frmSkills3.frx":03F6
      MousePointer    =   99  'Custom
      Top             =   1095
      Width           =   300
   End
   Begin VB.Image imgMas3 
      Height          =   300
      Left            =   3900
      MouseIcon       =   "frmSkills3.frx":0548
      MousePointer    =   99  'Custom
      Top             =   1455
      Width           =   300
   End
   Begin VB.Image imgMenos3 
      Height          =   300
      Left            =   3135
      MouseIcon       =   "frmSkills3.frx":069A
      MousePointer    =   99  'Custom
      Top             =   1455
      Width           =   300
   End
   Begin VB.Image imgMas4 
      Height          =   300
      Left            =   3900
      MouseIcon       =   "frmSkills3.frx":07EC
      MousePointer    =   99  'Custom
      Top             =   1830
      Width           =   300
   End
   Begin VB.Image imgMenos4 
      Height          =   300
      Left            =   3135
      MouseIcon       =   "frmSkills3.frx":093E
      MousePointer    =   99  'Custom
      Top             =   1830
      Width           =   300
   End
   Begin VB.Image imgMas5 
      Height          =   300
      Left            =   3900
      MouseIcon       =   "frmSkills3.frx":0A90
      MousePointer    =   99  'Custom
      Top             =   2205
      Width           =   300
   End
   Begin VB.Image imgMenos5 
      Height          =   300
      Left            =   3135
      MouseIcon       =   "frmSkills3.frx":0BE2
      MousePointer    =   99  'Custom
      Top             =   2205
      Width           =   300
   End
   Begin VB.Image imgMas6 
      Height          =   300
      Left            =   3900
      MouseIcon       =   "frmSkills3.frx":0D34
      MousePointer    =   99  'Custom
      Top             =   2550
      Width           =   300
   End
   Begin VB.Image imgMenos6 
      Height          =   300
      Left            =   3135
      MouseIcon       =   "frmSkills3.frx":0E86
      MousePointer    =   99  'Custom
      Top             =   2550
      Width           =   300
   End
   Begin VB.Image imgMas7 
      Height          =   300
      Left            =   3900
      MouseIcon       =   "frmSkills3.frx":0FD8
      MousePointer    =   99  'Custom
      Top             =   2970
      Width           =   300
   End
   Begin VB.Image imgMenos7 
      Height          =   300
      Left            =   3135
      MouseIcon       =   "frmSkills3.frx":112A
      MousePointer    =   99  'Custom
      Top             =   2970
      Width           =   300
   End
   Begin VB.Image imgMas8 
      Height          =   300
      Left            =   3900
      MouseIcon       =   "frmSkills3.frx":127C
      MousePointer    =   99  'Custom
      Top             =   3330
      Width           =   300
   End
   Begin VB.Image imgMenos8 
      Height          =   300
      Left            =   3135
      MouseIcon       =   "frmSkills3.frx":13CE
      MousePointer    =   99  'Custom
      Top             =   3330
      Width           =   300
   End
   Begin VB.Image imgMas9 
      Height          =   300
      Left            =   3900
      MouseIcon       =   "frmSkills3.frx":1520
      MousePointer    =   99  'Custom
      Top             =   3720
      Width           =   300
   End
   Begin VB.Image imgMenos9 
      Height          =   300
      Left            =   3135
      MouseIcon       =   "frmSkills3.frx":1672
      MousePointer    =   99  'Custom
      Top             =   3720
      Width           =   300
   End
   Begin VB.Image imgMas10 
      Height          =   300
      Left            =   3900
      MouseIcon       =   "frmSkills3.frx":17C4
      MousePointer    =   99  'Custom
      Top             =   4080
      Width           =   300
   End
   Begin VB.Image imgMenos10 
      Height          =   300
      Left            =   3135
      MouseIcon       =   "frmSkills3.frx":1916
      MousePointer    =   99  'Custom
      Top             =   4080
      Width           =   300
   End
   Begin VB.Image imgMas11 
      Height          =   285
      Left            =   8340
      MouseIcon       =   "frmSkills3.frx":1A68
      MousePointer    =   99  'Custom
      Top             =   720
      Width           =   345
   End
   Begin VB.Image imgMenos11 
      Height          =   285
      Left            =   7500
      MouseIcon       =   "frmSkills3.frx":1BBA
      MousePointer    =   99  'Custom
      Top             =   720
      Width           =   345
   End
   Begin VB.Image imgMas12 
      Height          =   285
      Left            =   8340
      MouseIcon       =   "frmSkills3.frx":1D0C
      MousePointer    =   99  'Custom
      Top             =   1095
      Width           =   345
   End
   Begin VB.Image imgMenos12 
      Height          =   285
      Left            =   7500
      MouseIcon       =   "frmSkills3.frx":1E5E
      MousePointer    =   99  'Custom
      Top             =   1095
      Width           =   345
   End
   Begin VB.Image imgMas13 
      Height          =   285
      Left            =   8340
      MouseIcon       =   "frmSkills3.frx":1FB0
      MousePointer    =   99  'Custom
      Top             =   1455
      Width           =   345
   End
   Begin VB.Image imgMenos13 
      Height          =   285
      Left            =   7500
      MouseIcon       =   "frmSkills3.frx":2102
      MousePointer    =   99  'Custom
      Top             =   1455
      Width           =   345
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   13
      Left            =   7905
      TabIndex        =   8
      Top             =   1500
      Width           =   405
   End
   Begin VB.Image imgMas14 
      Height          =   285
      Left            =   8340
      MouseIcon       =   "frmSkills3.frx":2254
      MousePointer    =   99  'Custom
      Top             =   1830
      Width           =   345
   End
   Begin VB.Image imgMenos14 
      Height          =   285
      Left            =   7500
      MouseIcon       =   "frmSkills3.frx":23A6
      MousePointer    =   99  'Custom
      Top             =   1830
      Width           =   345
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   14
      Left            =   7905
      TabIndex        =   7
      Top             =   1875
      Width           =   405
   End
   Begin VB.Image imgMas15 
      Height          =   285
      Left            =   8340
      MouseIcon       =   "frmSkills3.frx":24F8
      MousePointer    =   99  'Custom
      Top             =   2205
      Width           =   345
   End
   Begin VB.Image imgMenos15 
      Height          =   285
      Left            =   7500
      MouseIcon       =   "frmSkills3.frx":264A
      MousePointer    =   99  'Custom
      Top             =   2205
      Width           =   345
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   15
      Left            =   7905
      TabIndex        =   6
      Top             =   2250
      Width           =   405
   End
   Begin VB.Image imgMas16 
      Height          =   285
      Left            =   8340
      MouseIcon       =   "frmSkills3.frx":279C
      MousePointer    =   99  'Custom
      Top             =   2550
      Width           =   345
   End
   Begin VB.Image imgMenos16 
      Height          =   285
      Left            =   7500
      MouseIcon       =   "frmSkills3.frx":28EE
      MousePointer    =   99  'Custom
      Top             =   2550
      Width           =   345
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   16
      Left            =   7905
      TabIndex        =   5
      Top             =   2625
      Width           =   405
   End
   Begin VB.Image imgMas17 
      Height          =   285
      Left            =   8340
      MouseIcon       =   "frmSkills3.frx":2A40
      MousePointer    =   99  'Custom
      Top             =   2970
      Width           =   345
   End
   Begin VB.Image imgMenos17 
      Height          =   285
      Left            =   7500
      MouseIcon       =   "frmSkills3.frx":2B92
      MousePointer    =   99  'Custom
      Top             =   2970
      Width           =   345
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   17
      Left            =   7905
      TabIndex        =   4
      Top             =   3000
      Width           =   405
   End
   Begin VB.Image imgMas18 
      Height          =   285
      Left            =   8340
      MouseIcon       =   "frmSkills3.frx":2CE4
      MousePointer    =   99  'Custom
      Top             =   3330
      Width           =   345
   End
   Begin VB.Image imgMenos18 
      Height          =   285
      Left            =   7500
      MouseIcon       =   "frmSkills3.frx":2E36
      MousePointer    =   99  'Custom
      Top             =   3330
      Width           =   345
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   18
      Left            =   7905
      TabIndex        =   3
      Top             =   3375
      Width           =   405
   End
   Begin VB.Image imgMenos1 
      Height          =   300
      Left            =   3135
      MouseIcon       =   "frmSkills3.frx":2F88
      MousePointer    =   99  'Custom
      Top             =   720
      Width           =   300
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   19
      Left            =   7905
      TabIndex        =   2
      Top             =   3750
      Width           =   405
   End
   Begin VB.Image imgMas19 
      Height          =   285
      Left            =   8340
      MouseIcon       =   "frmSkills3.frx":30DA
      MousePointer    =   99  'Custom
      Top             =   3720
      Width           =   345
   End
   Begin VB.Image imgMenos19 
      Height          =   285
      Left            =   7500
      MouseIcon       =   "frmSkills3.frx":322C
      MousePointer    =   99  'Custom
      Top             =   3720
      Width           =   345
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   20
      Left            =   7905
      TabIndex        =   1
      Top             =   4125
      Width           =   405
   End
   Begin VB.Image imgMas20 
      Height          =   285
      Left            =   8340
      MouseIcon       =   "frmSkills3.frx":337E
      MousePointer    =   99  'Custom
      Top             =   4080
      Width           =   345
   End
   Begin VB.Image imgMenos20 
      Height          =   285
      Left            =   7500
      MouseIcon       =   "frmSkills3.frx":34D0
      MousePointer    =   99  'Custom
      Top             =   4080
      Width           =   345
   End
   Begin VB.Image imgAceptar 
      Height          =   360
      Left            =   7200
      MouseIcon       =   "frmSkills3.frx":3622
      MousePointer    =   99  'Custom
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Label puntos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   300
      Left            =   5010
      TabIndex        =   0
      Top             =   255
      Width           =   165
   End
End
Attribute VB_Name = "frmSkills3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Evolution Online 0.11.6
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Evolution Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'

Option Explicit

Private clsFormulario As clsFormMovementManager

Private cBotonMas(1 To NUMSKILLS) As clsGraphicalButton
Private cBotonMenos(1 To NUMSKILLS) As clsGraphicalButton
Private cSkillNames(1 To NUMSKILLS) As clsGraphicalButton
Private cBtonAceptar As clsGraphicalButton
Private cBotonCancelar As clsGraphicalButton

Public LastPressed As clsGraphicalButton

Private bPuedeMagia As Boolean
Private bPuedeMeditar As Boolean
Private bPuedeEscudo As Boolean
Private bPuedeCombateDistancia As Boolean

Private vsHelp(1 To NUMSKILLS) As String

Private Sub Form_Load()

    MirandoAsignarSkills = True

    ' Handles Form movement (drag and drop).
    Set clsFormulario = New clsFormMovementManager
    clsFormulario.Initialize Me

    'Flags para saber que skills se modificaron
    ReDim flags(1 To NUMSKILLS)

    Call ValidarSkills

    Me.Picture = LoadPicture(DirGraficos & "VentanaSkills.jpg")
    Call LoadButtons

    Call LoadHelp
End Sub

Private Sub LoadButtons()
    Dim GrhPath As String
    Dim i As Long

    GrhPath = DirGraficos


    For i = 1 To NUMSKILLS
        Set cBotonMas(i) = New clsGraphicalButton
        Set cBotonMenos(i) = New clsGraphicalButton
        Set cSkillNames(i) = New clsGraphicalButton
    Next i

    Set cBtonAceptar = New clsGraphicalButton
    Set cBotonCancelar = New clsGraphicalButton

    Set LastPressed = New clsGraphicalButton
End Sub

Private Sub SumarSkillPoint(ByVal SkillIndex As Integer)
    If Alocados > 0 Then

        If Val(Text1(SkillIndex).Caption) < MAXSKILLPOINTS Then
            Text1(SkillIndex).Caption = Val(Text1(SkillIndex).Caption) + 1
            flags(SkillIndex) = flags(SkillIndex) + 1
            Alocados = Alocados - 1
        End If

    End If

    puntos.Caption = Alocados
End Sub

Private Sub RestarSkillPoint(ByVal SkillIndex As Integer)
    If Alocados < SkillPoints Then

        If Val(Text1(SkillIndex).Caption) > 0 And flags(SkillIndex) > 0 Then
            Text1(SkillIndex).Caption = Val(Text1(SkillIndex).Caption) - 1
            flags(SkillIndex) = flags(SkillIndex) - 1
            Alocados = Alocados + 1
        End If
    End If

    puntos.Caption = Alocados
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastPressed.ToggleToNormal
    lblHelp.Caption = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MirandoAsignarSkills = False
End Sub

Private Sub imgAceptar_Click()
    Dim skillChanges(NUMSKILLS) As Byte
    Dim i As Long

    For i = 1 To NUMSKILLS
        skillChanges(i) = CByte(Text1(i).Caption) - UserSkills(i)
        'Actualizamos nuestros datos locales
        UserSkills(i) = Val(Text1(i).Caption)
    Next i

    Call WriteModifySkills(skillChanges())

    If Alocados = 0 Then Call frmMain.LightSkillStar(False)

    SkillPoints = Alocados

    Unload Me
End Sub

Private Sub imgApunialar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ShowHelp(eSkill.Apuñalar)
End Sub

Private Sub imgCancelar_Click()
    Unload Me
End Sub

Private Sub imgCarpinteria_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ShowHelp(eSkill.Carpinteria)
End Sub

Private Sub imgCombateArmas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ShowHelp(eSkill.Armas)
End Sub

Private Sub imgCombateDistancia_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ShowHelp(eSkill.Proyectiles)
End Sub

Private Sub imgCombateSinArmas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ShowHelp(eSkill.Wrestling)
End Sub

Private Sub imgComercio_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ShowHelp(eSkill.Comerciar)
End Sub

Private Sub imgDomar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ShowHelp(eSkill.Domar)
End Sub

Private Sub imgEscudos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ShowHelp(eSkill.Defensa)
End Sub

Private Sub imgEvasion_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ShowHelp(eSkill.Tacticas)
End Sub

Private Sub imgHerreria_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ShowHelp(eSkill.Herreria)
End Sub

Private Sub imgLiderazgo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ShowHelp(eSkill.Liderazgo)
End Sub

Private Sub imgMagia_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ShowHelp(eSkill.Magia)
End Sub

Private Sub imgMas1_Click()
    Call SumarSkillPoint(1)
End Sub

Private Sub imgMas10_Click()
    Call SumarSkillPoint(10)
End Sub

Private Sub imgMas11_Click()
    Call SumarSkillPoint(11)
End Sub

Private Sub imgMas12_Click()
    Call SumarSkillPoint(12)
End Sub

Private Sub imgMas13_Click()
    Call SumarSkillPoint(13)
End Sub

Private Sub imgMas14_Click()
    Call SumarSkillPoint(14)
End Sub

Private Sub imgMas15_Click()
    Call SumarSkillPoint(15)
End Sub

Private Sub imgMas16_Click()
    Call SumarSkillPoint(16)
End Sub

Private Sub imgMas17_Click()
    Call SumarSkillPoint(17)
End Sub

Private Sub imgMas18_Click()
    Call SumarSkillPoint(18)
End Sub

Private Sub imgMas19_Click()
    Call SumarSkillPoint(19)
End Sub

Private Sub imgMas2_Click()
    Call SumarSkillPoint(2)
End Sub

Private Sub imgMas20_Click()
    Call SumarSkillPoint(20)
End Sub

Private Sub imgMas3_Click()
    Call SumarSkillPoint(3)
End Sub

Private Sub imgMas4_Click()
    Call SumarSkillPoint(4)
End Sub

Private Sub imgMas5_Click()
    Call SumarSkillPoint(5)
End Sub

Private Sub imgMas6_Click()
    Call SumarSkillPoint(6)
End Sub

Private Sub imgMas7_Click()
    Call SumarSkillPoint(7)
End Sub

Private Sub imgMas8_Click()
    Call SumarSkillPoint(8)
End Sub

Private Sub imgMas9_Click()
    Call SumarSkillPoint(9)
End Sub

Private Sub imgMeditar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ShowHelp(eSkill.Meditar)
End Sub

Private Sub imgMenos1_Click()
    Call RestarSkillPoint(1)
End Sub

Private Sub imgMenos10_Click()
    Call RestarSkillPoint(10)
End Sub

Private Sub imgMenos11_Click()
    Call RestarSkillPoint(11)
End Sub

Private Sub imgMenos12_Click()
    Call RestarSkillPoint(12)
End Sub

Private Sub imgMenos13_Click()
    Call RestarSkillPoint(13)
End Sub

Private Sub imgMenos14_Click()
    Call RestarSkillPoint(14)
End Sub

Private Sub imgMenos15_Click()
    Call RestarSkillPoint(15)
End Sub

Private Sub imgMenos16_Click()
    Call RestarSkillPoint(16)
End Sub

Private Sub imgMenos17_Click()
    Call RestarSkillPoint(17)
End Sub

Private Sub imgMenos18_Click()
    Call RestarSkillPoint(18)
End Sub

Private Sub imgMenos19_Click()
    Call RestarSkillPoint(19)
End Sub

Private Sub imgMenos2_Click()
    Call RestarSkillPoint(2)
End Sub

Private Sub imgMenos20_Click()
    Call RestarSkillPoint(20)
End Sub

Private Sub imgMenos3_Click()
    Call RestarSkillPoint(3)
End Sub

Private Sub imgMenos4_Click()
    Call RestarSkillPoint(4)
End Sub

Private Sub imgMenos5_Click()
    Call RestarSkillPoint(5)
End Sub

Private Sub imgMenos6_Click()
    Call RestarSkillPoint(6)
End Sub

Private Sub imgMenos7_Click()
    Call RestarSkillPoint(7)
End Sub

Private Sub imgMenos8_Click()
    Call RestarSkillPoint(8)
End Sub

Private Sub imgMenos9_Click()
    Call RestarSkillPoint(9)
End Sub

Private Sub LoadHelp()

    vsHelp(eSkill.Magia) = "Magia:" & vbCrLf & _
                           "- Representa la habilidad de un personaje de las áreas mágica." & vbCrLf & _
                           "- Indica la variedad de hechizos que es capaz de dominar el personaje."
    If Not bPuedeMagia Then
        vsHelp(eSkill.Magia) = vsHelp(eSkill.Magia) & vbCrLf & _
                               "* Habilidad inhabilitada para tu clase."
    End If

    vsHelp(eSkill.Robar) = "Robar:" & vbCrLf & _
                           "- Habilidades de hurto. Nunca por medio de la violencia." & vbCrLf & _
                           "- Indica la probabilidad de éxito del personaje al intentar apoderarse de oro de otro, en caso de ser Ladrón, tambien podrá apoderarse de items."

    vsHelp(eSkill.Tacticas) = "Evasión en Combate:" & vbCrLf & _
                              "- Representa la habilidad general para moverse en combate entre golpes enemigos sin morir o tropezar en el intento." & vbCrLf & _
                              "- Indica la posibilidad de evadir un golpe físico del personaje."

    vsHelp(eSkill.Armas) = "Combate con Armas:" & vbCrLf & _
                           "- Representa la habilidad del personaje para manejar armas de combate cuerpo a cuerpo." & vbCrLf & _
                           "- Indica la probabilidad de impactar al oponente con armas cuerpo a cuerpo."

    vsHelp(eSkill.Meditar) = "Meditar:" & vbCrLf & _
                             "- Representa la capacidad del personaje de concentrarse para abstrarse dentro de su mente, y así revitalizar su fuerza espiritual." & vbCrLf & _
                             "- Indica la velocidad a la que el personaje recupera maná (Clases mágicas)."

    If Not bPuedeMeditar Then
        vsHelp(eSkill.Meditar) = vsHelp(eSkill.Meditar) & vbCrLf & _
                                 "* Habilidad inhabilitada para tu clase."
    End If

    vsHelp(eSkill.Apuñalar) = "Apuñalar:" & vbCrLf & _
                              "- Representa la destreza para inflingir daño grave con armas cortas." & vbCrLf & _
                              "- Indica la posibilidad de apuñalar al enemigo en un ataque. El Asesino es la única clase que no necesitará 10 skills para comenzar a entrenar esta habilidad."

    vsHelp(eSkill.Ocultarse) = "Ocultarse:" & vbCrLf & _
                               "- La habilidad propia de un personaje para mimetizarse con el medio y evitar se perciba su presencia." & vbCrLf & _
                               "- Indica la facilidad con la que uno puede desaparecer de la vista de los demás y por cuanto tiempo."

    vsHelp(eSkill.Supervivencia) = "Superivencia:" & vbCrLf & _
                                   "- Es el conjunto de habilidades necesarias para sobrevivir fuera de una ciudad en base a lo que la naturaleza ofrece." & vbCrLf & _
                                   "- Permite conocer la salud de las criaturas guiándose exclusivamente por su aspecto, así como encender fogatas junto a las que descansar."

    vsHelp(eSkill.Talar) = "Talar:" & vbCrLf & _
                           "- Es la habilidad en el uso del hacha para evitar desperdiciar leña y maximizar la efectividad de cada golpe dado." & vbCrLf & _
                           "- Indica la probabilidad de obtener leña por golpe."

    vsHelp(eSkill.Comerciar) = "Comercio:" & vbCrLf & _
                               "- Es la habilidad para regatear los precios exigidos en la compra y evitar ser regateado al vender." & vbCrLf & _
                               "- Indica que tan caro se compra en el comercio con NPCs."

    vsHelp(eSkill.Defensa) = "Defensa con Escudos:" & vbCrLf & _
                             "- Es la habilidad de interponer correctamente el escudo ante cada embate enemigo para evitar ser impactado sin perder el equilibrio y poder responder rápidamente con la otra mano." & vbCrLf & _
                             "- Indica las probabilidades de bloquear un impacto con el escudo."

    If Not bPuedeEscudo Then
        vsHelp(eSkill.Defensa) = vsHelp(eSkill.Defensa) & vbCrLf & _
                                 "* Habilidad inhabilitada para tu clase."
    End If


    vsHelp(eSkill.Pesca) = "Pesca:" & vbCrLf & _
                           "- Es el conjunto de conocimientos básicos para poder armar un señuelo, poner la carnada en el anzuelo y saber dónde buscar peces." & vbCrLf & _
                           "- Indica la probabilidad de tener éxito en cada intento de pescar."

    vsHelp(eSkill.Mineria) = "Minería:" & vbCrLf & _
                             "- Es el conjunto de conocimientos sobre los distintos minerales, el dónde se obtienen, cómo deben ser extraídos y trabajados." & vbCrLf & _
                             "- Indica la probabilidad de tener éxito en cada intento de minar y la capacidad, o no de convertir estos minerales en lingotes."

    vsHelp(eSkill.Carpinteria) = "Carpintería:" & vbCrLf & _
                                 "- Es el conjunto de conocimientos para saber serruchar, lijar, encolar y clavar madera con un buen nivel de terminación." & vbCrLf & _
                                 "- Indica la habilidad en el manejo de estas herramientas, el que tan bueno se es en el oficio de carpintero."

    vsHelp(eSkill.Herreria) = "Herrería:" & vbCrLf & _
                              "- Es el conjunto de conocimientos para saber procesar cada tipo de mineral para fundirlo, forjarlo y crear aleaciones." & vbCrLf & _
                              "- Indica la habilidad en el manejo de estas técnicas, el que tan bueno se es en el oficio de herrero."

    vsHelp(eSkill.Liderazgo) = "Liderazgo:" & vbCrLf & _
                               "- Es la habilidad propia del personaje para convencer a otros a seguirlo en batalla." & vbCrLf & _
                               "- Permite crear clanes y partys"

    vsHelp(eSkill.Domar) = "Domar Animales:" & vbCrLf & _
                           "- Es la habilidad en el trato con animales para que estos te sigan y ayuden en combate." & vbCrLf & _
                           "- Indica la posibilidad de lograr domar a una criatura y qué clases de criaturas se puede domar."

    vsHelp(eSkill.Proyectiles) = "Combate a distancia:" & vbCrLf & _
                                 "- Es el manejo de las armas de largo alcance." & vbCrLf & _
                                 "- Indica la probabilidad de éxito para impactar a un enemigo con este tipo de armas."

    If Not bPuedeCombateDistancia Then
        vsHelp(eSkill.Proyectiles) = vsHelp(eSkill.Proyectiles) & vbCrLf & _
                                     "* Habilidad inhabilitada para tu clase."
    End If

    vsHelp(eSkill.Wrestling) = "Combate sin armas:" & vbCrLf & _
                               "- Es la habilidad del personaje para entrar en combate sin arma alguna salvo sus propios brazos." & vbCrLf & _
                               "- Indica la probabilidad de éxito para impactar a un enemigo estando desarmado. El Bandido y Ladrón tienen habilidades extras asociadas a esta habilidad."

    vsHelp(eSkill.Navegacion) = "Navegación:" & vbCrLf & _
                                "- Es la habilidad para controlar barcos en el mar sin naufragar." & vbCrLf & _
                                "- Indica que clase de barcos se pueden utilizar."

End Sub

Private Sub imgMineria_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ShowHelp(eSkill.Mineria)
End Sub

Private Sub imgNavegacion_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ShowHelp(eSkill.Navegacion)
End Sub

Private Sub imgOcultarse_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ShowHelp(eSkill.Ocultarse)
End Sub

Private Sub imgPesca_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ShowHelp(eSkill.Pesca)
End Sub

Private Sub imgRobar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ShowHelp(eSkill.Robar)
End Sub

Private Sub imgSupervivencia_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ShowHelp(eSkill.Supervivencia)
End Sub

Private Sub imgTalar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ShowHelp(eSkill.Talar)
End Sub

Private Sub lblHelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LastPressed.ToggleToNormal
End Sub

Private Sub ShowHelp(ByVal eeSkill As eSkill)
    lblHelp.Caption = vsHelp(eeSkill)
End Sub

Private Sub ValidarSkills()

    bPuedeMagia = True
    bPuedeMeditar = True
    bPuedeEscudo = True
    bPuedeCombateDistancia = True

    Select Case UserClase
        Case eClass.Warrior, eClass.Hunter, eClass.Worker, eClass.Thief
            bPuedeMagia = False
            bPuedeMeditar = False

        Case eClass.Pirat
            bPuedeMagia = False
            bPuedeMeditar = False
            bPuedeEscudo = False

        Case eClass.Mage, eClass.Druid
            bPuedeEscudo = False
            bPuedeCombateDistancia = False

    End Select

    ' Magia
    imgMas1.Enabled = bPuedeMagia
    imgMenos1.Enabled = bPuedeMagia

    ' Meditar
    imgMas5.Enabled = bPuedeMeditar
    imgMenos5.Enabled = bPuedeMeditar

    ' Escudos
    imgMas11.Enabled = bPuedeEscudo
    imgMenos11.Enabled = bPuedeEscudo

    ' Proyectiles
    imgMas18.Enabled = bPuedeCombateDistancia
    imgMenos18.Enabled = bPuedeCombateDistancia
End Sub
