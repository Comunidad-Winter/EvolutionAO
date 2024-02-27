VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.ocx"
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "CSWSK32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   ClientHeight    =   8700
   ClientLeft      =   360
   ClientTop       =   300
   ClientWidth     =   12000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "frmMain.frx":000C
   ScaleHeight     =   580
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin SocketWrenchCtrl.Socket Socket1 
      Left            =   6960
      Top             =   6360
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   0   'False
      Backlog         =   1
      Binary          =   -1  'True
      Blocking        =   0   'False
      Broadcast       =   0   'False
      BufferSize      =   10240
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   10000
      Type            =   1
      Urgent          =   0   'False
   End
   Begin VB.PictureBox Minimap 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   1500
      Left            =   6795
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   41
      Top             =   75
      Width           =   1500
      Begin VB.Image imgChar 
         Height          =   75
         Left            =   720
         Top             =   720
         Width           =   75
      End
   End
   Begin VB.Timer TimerPociones 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6480
      Top             =   7440
   End
   Begin VB.TextBox SendTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   165
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1650
      Visible         =   0   'False
      Width           =   8145
   End
   Begin VB.PictureBox picSM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   3
      Left            =   11325
      MouseIcon       =   "frmMain.frx":2BBF7
      MousePointer    =   99  'Custom
      ScaleHeight     =   450
      ScaleWidth      =   420
      TabIndex        =   29
      ToolTipText     =   "Macro trabajo"
      Top             =   8400
      Width           =   420
   End
   Begin VB.PictureBox picSM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   2
      Left            =   10875
      MouseIcon       =   "frmMain.frx":2BD49
      MousePointer    =   99  'Custom
      ScaleHeight     =   450
      ScaleWidth      =   420
      TabIndex        =   28
      ToolTipText     =   "Macro hechizo"
      Top             =   8400
      Width           =   420
   End
   Begin VB.PictureBox picSM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   1
      Left            =   10440
      MouseIcon       =   "frmMain.frx":2BE9B
      MousePointer    =   99  'Custom
      ScaleHeight     =   450
      ScaleWidth      =   420
      TabIndex        =   27
      ToolTipText     =   "Seguro criminalidad"
      Top             =   8400
      Width           =   420
   End
   Begin VB.PictureBox picSM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   0
      Left            =   10005
      MouseIcon       =   "frmMain.frx":2BFED
      MousePointer    =   99  'Custom
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   26
      ToolTipText     =   "Seguro resurrección"
      Top             =   8400
      Width           =   420
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   7440
      Top             =   6360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer macrotrabajo 
      Enabled         =   0   'False
      Left            =   6960
      Top             =   6960
   End
   Begin VB.Timer TrainingMacro 
      Enabled         =   0   'False
      Interval        =   3121
      Left            =   7440
      Top             =   6960
   End
   Begin VB.TextBox SendCMSTXT 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   240
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   8355
      Visible         =   0   'False
      Width           =   8025
   End
   Begin VB.Timer Macro 
      Interval        =   750
      Left            =   6480
      Top             =   6960
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   6480
      Top             =   6360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Second 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   7440
      Top             =   7440
   End
   Begin VB.Timer SpoofCheck 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   6960
      Top             =   7440
   End
   Begin RichTextLib.RichTextBox RecTxt 
      Height          =   1500
      Left            =   165
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   75
      Width           =   6540
      _ExtentX        =   11536
      _ExtentY        =   2646
      _Version        =   393217
      BackColor       =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmMain.frx":2C13F
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2400
      Left            =   9000
      ScaleHeight     =   160
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   163
      TabIndex        =   16
      Top             =   2610
      Width           =   2445
   End
   Begin VB.ListBox hlst 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Source Sans Pro"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   1950
      Left            =   9000
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2610
      Visible         =   0   'False
      Width           =   2430
   End
   Begin VB.Label imgRetos 
      BackStyle       =   0  'Transparent
      Height          =   240
      Left            =   10395
      MouseIcon       =   "frmMain.frx":2C1BC
      MousePointer    =   99  'Custom
      TabIndex        =   40
      Top             =   7830
      Width           =   1230
   End
   Begin VB.Label btnFacebook 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   11280
      MouseIcon       =   "frmMain.frx":2C30E
      MousePointer    =   99  'Custom
      TabIndex        =   39
      Top             =   5310
      Width           =   375
   End
   Begin VB.Label btnDiscord 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   8640
      MouseIcon       =   "frmMain.frx":2C460
      MousePointer    =   99  'Custom
      TabIndex        =   38
      Top             =   5295
      Width           =   375
   End
   Begin VB.Label lblFPS 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Fps: 100"
      ForeColor       =   &H0080C0FF&
      Height          =   180
      Left            =   8760
      MouseIcon       =   "frmMain.frx":2C5B2
      MousePointer    =   99  'Custom
      TabIndex        =   25
      Top             =   8400
      Width           =   915
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   10170
      MouseIcon       =   "frmMain.frx":2C704
      MousePointer    =   99  'Custom
      TabIndex        =   19
      Top             =   2040
      Width           =   1515
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   8565
      MouseIcon       =   "frmMain.frx":2C856
      MousePointer    =   99  'Custom
      TabIndex        =   18
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label lblDonaciones 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   210
      Left            =   600
      MouseIcon       =   "frmMain.frx":2C9A8
      MousePointer    =   99  'Custom
      TabIndex        =   37
      Top             =   1680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label btnManual 
      BackStyle       =   0  'Transparent
      Height          =   300
      Left            =   9705
      MouseIcon       =   "frmMain.frx":2CAFA
      MousePointer    =   99  'Custom
      TabIndex        =   36
      ToolTipText     =   "Manual"
      Top             =   5280
      Width           =   300
   End
   Begin VB.Label btnCerrar 
      BackStyle       =   0  'Transparent
      Height          =   150
      Left            =   11610
      MouseIcon       =   "frmMain.frx":2CC4C
      MousePointer    =   99  'Custom
      TabIndex        =   35
      Top             =   255
      Width           =   135
   End
   Begin VB.Label btnMinimizar 
      BackStyle       =   0  'Transparent
      Height          =   135
      Left            =   11340
      MouseIcon       =   "frmMain.frx":2CD9E
      MousePointer    =   99  'Custom
      TabIndex        =   34
      Top             =   240
      Width           =   135
   End
   Begin VB.Label btnResetPJ 
      BackStyle       =   0  'Transparent
      Height          =   300
      Left            =   10800
      MouseIcon       =   "frmMain.frx":2CEF0
      MousePointer    =   99  'Custom
      TabIndex        =   32
      ToolTipText     =   "Resetear PJ"
      Top             =   5280
      Width           =   300
   End
   Begin VB.Label btnRanking 
      BackStyle       =   0  'Transparent
      Height          =   300
      Left            =   9195
      MouseIcon       =   "frmMain.frx":2D042
      MousePointer    =   99  'Custom
      TabIndex        =   33
      ToolTipText     =   "Ranking"
      Top             =   5280
      Width           =   300
   End
   Begin VB.Label lblPorcLvl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   210
      Left            =   9450
      TabIndex        =   21
      Top             =   1125
      Width           =   315
   End
   Begin VB.Label lblOnline 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   180
      Left            =   1800
      TabIndex        =   31
      Top             =   1680
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image btnMapa 
      Height          =   300
      Left            =   10275
      MouseIcon       =   "frmMain.frx":2D194
      MousePointer    =   99  'Custom
      ToolTipText     =   "Mapa"
      Top             =   5280
      Width           =   300
   End
   Begin VB.Image imgClanes 
      Height          =   240
      Left            =   10395
      MouseIcon       =   "frmMain.frx":2D2E6
      MousePointer    =   99  'Custom
      Top             =   7080
      Width           =   1230
   End
   Begin VB.Image imgEstadisticas 
      Height          =   240
      Left            =   10395
      MouseIcon       =   "frmMain.frx":2D438
      MousePointer    =   99  'Custom
      Top             =   6360
      Width           =   1230
   End
   Begin VB.Image imgOpciones 
      Height          =   240
      Left            =   10395
      MouseIcon       =   "frmMain.frx":2D58A
      MousePointer    =   99  'Custom
      Top             =   6705
      Width           =   1230
   End
   Begin VB.Image imgGrupo 
      Height          =   240
      Left            =   10395
      MouseIcon       =   "frmMain.frx":2D6DC
      MousePointer    =   99  'Custom
      Top             =   7455
      Width           =   1230
   End
   Begin VB.Image imgAsignarSkill 
      Height          =   150
      Left            =   8730
      MouseIcon       =   "frmMain.frx":2D82E
      MousePointer    =   99  'Custom
      Top             =   1200
      Width           =   165
   End
   Begin VB.Label lblDropGold 
      BackStyle       =   0  'Transparent
      Height          =   345
      Left            =   10440
      MouseIcon       =   "frmMain.frx":2D980
      MousePointer    =   99  'Custom
      TabIndex        =   30
      Top             =   5820
      Width           =   300
   End
   Begin VB.Image cmdInfo 
      Height          =   450
      Left            =   11520
      MouseIcon       =   "frmMain.frx":2DAD2
      MousePointer    =   99  'Custom
      ToolTipText     =   "Información"
      Top             =   3885
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   240
      Index           =   0
      Left            =   11550
      MouseIcon       =   "frmMain.frx":2DC24
      MousePointer    =   99  'Custom
      Top             =   3330
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   240
      Index           =   1
      Left            =   11550
      MouseIcon       =   "frmMain.frx":2DD76
      MousePointer    =   99  'Custom
      Top             =   3000
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Bramh"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   225
      Left            =   9255
      TabIndex        =   23
      Top             =   690
      Width           =   2325
   End
   Begin VB.Label lblLvl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   240
      Left            =   8745
      TabIndex        =   22
      ToolTipText     =   "Nivel"
      Top             =   945
      Width           =   135
   End
   Begin VB.Label lblExp 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exp: 10/10"
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
      Height          =   210
      Left            =   9735
      MouseIcon       =   "frmMain.frx":2DEC8
      MousePointer    =   99  'Custom
      TabIndex        =   20
      Top             =   1125
      Width           =   1980
   End
   Begin VB.Image CmdLanzar 
      Height          =   435
      Left            =   8760
      MouseIcon       =   "frmMain.frx":2E01A
      MousePointer    =   99  'Custom
      ToolTipText     =   "Lanzar hechizo"
      Top             =   4650
      Visible         =   0   'False
      Width           =   2925
   End
   Begin VB.Label GldLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Left            =   10860
      TabIndex        =   15
      Top             =   5880
      Width           =   105
   End
   Begin VB.Label lblStrg 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   210
      Left            =   7875
      TabIndex        =   9
      ToolTipText     =   "Fuerza"
      Top             =   8535
      Width           =   210
   End
   Begin VB.Label lblDext 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Left            =   7290
      TabIndex        =   8
      ToolTipText     =   "Evasión"
      Top             =   8535
      Width           =   210
   End
   Begin VB.Label lblWeapon 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "000/000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Left            =   4245
      TabIndex        =   6
      Top             =   8535
      Width           =   855
   End
   Begin VB.Label lblShielder 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00/00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Left            =   5850
      TabIndex        =   5
      Top             =   8535
      Width           =   855
   End
   Begin VB.Label lblHelm 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00/00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Left            =   2520
      TabIndex        =   4
      Top             =   8535
      Width           =   855
   End
   Begin VB.Label lblArmor 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00/00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   8535
      Width           =   855
   End
   Begin VB.Shape MainViewShp 
      BackColor       =   &H80000001&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   6240
      Left            =   195
      Top             =   1995
      Width           =   8160
   End
   Begin VB.Shape ExpShp 
      BackColor       =   &H00000040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000002&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00404040&
      Height          =   195
      Left            =   9270
      Top             =   1140
      Width           =   2415
   End
   Begin VB.Label Coord 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "000 X:00 Y: 00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Left            =   8385
      MouseIcon       =   "frmMain.frx":2E16C
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   8595
      Width           =   1695
   End
   Begin VB.Label lblMapName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   8325
      MouseIcon       =   "frmMain.frx":2E2BE
      MousePointer    =   99  'Custom
      TabIndex        =   24
      Top             =   8595
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label lblMana 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "9999/9999"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   180
      Left            =   8610
      TabIndex        =   11
      Top             =   6951
      Width           =   1335
   End
   Begin VB.Shape shpMana 
      FillStyle       =   0  'Solid
      Height          =   180
      Left            =   8610
      Top             =   6960
      Width           =   1335
   End
   Begin VB.Label lblSed 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      Caption         =   "999/999"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   180
      Left            =   8610
      TabIndex        =   14
      Top             =   7882
      Width           =   1320
   End
   Begin VB.Label lblVida 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "999/999"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   180
      Left            =   8640
      TabIndex        =   12
      Top             =   6465
      Width           =   1335
   End
   Begin VB.Shape shpVida 
      FillStyle       =   0  'Solid
      Height          =   180
      Left            =   8610
      Top             =   6475
      Width           =   1335
   End
   Begin VB.Shape shpSed 
      FillStyle       =   0  'Solid
      Height          =   180
      Left            =   8610
      Top             =   7874
      Width           =   1335
   End
   Begin VB.Label lblEnergia 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "999/999"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   180
      Left            =   8610
      TabIndex        =   10
      Top             =   6008
      Width           =   1335
   End
   Begin VB.Shape shpEnergia 
      FillStyle       =   0  'Solid
      Height          =   180
      Left            =   8610
      Top             =   6030
      Width           =   1335
   End
   Begin VB.Label lblHambre 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "999/999"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   180
      Left            =   8610
      TabIndex        =   13
      Top             =   7431
      Width           =   1335
   End
   Begin VB.Shape shpHambre 
      FillStyle       =   0  'Solid
      Height          =   180
      Left            =   8610
      Top             =   7440
      Width           =   1335
   End
   Begin VB.Image imgRecHechizos 
      Appearance      =   0  'Flat
      Height          =   3720
      Index           =   1
      Left            =   8430
      Picture         =   "frmMain.frx":2E410
      Top             =   1470
      Visible         =   0   'False
      Width           =   3390
   End
   Begin VB.Image imgRecHechizos 
      Appearance      =   0  'Flat
      Height          =   3720
      Index           =   0
      Left            =   8430
      Picture         =   "frmMain.frx":33BD0
      Top             =   1470
      Visible         =   0   'False
      Width           =   3390
   End
   Begin VB.Menu mnuObj 
      Caption         =   "Objeto"
      Visible         =   0   'False
      Begin VB.Menu mnuTirar 
         Caption         =   "Tirar"
      End
      Begin VB.Menu mnuUsar 
         Caption         =   "Usar"
      End
      Begin VB.Menu mnuEquipar 
         Caption         =   "Equipar"
      End
   End
   Begin VB.Menu mnuNpc 
      Caption         =   "NPC"
      Visible         =   0   'False
      Begin VB.Menu mnuNpcDesc 
         Caption         =   "Descripcion"
      End
      Begin VB.Menu mnuNpcComerciar 
         Caption         =   "Comerciar"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmMain"
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

' @@ Anti XMouseButton
Private MOUSE_DOWN As Boolean
Private MOUSE_UP As Boolean

Public InMouseExp As Boolean
Private enInventario As Boolean
Private enHechizos As Boolean
Private enFondo As Boolean
Private enBoton As Boolean

Public Attack As Boolean
Public tX As Byte
Public tY As Byte
Public MouseX As Long
Public MouseY As Long
Public MouseBoton As Long
Public MouseShift As Long
Private clicX As Long
Private clicY As Long

Public IsPlaying As Byte

Private clsFormulario As clsFormMovementManager

Private cBotonDiamArriba As clsGraphicalButton
Private cBotonDiamAbajo As clsGraphicalButton
Private cBotonMapa As clsGraphicalButton
Private cBotonGrupo As clsGraphicalButton
Private cBotonOpciones As clsGraphicalButton
Private cBotonEstadisticas As clsGraphicalButton
Private cBotonClanes As clsGraphicalButton
Private cBotonAsignarSkill As clsGraphicalButton

Public LastPressed As clsGraphicalButton

Public picSkillStar As Picture

Public WithEvents dragInventory As clsGrapchicalInventory
Attribute dragInventory.VB_VarHelpID = -1

Dim PuedeMacrear As Boolean

Private Sub Command1_Click()
    frmRetos1vs1.Show
End Sub

Private Sub btnCerrar_Click()
    prgRun = False
End Sub

Private Sub btnDiscord_Click()
    Call ShellExecute(0, "Open", "https://discord.gg/gtwRYZx", "", App.path, SW_SHOWNORMAL)
End Sub

Private Sub btnFacebook_Click()
    Call ShellExecute(0, "Open", "https://www.facebook.com/groups/EvolutionAO", "", App.path, SW_SHOWNORMAL)
End Sub

Private Sub btnManual_Click()
    frmManual.Show
End Sub

Private Sub btnMapa_Click()
    Call frmMapa.Show(vbModeless, frmMain)
End Sub

Private Sub btnMinimizar_Click()
    Me.WindowState = 1
End Sub

Private Sub btnRanking_Click()
    Call frmMenuRanking.Show(vbModeless, frmMain)
End Sub

Private Sub btnResetPJ_Click()
    Select Case MsgBox("¿Quieres resetear el personaje? (Solo funciona hasta nivel 15)", vbInformation + vbYesNo, "Mensaje de Alerta")
        Case vbYes
            Call WriteReset
    End Select
End Sub

Private Sub Form_Load()

    ' Detect links in console
    EnableURLDetect RecTxt.hWnd, Me.hWnd
    Set dragInventory = Inventario

    If NoRes Then
        ' Handles Form movement (drag and drop).
        Set clsFormulario = New clsFormMovementManager
        clsFormulario.Initialize Me, 120
    End If

    If Not FileExist(App.path & "\init\cabezas.ind", vbNormal) Or Not FileExist(App.path & "\init\cabezas.ind", vbHidden) Then
        End
    End If

    Call LoadButtons

    Me.Left = 0
    Me.Top = 0
    
    imgChar.Picture = LoadPicture(DirGraficos & "Char.JPG")
    
End Sub

Private Sub LoadButtons()
    Dim GrhPath As String
    Dim i As Integer

    GrhPath = DirGraficos

    Set cBotonDiamArriba = New clsGraphicalButton
    Set cBotonDiamAbajo = New clsGraphicalButton
    Set cBotonGrupo = New clsGraphicalButton
    Set cBotonOpciones = New clsGraphicalButton
    Set cBotonEstadisticas = New clsGraphicalButton
    Set cBotonClanes = New clsGraphicalButton
    Set cBotonAsignarSkill = New clsGraphicalButton
    Set cBotonMapa = New clsGraphicalButton

    Set LastPressed = New clsGraphicalButton


    For i = 0 To 3
        picSM(i).MouseIcon = picMouseIcon
    Next i
End Sub

Public Sub LightSkillStar(ByVal bTurnOn As Boolean)
    If bTurnOn Then
        'Si hay skill call cambiar imagen
    Else
        'Si no hay skill no cambiar imagen
    End If
End Sub

Private Sub cmdMoverHechi_Click(Index As Integer)
    If hlst.Visible = True Then
        If hlst.ListIndex = -1 Then Exit Sub
        Dim sTemp As String

        Select Case Index
            Case 1    'subir
                If hlst.ListIndex = 0 Then Exit Sub
            Case 0    'bajar
                If hlst.ListIndex = hlst.ListCount - 1 Then Exit Sub
        End Select

        Call WriteMoveSpell(Index = 1, hlst.ListIndex + 1)

        Select Case Index
            Case 1    'subir
                sTemp = hlst.List(hlst.ListIndex - 1)
                hlst.List(hlst.ListIndex - 1) = hlst.List(hlst.ListIndex)
                hlst.List(hlst.ListIndex) = sTemp
                hlst.ListIndex = hlst.ListIndex - 1
            Case 0    'bajar
                sTemp = hlst.List(hlst.ListIndex + 1)
                hlst.List(hlst.ListIndex + 1) = hlst.List(hlst.ListIndex)
                hlst.List(hlst.ListIndex) = sTemp
                hlst.ListIndex = hlst.ListIndex + 1
        End Select
    End If
End Sub

Public Sub ActivarMacroHechizos()
    If Not hlst.Visible Then
        Call AddtoRichTextBox(frmMain.RecTxt, "Debes tener seleccionado el hechizo para activar el auto-lanzar", 0, 200, 200, False, True, True)
        Exit Sub
    End If

    TrainingMacro.Interval = INT_MACRO_HECHIS
    TrainingMacro.Enabled = True
    Call AddtoRichTextBox(frmMain.RecTxt, "Auto lanzar hechizos activado", 0, 200, 200, False, True, True)
    Call ControlSM(eSMType.mSpells, True)
End Sub

Public Sub DesactivarMacroHechizos()
    TrainingMacro.Enabled = False
    Call AddtoRichTextBox(frmMain.RecTxt, "Auto lanzar hechizos desactivado", 0, 150, 150, False, True, True)
    Call ControlSM(eSMType.mSpells, False)
End Sub

Public Sub ControlSM(ByVal Index As Byte, ByVal Mostrar As Boolean)
    Dim GrhIndex As Long
    Dim SR As RECT
    Dim DR As RECT

    GrhIndex = GRH_INI_SM + Index + SM_CANT * (CInt(Mostrar) + 1)

    With GrhData(GrhIndex)
        SR.Left = .sX
        SR.Right = SR.Left + .pixelWidth
        SR.Top = .sY
        SR.Bottom = SR.Top + .pixelHeight

        DR.Left = 0
        DR.Right = .pixelWidth
        DR.Top = 0
        DR.Bottom = .pixelHeight
    End With

    Call DrawGrhtoHdc(picSM(Index).hdc, GrhIndex, SR, DR)
    picSM(Index).Refresh

    Select Case Index
        Case eSMType.sResucitation
            If Mostrar Then
                Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_RESU_ON, 0, 255, 0, True, False, True)
                picSM(Index).ToolTipText = "Seguro de resucitación activado."
            Else
                Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_RESU_OFF, 255, 37, 37, True, False, True)
                picSM(Index).ToolTipText = "Seguro de resucitación desactivado."
            End If

        Case eSMType.sSafemode
            If Mostrar Then
                Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_ACTIVADO, 0, 255, 0, True, False, True)
                picSM(Index).ToolTipText = "Seguro activado."
            Else
                Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_DESACTIVADO, 255, 37, 37, True, False, True)
                picSM(Index).ToolTipText = "Seguro desactivado."
            End If

        Case eSMType.mSpells
            If Mostrar Then
                picSM(Index).ToolTipText = "Macro de hechizos activado."
            Else
                picSM(Index).ToolTipText = "Macro de hechizos desactivado."
            End If

        Case eSMType.mWork
            If Mostrar Then
                picSM(Index).ToolTipText = "Macro de trabajo activado."
            Else
                picSM(Index).ToolTipText = "Macro de trabajo desactivado."
            End If
    End Select

    SMStatus(Index) = Mostrar
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    '***************************************************
    'Autor: Unknown
    'Last Modification: 18/11/2009
    '18/11/2009: ZaMa - Ahora se pueden poner comandos en los mensajes personalizados (execpto guildchat y privados)
    '***************************************************

    If (Not SendTxt.Visible) And (Not SendCMSTXT.Visible) Then

        Select Case KeyCode
            Case vbKeyI
                If esGM(UserCharIndex) Then Call WriteInvisible
            Case vbKeyP
                If esGM(UserCharIndex) Then Call frmPanelGm.Show(vbModeless, frmMain)
        End Select

        'Checks if the key is valid
        If LenB(CustomKeys.ReadableName(KeyCode)) > 0 Then
            Select Case KeyCode
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleMusic)
                    Audio.MusicActivated = Not Audio.MusicActivated

                Case CustomKeys.BindedKey(eKeyType.mKeyToggleSound)
                    Audio.SoundActivated = Not Audio.SoundActivated

                Case CustomKeys.BindedKey(eKeyType.mKeyToggleFxs)
                    Audio.SoundEffectsActivated = Not Audio.SoundEffectsActivated

                Case CustomKeys.BindedKey(eKeyType.mKeyGetObject)
                    Call AgarrarItem

                Case CustomKeys.BindedKey(eKeyType.mKeyEquipObject)
                    Call EquiparItem

                Case CustomKeys.BindedKey(eKeyType.mKeyToggleNames)
                    Nombres = Not Nombres

                Case CustomKeys.BindedKey(eKeyType.mKeyTamAnimal)
                    If UserEstado = 1 Then
                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                        End With
                    Else
                        Call WriteWork(eSkill.Domar)
                    End If

                Case CustomKeys.BindedKey(eKeyType.mKeySteal)
                    If UserEstado = 1 Then
                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                        End With
                    Else
                        Call WriteWork(eSkill.Robar)
                    End If

                Case CustomKeys.BindedKey(eKeyType.mKeyHide)
                    If UserEstado = 1 Then
                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                        End With
                    Else
                        Call WriteWork(eSkill.Ocultarse)
                    End If

                Case CustomKeys.BindedKey(eKeyType.mKeyDropObject)
                    Call TirarItem

                Case CustomKeys.BindedKey(eKeyType.mKeyUseObject)
                    If macrotrabajo.Enabled Then Call DesactivarMacroTrabajo

                    If MainTimer.Check(TimersIndex.UseItemWithU) Then
                        Call UsarItem(1)
                    End If

                Case CustomKeys.BindedKey(eKeyType.mKeyRequestRefresh)
                    If MainTimer.Check(TimersIndex.SendRPU) Then
                        Call WriteRequestPositionUpdate
                        Beep
                    End If
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleSafeMode)
                    Call WriteSafeToggle

                Case CustomKeys.BindedKey(eKeyType.mKeyToggleResuscitationSafe)
                    Call WriteResuscitationToggle
            End Select
        Else
            Select Case KeyCode
                    'Custom messages!
                Case vbKey0 To vbKey9
                    Dim CustomMessage As String

                    CustomMessage = CustomMessages.Message((KeyCode - 39) Mod 10)
                    If LenB(CustomMessage) <> 0 Then
                        ' No se pueden mandar mensajes personalizados de clan o privado!
                        If UCase(Left(CustomMessage, 5)) <> "/CMSG" And _
                           Left(CustomMessage, 1) <> "\" Then

                            Call ParseUserCommand(CustomMessage)
                        End If
                    End If
            End Select
        End If
    End If

    Select Case KeyCode
        Case CustomKeys.BindedKey(eKeyType.mKeyTalkWithGuild)
            If SendTxt.Visible Then Exit Sub

            If (Not Comerciando) And (Not MirandoAsignarSkills) And _
               (Not frmMSG.Visible) And (Not MirandoForo) And _
               (Not frmEstadisticas.Visible) And (Not frmCantidad.Visible) Then
                SendCMSTXT.Visible = True
                SendCMSTXT.SetFocus
            End If

        Case CustomKeys.BindedKey(eKeyType.mKeyTakeScreenShot)
            Call ScreenCapture

        Case CustomKeys.BindedKey(eKeyType.mKeyShowOptions)
            Call frmOpciones.Show(vbModeless, frmMain)

        Case CustomKeys.BindedKey(eKeyType.mKeyMeditate)
            If UserMinMAN = UserMaxMAN Then Exit Sub

            If UserEstado = 1 Then
                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                End With
                Exit Sub
            End If

            If Not PuedeMacrear Then
                AddtoRichTextBox frmMain.RecTxt, "No tan rápido..!", 255, 255, 255, False, False, True
            Else
                Call WriteMeditate
                PuedeMacrear = False
            End If

        Case CustomKeys.BindedKey(eKeyType.mKeyCastSpellMacro)
            If UserEstado = 1 Then
                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                End With
                Exit Sub
            End If

            If TrainingMacro.Enabled Then
                DesactivarMacroHechizos
            Else
                ActivarMacroHechizos
            End If

        Case CustomKeys.BindedKey(eKeyType.mKeyWorkMacro)
            If UserEstado = 1 Then
                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                End With
                Exit Sub
            End If

            If macrotrabajo.Enabled Then
                Call DesactivarMacroTrabajo
            Else
                Call ActivarMacroTrabajo
            End If

        Case CustomKeys.BindedKey(eKeyType.mKeyExitGame)
            If frmMain.macrotrabajo.Enabled Then Call DesactivarMacroTrabajo
            Call WriteQuit

        Case CustomKeys.BindedKey(eKeyType.mKeyAttack)
            If Shift <> 0 Then Exit Sub

            If Not MainTimer.Check(TimersIndex.Arrows, False) Then Exit Sub    'Check if arrows interval has finished.
            If Not MainTimer.Check(TimersIndex.CastSpell, False) Then    'Check if spells interval has finished.
                If Not MainTimer.Check(TimersIndex.CastAttack) Then Exit Sub    'Corto intervalo Golpe-Hechizo
            Else
                If Not MainTimer.Check(TimersIndex.Attack) Or UserDescansar Or UserMeditar Then Exit Sub
            End If

            If TrainingMacro.Enabled Then DesactivarMacroHechizos
            If macrotrabajo.Enabled Then DesactivarMacroTrabajo
            Call WriteAttack
            Attack = True
            charlist(UserCharIndex).Arma.WeaponWalk(charlist(UserCharIndex).Heading).Started = 1
            charlist(UserCharIndex).Escudo.ShieldWalk(charlist(UserCharIndex).Heading).Started = 1

        Case CustomKeys.BindedKey(eKeyType.mKeyTalk)
            If SendCMSTXT.Visible Then Exit Sub

            If (Not Comerciando) And (Not MirandoAsignarSkills) And _
               (Not frmMSG.Visible) And (Not MirandoForo) And _
               (Not frmEstadisticas.Visible) And (Not frmCantidad.Visible) Then
                SendTxt.Visible = True
                SendTxt.SetFocus
            End If

    End Select
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseBoton = Button
    MouseShift = Shift
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    clicX = X
    clicY = Y
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If prgRun = True Then
        prgRun = False
        Cancel = 1
    End If
End Sub

Private Sub Form_Resize()
    MainWindowState = Me.WindowState
End Sub

Private Sub imgAsignarSkill_Click()
    Dim i As Integer

    LlegaronSkills = False
    Call WriteRequestSkills
    Call FlushBuffer

    Do While Not LlegaronSkills
        DoEvents    'esperamos a que lleguen y mantenemos la interfaz viva
    Loop
    LlegaronSkills = False

    For i = 1 To NUMSKILLS
        frmSkills3.Text1(i).Caption = UserSkills(i)
    Next i

    Alocados = SkillPoints
    frmSkills3.puntos.Caption = SkillPoints
    frmSkills3.Show , frmMain

End Sub

Private Sub ImgClanes_Click()
    If frmGuildLeader.Visible Then Unload frmGuildLeader
    Call WriteRequestGuildLeaderInfo
End Sub

Private Sub imgEstadisticas_Click()
    LlegaronAtrib = False
    LlegaronSkills = False
    LlegoFama = False
    Call WriteRequestAtributes
    Call WriteRequestSkills
    Call WriteRequestMiniStats
    Call WriteRequestFame
    Call FlushBuffer
    Do While Not LlegaronSkills Or Not LlegaronAtrib Or Not LlegoFama
        DoEvents    'esperamos a que lleguen y mantenemos la interfaz viva
    Loop
    frmEstadisticas.Iniciar_Labels
    frmEstadisticas.Show , frmMain
    LlegaronAtrib = False
    LlegaronSkills = False
    LlegoFama = False
End Sub


Private Sub imgGrupo_Click()
    Call WriteRequestPartyForm
End Sub

Private Sub imgInvScrollDown_Click()
    Call Inventario.ScrollInventory(True)
End Sub

Private Sub imgInvScrollUp_Click()
    Call Inventario.ScrollInventory(False)
End Sub


Private Sub imgOpciones_Click()
    Call frmOpciones.Show(vbModeless, frmMain)
End Sub

Private Sub lblScroll_Click(Index As Integer)
    Inventario.ScrollInventory (Index = 0)
End Sub

Private Sub Label1_Click()
    prgRun = False
End Sub

Private Sub Label2_Click()

End Sub



Private Sub lblEstadisticas_Click()
    Call WriteRequestStats
End Sub

Private Sub imgRetos_Click()
    frmRetos1vs1.Show
End Sub

Private Sub lblDonaciones_Click()
    frmDonaciones.Show
End Sub

Private Sub lblExp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    InMouseExp = True
    lblExp.Caption = UserExp & "/" & UserPasarNivel
    If UserPasarNivel = 0 Then
        lblExp.Caption = "¡Nivel máximo!"
    End If
End Sub

Private Sub lblResetearPersonaje_Click()
    Select Case MsgBox("¿Quieres resetear el personaje? (Solo funciona hasta nivel 15)", vbInformation + vbYesNo, "Mensaje de Alerta")
        Case vbYes
            Call WriteReset
    End Select
End Sub

Private Sub Macro_Timer()
    PuedeMacrear = True
End Sub

Private Sub macrotrabajo_Timer()
    If Inventario.SelectedItem = 0 Then
        Call DesactivarMacroTrabajo
        Exit Sub
    End If

    If UsingSkill = eSkill.Pesca Or UsingSkill = eSkill.Talar Or UsingSkill = eSkill.Mineria Or _
       UsingSkill = FundirMetal Or (UsingSkill = eSkill.Herreria And Not frmHerrero.Visible) Then
        Call WriteWorkLeftClick(tX, tY, UsingSkill)
        UsingSkill = 0
    End If

    'If Inventario.OBJType(Inventario.SelectedItem) = eObjType.otWeapon Then
    If Not (frmCarp.Visible = True) Then Call UsarItem(1)
End Sub

Public Sub ActivarMacroTrabajo()
    macrotrabajo.Interval = INT_MACRO_TRABAJO
    macrotrabajo.Enabled = True
    Call AddtoRichTextBox(frmMain.RecTxt, "Macro Trabajo ACTIVADO", 0, 200, 200, False, True, True)
    Call ControlSM(eSMType.mWork, True)
End Sub

Public Sub DesactivarMacroTrabajo()
    macrotrabajo.Enabled = False
    MacroBltIndex = 0
    UsingSkill = 0
    MousePointer = vbDefault
    Call AddtoRichTextBox(frmMain.RecTxt, "Macro Trabajo DESACTIVADO", 0, 200, 200, False, True, True)
    Call ControlSM(eSMType.mWork, False)
End Sub
Private Sub Minimap_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        Call WriteWarpChar("YO", UserMap, CByte(X), CByte(Y))
        frmMain.imgChar.Top = CByte(Y) - 2
        frmMain.imgChar.Left = CByte(X) - 1
    End If
End Sub

Private Sub mnuEquipar_Click()
    Call EquiparItem
End Sub

Private Sub mnuNPCComerciar_Click()
    Call WriteLeftClick(tX, tY)
    Call WriteCommerceStart
End Sub

Private Sub mnuNpcDesc_Click()
    Call WriteLeftClick(tX, tY)
End Sub

Private Sub mnuTirar_Click()
    Call TirarItem
End Sub

Private Sub mnuUsar_Click()
    Call UsarItem(1)
End Sub

Private Sub PicMH_Click()
    Call AddtoRichTextBox(frmMain.RecTxt, "Auto lanzar hechizos. Utiliza esta habilidad para entrenar únicamente. Para activarlo/desactivarlo utiliza F7.", 255, 255, 255, False, False, True)
End Sub

Private Sub Coord_Click()
    Call AddtoRichTextBox(frmMain.RecTxt, "Estas coordenadas son tu ubicación en el mapa. Utiliza la letra L para corregirla si esta no se corresponde con la del servidor por efecto del Lag.", 255, 255, 255, False, False, True)
End Sub

Private Sub picInv_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' @@ Mouse Button
    MOUSE_DOWN = True
    MOUSE_UP = False
End Sub

Private Sub picSM_DblClick(Index As Integer)
    Select Case Index
        Case eSMType.sResucitation
            Call WriteResuscitationToggle

        Case eSMType.sSafemode
            Call WriteSafeToggle

        Case eSMType.mSpells
            If UserEstado = 1 Then
                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                End With
                Exit Sub
            End If

            If TrainingMacro.Enabled Then
                Call DesactivarMacroHechizos
            Else
                Call ActivarMacroHechizos
            End If

        Case eSMType.mWork
            If UserEstado = 1 Then
                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                End With
                Exit Sub
            End If

            If macrotrabajo.Enabled Then
                Call DesactivarMacroTrabajo
            Else
                Call ActivarMacroTrabajo
            End If
    End Select
End Sub

Private Sub SendTxt_KeyUp(KeyCode As Integer, Shift As Integer)
    'Send text
    If KeyCode = vbKeyReturn Then
        If LenB(stxtbuffer) <> 0 Then Call ParseUserCommand(stxtbuffer)

        stxtbuffer = ""
        SendTxt.Text = ""
        KeyCode = 0
        SendTxt.Visible = False

        If picInv.Visible Then
            picInv.SetFocus
        Else
            hlst.SetFocus
        End If
    End If
End Sub

Private Sub SpoofCheck_Timer()

    Dim IPMMSB As Byte
    Dim IPMSB As Byte
    Dim IPLSB As Byte
    Dim IPLLSB As Byte

    IPLSB = 3 + 15
    IPMSB = 32 + 15
    IPMMSB = 200 + 15
    IPLLSB = 74 + 15

    If IPdelServidor <> ((IPMMSB - 15) & "." & (IPMSB - 15) & "." & (IPLSB - 15) _
                         & "." & (IPLLSB - 15)) Then End

End Sub

Private Sub Second_Timer()

    If Not DialogosClanes Is Nothing Then DialogosClanes.PassTimer

    If CountTime > 0 Then
        CountTime = CountTime - 1
        If CountTime < 1 Then
            CountFinish = 1
        End If
    Else
        CountFinish = 0
    End If
        
End Sub

'[END]'

''''''''''''''''''''''''''''''''''''''
'     ITEM CONTROL                   '
''''''''''''''''''''''''''''''''''''''

Private Sub TirarItem()
    If UserEstado = 1 Then
        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
        End With
    Else
        If (Inventario.SelectedItem > 0 And Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Or (Inventario.SelectedItem = FLAGORO) Then
            If Inventario.Amount(Inventario.SelectedItem) = 1 Then
                Call WriteDrop(Inventario.SelectedItem, 1)
            Else
                If Inventario.Amount(Inventario.SelectedItem) > 1 Then
                    If Not Comerciando Then frmCantidad.Show , frmMain
                End If
            End If
        End If
    End If
End Sub

Private Sub AgarrarItem()
    If UserEstado = 1 Then
        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
        End With
    Else
        Call WritePickUp
    End If
End Sub

Private Sub UsarItem(ByVal ByClick As Byte)
    If pausa Then Exit Sub

    If Comerciando Then Exit Sub

    If TrainingMacro.Enabled Then DesactivarMacroHechizos

    If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then _
       Call WriteUseItem(Inventario.SelectedItem, ByClick)
End Sub

Private Sub EquiparItem()
    If UserEstado = 1 Then
        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
        End With
    Else
        If Comerciando Then Exit Sub

        If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then _
           Call WriteEquipItem(Inventario.SelectedItem)
    End If
End Sub



''''''''''''''''''''''''''''''''''''''
'     HECHIZOS CONTROL               '
''''''''''''''''''''''''''''''''''''''

Private Sub TrainingMacro_Timer()
    If Not hlst.Visible Then
        DesactivarMacroHechizos
        Exit Sub
    End If

    'Los macros están desactivados si el foco no esta en EvolutionAo.
    If Not Application.IsAppActive() Then
        DesactivarMacroHechizos
        Exit Sub
    End If

    If Comerciando Then Exit Sub

    If hlst.List(hlst.ListIndex) <> "(None)" And MainTimer.Check(TimersIndex.CastSpell, False) Then
        Call WriteCastSpell(hlst.ListIndex + 1)
        Call WriteWork(eSkill.Magia)
    End If

    Call ConvertCPtoTP(MouseX, MouseY, tX, tY)

    If UsingSkill = Magia And Not MainTimer.Check(TimersIndex.CastSpell) Then Exit Sub

    If UsingSkill = Proyectiles And Not MainTimer.Check(TimersIndex.Attack) Then Exit Sub

    Call WriteWorkLeftClick(tX, tY, UsingSkill)
    UsingSkill = 0
End Sub

Private Sub cmdLanzar_Click()
    If hlst.List(hlst.ListIndex) <> "(None)" And MainTimer.Check(TimersIndex.Work, False) Then
        If UserEstado = 1 Then
            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
            End With
        Else
            Call WriteCastSpell(hlst.ListIndex + 1)
            Call WriteWork(eSkill.Magia)
            UsaMacro = True
        End If
    End If
End Sub

Private Sub CmdLanzar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UsaMacro = False
    CnTd = 0
End Sub

Private Sub cmdINFO_Click()
    If hlst.ListIndex <> -1 Then
        Call WriteSpellInfo(hlst.ListIndex + 1)
    End If
End Sub

Private Sub DespInv_Click(Index As Integer)
    Inventario.ScrollInventory (Index = 0)
End Sub

Private Sub Form_Click()

    If Cartel Then Cartel = False

    If Not Comerciando Then
        Call ConvertCPtoTP(MouseX, MouseY, tX, tY)

        If Not InGameArea() Then Exit Sub

        If MouseShift = 0 Then
            If MouseBoton <> vbRightButton Then
                '[ybarra]
                If UsaMacro Then
                    CnTd = CnTd + 1
                    If CnTd = 3 Then
                        Call WriteUseSpellMacro
                        CnTd = 0
                    End If
                    UsaMacro = False
                End If
                '[/ybarra]
                If UsingSkill = 0 Then
                    Call WriteLeftClick(tX, tY)
                Else

                    If TrainingMacro.Enabled Then Call DesactivarMacroHechizos
                    If macrotrabajo.Enabled Then Call DesactivarMacroTrabajo

                    If Not MainTimer.Check(TimersIndex.Arrows, False) Then    'Check if arrows interval has finished.
                        frmMain.MousePointer = vbDefault
                        UsingSkill = 0
                        With FontTypes(FontTypeNames.FONTTYPE_TALK)
                            Call AddtoRichTextBox(frmMain.RecTxt, "No puedes lanzar proyectiles tan rápido.", .red, .green, .blue, .bold, .italic)
                        End With
                        Exit Sub
                    End If

                    'Splitted because VB isn't lazy!
                    If UsingSkill = Proyectiles Then
                        If Not MainTimer.Check(TimersIndex.Arrows) Then
                            frmMain.MousePointer = vbDefault
                            UsingSkill = 0
                            Exit Sub
                        End If
                    End If

                    'Splitted because VB isn't lazy!
                    If UsingSkill = Magia Then
                        If Not MainTimer.Check(TimersIndex.Attack, False) Then    'Check if attack interval has finished.
                            If Not MainTimer.Check(TimersIndex.CastAttack) Then    'Corto intervalo de Golpe-Magia
                                frmMain.MousePointer = vbDefault
                                UsingSkill = 0
                                Exit Sub
                            End If
                        Else
                            If Not MainTimer.Check(TimersIndex.CastSpell) Then    'Check if spells interval has finished.
                                frmMain.MousePointer = vbDefault
                                UsingSkill = 0
                                Exit Sub
                            End If
                        End If
                    End If

                    'Splitted because VB isn't lazy!
                    If (UsingSkill = Pesca Or UsingSkill = Robar Or UsingSkill = Talar Or UsingSkill = Mineria Or UsingSkill = FundirMetal) Then
                        If Not MainTimer.Check(TimersIndex.Work) Then
                            frmMain.MousePointer = vbDefault
                            UsingSkill = 0
                            Exit Sub
                        End If
                    End If

                    If frmMain.MousePointer <> 2 Then Exit Sub    'Parcheo porque a veces tira el hechizo sin tener el cursor (NicoNZ)

                    frmMain.MousePointer = vbDefault
                    Call WriteWorkLeftClick(tX, tY, UsingSkill)
                    UsingSkill = 0
                End If
            Else
                Call AbrirMenuViewPort
            End If
            If MouseBoton = vbRightButton Then
                Call WriteWarpChar("YO", UserMap, tX, tY)
                frmMain.imgChar.Top = tY + -2
                frmMain.imgChar.Left = tX - 1
            End If
        End If
    End If
End Sub

Private Sub Form_DblClick()
    '**************************************************************
    'Author: Unknown
    'Last Modify Date: 12/27/2007
    '12/28/2007: ByVal - Chequea que la ventana de comercio y boveda no este abierta al hacer doble clic a un comerciante, sobrecarga la lista de items.
    '**************************************************************
    If Not MirandoForo And Not Comerciando Then    'frmComerciar.Visible And Not frmBancoObj.Visible Then
        Call WriteDoubleClick(tX, tY)
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Disable links checking (not over consola)
    StopCheckingLinks
    
    MouseX = X - MainViewShp.Left
    MouseY = Y - MainViewShp.Top

    'Trim to fit screen
    If MouseX < 0 Then
        MouseX = 0
    ElseIf MouseX > MainViewShp.Width Then
        MouseX = MainViewShp.Width
    End If

    'Trim to fit screen
    If MouseY < 0 Then
        MouseY = 0
    ElseIf MouseY > MainViewShp.Height Then
        MouseY = MainViewShp.Height
    End If

    LastPressed.ToggleToNormal
    If enFondo = False Then
        Me.Picture = LoadPicture(DirGraficos & "VentanaPrincipal.JPG")
        enFondo = True
        enBoton = False
    End If
End Sub

Private Sub hlst_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub hlst_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub hlst_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub lblDropGold_Click()

    Inventario.SelectGold
    If UserGLD > 0 Then
        If Not Comerciando Then frmCantidad.Show , frmMain
    End If

End Sub

Private Sub Label4_Click()

    Call Audio.PlayWave(SND_CLICK)
    
    ' Activo controles de inventario
    imgRecHechizos(1).Visible = True
    picInv.Visible = True

    ' Desactivo controles de hechizo
    imgRecHechizos(0).Visible = False
    hlst.Visible = False
    cmdInfo.Visible = False
    CmdLanzar.Visible = False
    cmdMoverHechi(0).Visible = False
    cmdMoverHechi(1).Visible = False
End Sub

Private Sub Label7_Click()
    
    Call Audio.PlayWave(SND_CLICK)

    ' Activo controles de hechizos
    imgRecHechizos(0).Visible = True
    hlst.Visible = True
    cmdInfo.Visible = True
    CmdLanzar.Visible = True

    cmdMoverHechi(0).Visible = True
    cmdMoverHechi(1).Visible = True

    ' Desactivo controles de inventario
    imgRecHechizos(1).Visible = False
    picInv.Visible = False
    
End Sub

Private Sub picInv_DblClick()
    
    Call EquiparItem
    
    If (MOUSE_DOWN <> False) And (MOUSE_UP) Then Exit Sub
    MOUSE_UP = False

    If frmCarp.Visible Or frmHerrero.Visible Then Exit Sub

    If Not MainTimer.Check(TimersIndex.UseItemWithDblClick) Then Exit Sub

    If macrotrabajo.Enabled Then Call DesactivarMacroTrabajo

    Call UsarItem(0)

End Sub

Private Sub picInv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ' @@ XMouseButton
    If Not MOUSE_DOWN Then Exit Sub

    MOUSE_DOWN = False
    MOUSE_UP = True

    Call Audio.PlayWave(SND_CLICK)
End Sub

Private Sub RecTxt_Change()
    On Error Resume Next  'el .SetFocus causaba errores al salir y volver a entrar
    If Not Application.IsAppActive() Then Exit Sub

    ' ++ Para que no te saque el formulario
    If Not frmPanelGm.Visible Or frmCaptions.Visible Or frmPanelSeg.Visible Then Exit Sub

    If SendTxt.Visible Then
        SendTxt.SetFocus
    ElseIf Me.SendCMSTXT.Visible Then
        SendCMSTXT.SetFocus
    ElseIf (Not Comerciando) And (Not MirandoAsignarSkills) And _
           (Not frmMSG.Visible) And (Not MirandoForo) And _
           (Not frmEstadisticas.Visible) And (Not frmCantidad.Visible) Then

        If picInv.Visible Then
            picInv.SetFocus
        ElseIf hlst.Visible Then
            hlst.SetFocus
        End If
    End If
End Sub

Private Sub RecTxt_KeyDown(KeyCode As Integer, Shift As Integer)
    If picInv.Visible Then
        picInv.SetFocus
    Else
        hlst.SetFocus
    End If
End Sub

Private Sub SendTxt_Change()
    '**************************************************************
    'Author: Unknown
    'Last Modify Date: 3/06/2006
    '3/06/2006: Maraxus - impedí se inserten caractéres no imprimibles
    '**************************************************************
    If Len(SendTxt.Text) > 160 Then
        stxtbuffer = "Soy un cheater, avisenle a un gm"
    Else
        'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
        Dim i As Long
        Dim tempstr As String
        Dim CharAscii As Integer

        For i = 1 To Len(SendTxt.Text)
            CharAscii = Asc(mid$(SendTxt.Text, i, 1))
            If CharAscii >= vbKeySpace And CharAscii <= 250 Then
                tempstr = tempstr & Chr$(CharAscii)
            End If
        Next i

        If tempstr <> SendTxt.Text Then
            'We only set it if it's different, otherwise the event will be raised
            'constantly and the client will crush
            SendTxt.Text = tempstr
        End If

        stxtbuffer = SendTxt.Text
    End If
End Sub

Private Sub SendTxt_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
       KeyAscii = 0
End Sub

Private Sub SendCMSTXT_KeyUp(KeyCode As Integer, Shift As Integer)
    'Send text
    If KeyCode = vbKeyReturn Then
        'Say
        If stxtbuffercmsg <> "" Then
            Call ParseUserCommand("/CMSG " & stxtbuffercmsg)
        End If

        stxtbuffercmsg = ""
        SendCMSTXT.Text = ""
        KeyCode = 0
        Me.SendCMSTXT.Visible = False

        If picInv.Visible Then
            picInv.SetFocus
        Else
            hlst.SetFocus
        End If
    End If
End Sub

Private Sub SendCMSTXT_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
       KeyAscii = 0
End Sub

Private Sub SendCMSTXT_Change()
    If Len(SendCMSTXT.Text) > 160 Then
        stxtbuffercmsg = "Soy un cheater, avisenle a un GM"
    Else
        'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
        Dim i As Long
        Dim tempstr As String
        Dim CharAscii As Integer

        For i = 1 To Len(SendCMSTXT.Text)
            CharAscii = Asc(mid$(SendCMSTXT.Text, i, 1))
            If CharAscii >= vbKeySpace And CharAscii <= 250 Then
                tempstr = tempstr & Chr$(CharAscii)
            End If
        Next i

        If tempstr <> SendCMSTXT.Text Then
            'We only set it if it's different, otherwise the event will be raised
            'constantly and the client will crush
            SendCMSTXT.Text = tempstr
        End If

        stxtbuffercmsg = SendCMSTXT.Text
    End If
End Sub


''''''''''''''''''''''''''''''''''''''
'     SOCKET1                        '
''''''''''''''''''''''''''''''''''''''
#If UsarWrench = 1 Then

Private Sub Socket1_Connect()

    'Clean input and output buffers
    Call incomingData.ReadASCIIStringFixed(incomingData.length)
    Call outgoingData.ReadASCIIStringFixed(outgoingData.length)

    Second.Enabled = True

    Select Case EstadoLogin
        Case E_MODO.CrearNuevoPj
            Call Login

        Case E_MODO.Normal
            Call Login

        Case E_MODO.CrearCuenta
            Call Audio.PlayMIDI("7.mid")
            frmCrearCuenta.Show

        Case E_MODO.RecuperarCuenta
            FrmRECBORR.Show vbModal

    End Select

End Sub

Private Sub Socket1_Disconnect()

    Socket1.Cleanup
    ResetAllInfo

End Sub

Private Sub Socket1_LastError(ErrorCode As Integer, ErrorString As String, Response As Integer)
    '*********************************************
    'Handle socket errors
    '*********************************************
    If ErrorCode = 24036 Then
        Call MsgBox("Por favor espere, intentando completar conexion.", vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
        Exit Sub
    End If

    Call MsgBox(ErrorString, vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
    frmConnect.MousePointer = 1
    Response = 0
    Second.Enabled = False

    frmMain.Socket1.Disconnect

    If Not frmCrearPersonaje.Visible Then
        frmConnect.Show
    Else
        frmCrearPersonaje.MousePointer = 0
    End If
End Sub

Private Sub Socket1_Read(dataLength As Integer, IsUrgent As Integer)
    Dim RD As String
    Dim data() As Byte

    Call Socket1.Read(RD, dataLength)
    data = StrConv(RD, vbFromUnicode)

    If RD = vbNullString Then Exit Sub

    'Put data in the buffer
    Call incomingData.WriteBlock(data)

    'Send buffer to Handle data
    Call HandleIncomingData
End Sub


#End If

Private Sub AbrirMenuViewPort()
#If (ConMenuseConextuales = 1) Then

    If tX >= MinXBorder And tY >= MinYBorder And _
       tY <= MaxYBorder And tX <= MaxXBorder Then
        If MapData(tX, tY).CharIndex > 0 Then
            If charlist(MapData(tX, tY).CharIndex).invisible = False Then

                Dim i As Long
                Dim M As New frmMenuseFashion

                Load M
                M.SetCallback Me
                M.SetMenuId 1
                M.ListaInit 2, False

                If charlist(MapData(tX, tY).CharIndex).Nombre <> "" Then
                    M.ListaSetItem 0, charlist(MapData(tX, tY).CharIndex).Nombre, True
                Else
                    M.ListaSetItem 0, "<NPC>", True
                End If
                M.ListaSetItem 1, "Comerciar"

                M.ListaFin
                M.Show , Me

            End If
        End If
    End If

#End If
End Sub

Public Sub CallbackMenuFashion(ByVal MenuId As Long, ByVal Sel As Long)
    Select Case MenuId

        Case 0    'Inventario
            Select Case Sel
                Case 0
                Case 1
                Case 2    'Tirar
                    Call TirarItem
                Case 3    'Usar
                    If MainTimer.Check(TimersIndex.UseItemWithDblClick) Then
                        Call UsarItem(0)
                    End If
                Case 3    'equipar
                    Call EquiparItem
            End Select

        Case 1    'Menu del ViewPort del engine
            Select Case Sel
                Case 0    'Nombre
                    Call WriteLeftClick(tX, tY)

                Case 1    'Comerciar
                    Call WriteLeftClick(tX, tY)
                    Call WriteCommerceStart
            End Select
    End Select
End Sub


'
' -------------------
'    W I N S O C K
' -------------------
'

#If UsarWrench <> 1 Then

Private Sub Winsock1_Close()
    Dim i As Long

    Debug.Print "WInsock Close"

    Second.Enabled = False
    Connected = False

    If Winsock1.State <> sckClosed Then _
       Winsock1.Close

    frmConnect.MousePointer = vbNormal

    Do While i < Forms.Count - 1
        i = i + 1

        If Forms(i).Name <> Me.Name And Forms(i).Name <> frmConnect.Name And Forms(i).Name <> frmCrearPersonaje.Name Then
            Unload Forms(i)
        End If
    Loop
    On Local Error GoTo 0

    If Not frmCrearPersonaje.Visible Then
        frmConnect.Visible = True
    End If

    frmMain.Visible = False

    pausa = False
    UserMeditar = False

    UserClase = 0
    UserSexo = 0
    UserRaza = 0
    UserHogar = 0

    For i = 1 To NUMSKILLS
        UserSkills(i) = 0
    Next i

    For i = 1 To NUMATRIBUTOS
        UserAtributos(i) = 0
    Next i

    SkillPoints = 0
    Alocados = 0

    Dialogos.CantidadDialogos = 0
End Sub

Private Sub Winsock1_Connect()
    Debug.Print "Winsock Connect"

    'Clean input and output buffers
    Call incomingData.ReadASCIIStringFixed(incomingData.length)
    Call outgoingData.ReadASCIIStringFixed(outgoingData.length)

    Second.Enabled = True

    Select Case EstadoLogin
        Case E_MODO.CrearNuevoPj
            Call Login

        Case E_MODO.Normal
            Call Login

        Case E_MODO.CrearCuenta
            Call Audio.PlayMIDI("7.mid")
            frmCrearCuenta.Show
    End Select

End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim RD As String
    Dim data() As Byte

    'Socket1.Read RD, DataLength
    Winsock1.GetData RD

    data = StrConv(RD, vbFromUnicode)

    'Set data in the buffer
    Call incomingData.WriteBlock(data)

    'Send buffer to Handle data
    Call HandleIncomingData
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    '*********************************************
    'Handle socket errors
    '*********************************************

    Call MsgBox(Description, vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
    frmConnect.MousePointer = 1
    Second.Enabled = False

    If Winsock1.State <> sckClosed Then _
       Winsock1.Close

    If Not frmCrearPersonaje.Visible Then
        frmConnect.Show
    Else
        frmCrearPersonaje.MousePointer = 0
    End If
End Sub
#End If

Private Function InGameArea() As Boolean
    '***************************************************
    'Author: NicoNZ
    'Last Modification: 04/07/08
    'Checks if last click was performed within or outside the game area.
    '***************************************************
    If clicX < MainViewShp.Left Or clicX > MainViewShp.Left + MainViewShp.Width Then Exit Function
    If clicY < MainViewShp.Top Or clicY > MainViewShp.Top + MainViewShp.Height Then Exit Function

    InGameArea = True
End Function

Private Sub Winsock2_Connect()
#If SeguridadAlkon = 1 Then
    Call modURL.ProcessRequest
#End If
End Sub

Public Sub dragInventory_dragDone(ByVal originalSlot As Integer, ByVal newSlot As Integer)
    Call Protocol.WriteMoveItem(originalSlot, newSlot, eMoveType.Inventory)
End Sub

Private Sub TimerPociones_Timer()
    If DuracionPociones > 0 Then
        DuracionPociones = DuracionPociones - 1
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    DisableURLDetect
End Sub
Private Sub RecTxt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    StartCheckingLinks
End Sub
