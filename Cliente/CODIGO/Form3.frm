VERSION 5.00
Begin VB.Form frmCaptions 
   BackColor       =   &H8000000A&
   Caption         =   "Form3"
   ClientHeight    =   3465
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   48
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   ScaleHeight     =   3465
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808080&
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   2760
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.ListBox List1 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   480
      TabIndex        =   1
      Top             =   600
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      Caption         =   "CAPTIONS DE "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "frmCaptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Esta función Api devuelve un valor  Boolean indicando si la ventana es una ventana visible
Private Declare Function IsWindowVisible _
                          Lib "user32" ( _
                              ByVal hWnd As Long) As Long

'Esta función retorna el número de caracteres del caption de la ventana
Private Declare Function GetWindowTextLength _
                          Lib "user32" _
                              Alias "GetWindowTextLengthA" ( _
                              ByVal hWnd As Long) As Long

'Esta devuelve el texto. Se le pasa el hwnd de la ventana, un buffer donde se
'almacenará el texto devuelto, y el Lenght de la cadena en el último parámetro
'que obtuvimos con el Api GetWindowTextLength
Private Declare Function GetWindowText _
                          Lib "user32" _
                              Alias "GetWindowTextA" ( _
                              ByVal hWnd As Long, _
                              ByVal lpString As String, _
                              ByVal cch As Long) As Long

'Esta es la función Api que busca las ventanas y retorna su handle o Hwnd
Private Declare Function GetWindow _
                          Lib "user32" ( _
                              ByVal hWnd As Long, _
                              ByVal wFlag As Long) As Long

'Constantes para buscar las ventanas mediante el Api GetWindow
Private Const GW_HWNDFIRST = 0&
Private Const GW_HWNDNEXT = 2&
Private Const GW_CHILD = 5&
Public CANTv As Byte

Public Function Listar() As String

    Me.List1.Clear    'Faltaba esto

    Dim buf As Long, handle As Long, titulo As String, lenT As Long, Ret As Long
    'Obtenemos el Hwnd de la primera ventana, usando la constante GW_HWNDFIRST
    handle = GetWindow(hWnd, GW_HWNDFIRST)

    'Este bucle va a recorrer todas las ventanas.
    'cuando GetWindow devielva un 0, es por que no hay mas
    Do While handle <> 0
        'Tenemos que comprobar que la ventana es una de tipo visible
        If IsWindowVisible(handle) Then
            'Obtenemos el número de caracteres de la ventana
            lenT = GetWindowTextLength(handle)
            'si es el número anterior es mayor a 0
            If lenT > 0 Then
                'Creamos un buffer. Este buffer tendrá el tamaño con la variable LenT
                titulo = String$(lenT, 0)
                'Ahora recuperamos el texto de la ventana en el buffer que le enviamos
                'y también debemos pasarle el Hwnd de dicha ventana
                Ret = GetWindowText(handle, titulo, lenT + 1)
                titulo$ = Left$(titulo, Ret)
                'La agregamos al ListBox
                Listar = titulo & "#" & Listar
                CANTv = CANTv + 1
            End If
        End If
        'Buscamos con GetWindow la próxima ventana usando la constante GW_HWNDNEXT
        handle = GetWindow(handle, GW_HWNDNEXT)
    Loop

End Function

