Attribute VB_Name = "Application"
Option Explicit

'Recupera la ventana si no está activa.

Private Declare Function GetActiveWindow Lib "user32" () As Long

'@return Verdadero

Public Function IsAppActive() As Boolean
    IsAppActive = (GetActiveWindow <> 0)
End Function
