Attribute VB_Name = "m_Recordar"
Option Explicit

'Contraseña de la encriptacion.
Private Const UserKey As String = "ClaveEncrypt1449"    'Cambiala alv despues

Private Type tRecu
    Password As String
    Nick As String
End Type

Public Recu() As tRecu
Private MaxRecu As Byte

Public Sub LoadRecup()

    Dim Leer As clsIniManager
    Set Leer = New clsIniManager

    'Cargamos el archivo que guarda la contraseña.
    Call Leer.Initialize(App.path & "\Librerias DLLS\DXDIALOG.OCX")

    MaxRecu = Val(Leer.GetValue("INIT", "MAX"))

    If MaxRecu <> 0 Then

        ReDim Recu(1 To MaxRecu) As tRecu
        Dim j As Long

        For j = 1 To MaxRecu
            Recu(j).Nick = DesEncryptString(Leer.GetValue("INIT", "NICK" & j))
            Recu(j).Password = DesEncryptString(Leer.GetValue("INIT", "PASS" & j))
        Next j

    End If

    Set Leer = Nothing

End Sub

Public Function NickExiste(ByVal Nombre As String) As Byte

    Nombre = UCase$(Nombre)

    Dim LoopC As Long

    For LoopC = 1 To MaxRecu
        If StrComp(Recu(LoopC).Nick, Nombre) = 0 Then
            NickExiste = LoopC
            Exit Function
        End If
    Next LoopC

    NickExiste = 0

End Function

Public Sub SaveRecu(ByVal Nombre As String, ByVal Password As String)

    Nombre = UCase$(Nombre)
    Password = UCase$(Password)

    Dim Leer As clsIniManager

    If NickExiste(Nombre) Then

        Dim i As Long

        For i = 1 To MaxRecu
            If StrComp(Recu(i).Nick, Nombre) = 0 Then
                If Password <> Recu(i).Password Then
                    Recu(i).Password = Password

                    Set Leer = New clsIniManager

                    Call Leer.Initialize(App.path & "\Librerias DLLS\DXDIALOG.OCX")
                    Call Leer.ChangeValue("INIT", "PASS" & i, EncryptString(Password))

                    Call Leer.DumpFile(App.path & "\Librerias DLLS\DXDIALOG.OCX")
                    Set Leer = Nothing

                End If
                Exit Sub
            End If
        Next i

    Else

        MaxRecu = MaxRecu + 1
        ReDim Preserve Recu(1 To MaxRecu) As tRecu

        Recu(MaxRecu).Nick = Nombre
        Recu(MaxRecu).Password = Password

        Set Leer = New clsIniManager
        Call Leer.Initialize(App.path & "\Librerias DLLS\DXDIALOG.OCX")

        Call Leer.ChangeValue("INIT", "MAX", MaxRecu)
        Call Leer.ChangeValue("INIT", "NICK" & MaxRecu, EncryptString(Nombre))
        Call Leer.ChangeValue("INIT", "PASS" & MaxRecu, EncryptString(Password))

        Call Leer.DumpFile(App.path & "\Librerias DLLS\DXDIALOG.OCX")
        Set Leer = Nothing

    End If

End Sub

Private Function EncryptString(ByVal Text As String) As String

    Dim Temp As Integer
    Dim i As Long
    Dim j As Integer
    Dim N As Integer
    Dim rtn As String
    Dim LenText As String

    N = Len(UserKey)

    ReDim UserKeyASCIIS(1 To N)

    For i = 1 To N
        UserKeyASCIIS(i) = Asc(mid$(UserKey, i, 1))
    Next i

    LenText = Len(Text)
    ReDim TextASCIIS(LenText) As Integer

    For i = 1 To LenText
        TextASCIIS(i) = Asc(mid$(Text, i, 1))
    Next i

    For i = 1 To LenText
        j = IIf(j + 1 >= N, 1, j + 1)
        Temp = TextASCIIS(i) + UserKeyASCIIS(j)
        If Temp > 255 Then
            Temp = Temp - 255
        End If
        rtn = rtn + Chr$(Temp)
    Next

    EncryptString = rtn

End Function

Private Function DesEncryptString(ByVal Text As String) As String

    Dim Temp As Integer
    Dim i As Long
    Dim j As Integer
    Dim N As Integer
    Dim rtn As String
    Dim LenText As String

    N = Len(UserKey)

    ReDim UserKeyASCIIS(1 To N)

    For i = 1 To N
        UserKeyASCIIS(i) = Asc(mid$(UserKey, i, 1))
    Next i

    LenText = Len(Text)
    ReDim TextASCIIS(LenText) As Integer

    For i = 1 To LenText
        TextASCIIS(i) = Asc(mid$(Text, i, 1))
    Next i

    For i = 1 To LenText
        j = IIf(j + 1 >= N, 1, j + 1)
        Temp = TextASCIIS(i) - UserKeyASCIIS(j)
        If Temp < 0 Then
            Temp = Temp + 255
        End If
        rtn = rtn + Chr$(Temp)
    Next i

    DesEncryptString = rtn

End Function
