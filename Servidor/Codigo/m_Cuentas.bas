Attribute VB_Name = "m_Cuentas"

Option Explicit

Public Type AccountUser
    Name As String
    Body As Integer
    Head As Integer
    Weapon As Integer
    Shield As Integer
    Helmet As Integer
    Class As Byte
    Race As Byte
    map As Integer
    level As Byte
    Gold As Long
    Estatus As Byte
    Genero As Byte
End Type

Public Sub CreateNewAccount(ByVal UserIndex As Integer, ByVal UserAccount As String, ByVal Password As String, ByVal Email As String, ByVal Pin As String)

    If LenB(UserAccount) < 1 Then Exit Sub 'Tiene que tener un caractere
    If Len(UserAccount) > 20 Then Exit Sub 'Tiene que tener menos de 20 carácteres
    If LenB(Password) < 1 Then Exit Sub

    If Not AsciiValidos(UserAccount) Then Exit Sub

    If CuentaExiste(UserAccount) Then
        Call WriteErrorMsg(UserIndex, "Ya existe la cuenta.")
        Exit Sub
    End If

    Call SaveNewAccountCharfile(UserAccount, Password, Email, Pin)
    Call ConnectAccount(UserIndex, UserAccount, Password)

End Sub

Public Sub ConnectAccount(ByVal UserIndex As Integer, ByVal UserAccount As String, ByVal Password As String)

    If LenB(UserAccount) < 1 Then Exit Sub 'Tiene que tener un caractere
    If Len(UserAccount) > 20 Then Exit Sub 'Tiene que tener menos de 20 carácteres
    If LenB(Password) < 1 Then Exit Sub

    If Not AsciiValidos(UserAccount) Then Exit Sub

    If Not CuentaExiste(UserAccount) Then
        Call WriteErrorMsg(UserIndex, "No existe la cuenta.")
        Exit Sub
    End If

    Dim Leer As clsIniManager
    Set Leer = New clsIniManager
    Call Leer.Initialize(AccountPath & UserAccount & ".acc")

    'Es la contraseña correcta?
    If UCase$(Password) = UCase$(Leer.GetValue("INIT", "PASSWORD")) Then
        If val(Leer.GetValue("INIT", "ESTADO")) <> 1 Then
            UserList(UserIndex).Account = UserAccount
            UserList(UserIndex).PasswordLogged = UCase$(Password)
            Call LoginAccountCharfile(UserIndex, UserAccount)
            Call WriteCaptchaCode(UserIndex)
            
            Call Leer.ChangeValue("INIT", "FECHA_ULTIMA_VISITA", Now)
            Call Leer.DumpFile(AccountPath & UserAccount & ".acc")
        Else
            Call WriteErrorMsg(UserIndex, "Cuenta deshabilitada.")
        End If
    Else
        Call WriteErrorMsg(UserIndex, "Password incorrecto.")
    End If

    Set Leer = Nothing

End Sub

Public Sub SaveNewAccountCharfile(ByVal UserName As String, ByVal Password As String, ByVal Email As String, ByVal Pin As String)

    Dim Manager As clsIniManager
    Set Manager = New clsIniManager

    Dim AccountFile As String
    AccountFile = AccountPath & UCase$(UserName) & ".acc"

    With Manager
        Call .ChangeValue("INIT", "ESTADO", "0")
        Call .ChangeValue("INIT", "PASSWORD", Password)
        Call .ChangeValue("INIT", "EMAIL", Email)
        Call .ChangeValue("INIT", "PIN", Pin)
        
        Call .ChangeValue("INIT", "FECHA_CREACION", Now)
        Call .ChangeValue("INIT", "FECHA_ULTIMA_VISITA", Now)
        Call .DumpFile(AccountFile)
    End With

    Set Manager = Nothing

End Sub

Public Sub LoginAccountCharfile(ByVal UserIndex As Integer, ByVal UserAccount As String)

    Dim Leer As clsIniManager
    Set Leer = New clsIniManager
    Call Leer.Initialize(AccountPath & UserAccount & ".acc")

    Dim NumberOfCharacters As Byte
    NumberOfCharacters = val(Leer.GetValue("INIT", "CANT_PJS"))

    Dim Characters(1 To 8) As AccountUser

    If NumberOfCharacters <> 0 Then

        Dim i As Long
        Dim Priv As Byte
        Dim CurrentCharacter As String

        For i = 1 To NumberOfCharacters
            CurrentCharacter = Leer.GetValue("PJS", "PJ" & i)

            If PersonajeExiste(CurrentCharacter) Then
                Priv = UserDarPrivilegioLevel(CurrentCharacter)

                With Characters(i)

                    .Name = CurrentCharacter

                    If val(GetVar(CharPath & CurrentCharacter & ".chr", "FLAGS", "Muerto")) = 0 Then
                        .Body = val(GetVar(CharPath & CurrentCharacter & ".chr", "INIT", "Body"))
                        .Head = val(GetVar(CharPath & CurrentCharacter & ".chr", "INIT", "Head"))
                    Else
                        .Body = 8
                        .Head = 500
                    End If

                    .Weapon = val(GetVar(CharPath & CurrentCharacter & ".chr", "INIT", "Arma"))
                    .Shield = val(GetVar(CharPath & CurrentCharacter & ".chr", "INIT", "Escudo"))
                    .Helmet = val(GetVar(CharPath & CurrentCharacter & ".chr", "INIT", "Casco"))
                    .Class = val(GetVar(CharPath & CurrentCharacter & ".chr", "INIT", "Clase"))
                    .Race = val(GetVar(CharPath & CurrentCharacter & ".chr", "INIT", "Raza"))
                    .Genero = val(GetVar(CharPath & CurrentCharacter & ".chr", "INIT", "Genero"))
                    .map = CInt(ReadField(1, GetVar(CharPath & CurrentCharacter & ".chr", "INIT", "Position"), 45))
                    .level = val(GetVar(CharPath & CurrentCharacter & ".chr", "STATS", "ELV"))
                    .Gold = val(GetVar(CharPath & CurrentCharacter & ".chr", "STATS", "GLD"))

                    If BodyIsBoat(.Body) Then
                        Dim ArmourEqpSlot As Byte
                        ArmourEqpSlot = CByte(GetVar(CharPath & CurrentCharacter & ".chr", "Inventory", "ArmourEqpSlot"))

                        If ArmourEqpSlot > 0 Then
                            Dim ArmourObjIndex As Integer
                            ArmourObjIndex = CInt(ReadField(1, GetVar(CharPath & CurrentCharacter & ".chr", "Inventory", "Obj" & ArmourEqpSlot), 45))

                            If ArmourObjIndex > 0 Then
                                .Body = ObjData(ArmourObjIndex).Ropaje
                            End If
                        Else
                            .Body = CuerpoDesnudo(.Genero, .Race)
                        End If
                    End If

                    If Priv <> PlayerType.User Then
                        Characters(i).Estatus = Priv
                    Else
                        If val(GetVar(CharPath & CurrentCharacter & ".chr", "REP", "Promedio")) < 0 Then
                            Characters(i).Estatus = 255
                        Else
                            Characters(i).Estatus = 0
                        End If
                    End If

                End With

            End If

        Next i

    End If

    Call WriteUserAccountLogged(UserIndex, UserAccount, NumberOfCharacters, Characters)
    Set Leer = Nothing

End Sub

Public Sub SaveUserToAccountCharfile(ByVal UserName As String, ByVal UserAccount As String)

    Dim AccountCharfile As String
    AccountCharfile = AccountPath & UserAccount & ".acc"

    If FileExist(AccountCharfile) Then
        Dim Leer As clsIniManager
        Set Leer = New clsIniManager
        Call Leer.Initialize(AccountCharfile)
            
        Dim CantPjs As Byte
        CantPjs = val(Leer.GetValue("INIT", "CANT_PJS"))
        CantPjs = CantPjs + 1

        If CantPjs < 9 Then
            Call Leer.ChangeValue("INIT", "CANT_PJS", CantPjs)
            Call Leer.ChangeValue("PJS", "PJ" & CantPjs, UserName)
            Call Leer.DumpFile(AccountCharfile)
        Else
            'Call LogError("Error in SaveUserToAccountCharfile. Se intento crear mas de 8 personajes. UserName: " & UserName)
        End If
        
        Set Leer = Nothing
    Else
        Call LogError("Error in SaveUserToAccountCharfile. Cuenta inexistente de " & UserAccount)
    End If

End Sub

Public Function GenerateRandomKey(ByVal UserAccount As String) As Integer

    On Error GoTo Erroraso

    GenerateRandomKey = RandomNumber(5000, 16000) + RandomNumber(5000, 16000)
    Call WriteVar(AccountPath & UserAccount & ".acc", "INIT", "PASSWORD", GenerateRandomKey)

    Exit Function

Erroraso:

    GenerateRandomKey = 32000

End Function

Private Function CuerpoDesnudo(ByVal Genero As Byte, ByVal raza As Byte) As Integer

    Select Case Genero
        Case eGenero.Hombre
            Select Case raza
                Case eRaza.Humano
                    CuerpoDesnudo = 21
                Case eRaza.Drow
                    CuerpoDesnudo = 32
                Case eRaza.Elfo
                    CuerpoDesnudo = 210
                Case eRaza.Gnomo
                    CuerpoDesnudo = 222
                Case eRaza.Enano
                    CuerpoDesnudo = 53
            End Select
        Case eGenero.Mujer
            Select Case raza
                Case eRaza.Humano
                    CuerpoDesnudo = 39
                Case eRaza.Drow
                    CuerpoDesnudo = 40
                Case eRaza.Elfo
                    CuerpoDesnudo = 259
                Case eRaza.Gnomo
                    CuerpoDesnudo = 260
                Case eRaza.Enano
                    CuerpoDesnudo = 60
            End Select
    End Select

End Function
