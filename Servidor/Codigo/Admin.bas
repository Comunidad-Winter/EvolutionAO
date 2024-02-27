Attribute VB_Name = "Admin"
Option Explicit

Public Type tMotd
    texto As String
    Formato As String
End Type

Public MaxLines As Integer
Public MOTD() As tMotd

Public Type tAPuestas
    Ganancias As Long
    Perdidas As Long
    Jugadas As Long
End Type
Public Apuestas As tAPuestas

Public tInicioServer As Long
Public EstadisticasWeb As New clsEstadisticasIPC

'INTERVALOS
Public SanaIntervaloSinDescansar As Integer
Public StaminaIntervaloSinDescansar As Integer
Public SanaIntervaloDescansar As Integer
Public StaminaIntervaloDescansar As Integer
Public IntervaloSed As Integer
Public IntervaloHambre As Integer
Public IntervaloVeneno As Integer
Public IntervaloParalizado As Integer
Public IntervaloInvisible As Integer
Public IntervaloFrio As Integer
Public IntervaloWavFx As Integer
Public IntervaloLanzaHechizo As Integer
Public IntervaloNPCPuedeAtacar As Integer
Public IntervaloNPCAI As Integer
Public IntervaloInvocacion As Integer
Public IntervaloOculto As Integer    '[Nacho]
Public IntervaloUserPuedeAtacar As Long
Public IntervaloGolpeUsar As Long
Public IntervaloMagiaGolpe As Long
Public IntervaloGolpeMagia As Long
Public IntervaloUserPuedeCastear As Long
Public IntervaloUserPuedeTrabajar As Long
Public IntervaloParaConexion As Long
Public IntervaloCerrarConexion As Long    '[Gonzalo]
Public IntervaloUserPuedeUsar As Long
Public IntervaloUserPuedeUsarClick As Long
Public IntervaloFlechasCazadores As Long
Public IntervaloPuedeSerAtacado As Long
Public IntervaloAtacable As Long
Public IntervaloOwnedNpc As Long

'BALANCE

Public PorcentajeRecuperoMana As Integer

Public MinutosWs As Long
Public Puerto As Integer

Public BootDelBackUp As Byte
Public Lloviendo As Boolean
Public DeNoche As Boolean

Function VersionOK(ByVal Ver As String) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    VersionOK = (Ver = ULTIMAVERSION)
End Function

Sub ReSpawnOrigPosNpcs()
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error Resume Next

    Dim i As Long
    Dim MiNPC As npc

    For i = 1 To LastNPC

        'OJO
        If Npclist(i).flags.NPCActive Then
            If InMapBounds(Npclist(i).Orig.map, Npclist(i).Orig.X, Npclist(i).Orig.Y) And Npclist(i).Numero = Guardias Then
                MiNPC = Npclist(i)
                Call QuitarNPC(i)
                Call ReSpawnNpc(MiNPC)
            End If
        End If

    Next i

End Sub

Sub WorldSave()

    On Error Resume Next

    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Iniciando WorldSave", FontTypeNames.FONTTYPE_SERVER))
    Call ReSpawnOrigPosNpcs    'respawn de los guardias en las pos originales

    FrmStat.ProgressBar1.min = 0
    FrmStat.ProgressBar1.max = CantMaps
    FrmStat.ProgressBar1.Value = 0

    If CantMaps > 0 Then

        Dim loopX As Long, ID As Integer, RutaBackup As String
        RutaBackup = App.Path & "\WorldBackUp\Mapa"

        For loopX = 1 To CantMaps
            ID = MapBackup(loopX)

            If ID > 0 Then
                Call GrabarMapa(ID, RutaBackup & ID)
                FrmStat.ProgressBar1.Value = FrmStat.ProgressBar1.Value + 1
            End If
        Next loopX

    End If

    FrmStat.Visible = False

    Call SaveForums
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> WorldSave ha concluído.", FontTypeNames.FONTTYPE_SERVER))

End Sub

Public Sub PurgarPenas()

    Dim i As Long

    For i = 1 To LastUser
        If UserList(i).flags.UserLogged Then
            If UserList(i).Counters.Pena > 0 Then
                UserList(i).Counters.Pena = UserList(i).Counters.Pena - 1

                If UserList(i).Counters.Pena < 1 Then
                    UserList(i).Counters.Pena = 0
                    Call WarpUserChar(i, Libertad.map, Libertad.X, Libertad.Y, True)
                    Call WriteConsoleMsg(i, "¡Has sido liberado!", FontTypeNames.FONTTYPE_INFO)

                    Call FlushBuffer(i)
                End If
            End If
        End If
    Next i
End Sub

Public Sub Encarcelar(ByVal UserIndex As Integer, ByVal Minutos As Long, Optional ByVal GmName As String = vbNullString)

    UserList(UserIndex).Counters.Pena = Minutos

    Call WarpUserChar(UserIndex, Prision.map, Prision.X, Prision.Y, True)

    If LenB(GmName) = 0 Then
        Call WriteConsoleMsg(UserIndex, "Has sido encarcelado, deberás permanecer en la cárcel " & Minutos & " minutos.", FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(UserIndex, GmName & " te ha encarcelado, deberás permanecer en la cárcel " & Minutos & " minutos.", FontTypeNames.FONTTYPE_INFO)
    End If
    If UserList(UserIndex).flags.Traveling = 1 Then
        UserList(UserIndex).flags.Traveling = 0
        UserList(UserIndex).Counters.goHome = 0
        Call WriteMultiMessage(UserIndex, eMessages.CancelHome)
    End If
End Sub


Public Sub BorrarUsuario(ByVal UserName As String)
    On Error Resume Next
    If FileExist(CharPath & UCase$(UserName) & ".chr", vbNormal) Then
        Kill CharPath & UCase$(UserName) & ".chr"
    End If
End Sub

Public Function BANCheck(ByVal Name As String) As Boolean
    BANCheck = (val(GetVar(App.Path & "\charfile\" & Name & ".chr", "FLAGS", "Ban")) = 1)
End Function

Public Function PersonajeExiste(ByVal Name As String) As Boolean
    PersonajeExiste = FileExist(CharPath & UCase$(Name) & ".chr", vbNormal)
End Function

Public Function CuentaExiste(ByVal Name As String) As Boolean
    CuentaExiste = FileExist(AccountPath & UCase$(Name) & ".acc", vbNormal)
End Function

Public Function UnBan(ByVal Name As String) As Boolean

    'Unban the character
    Call WriteVar(App.Path & "\charfile\" & Name & ".chr", "FLAGS", "Ban", "0")

    'Remove it from the banned people database
    Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", Name, "BannedBy", "NOBODY")
    Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", Name, "Reason", "NO REASON")
End Function

Public Function MD5ok(ByVal md5formateado As String) As Boolean

    Dim i As Integer

    If MD5ClientesActivado = 1 Then
        For i = 0 To UBound(MD5s)
            If (md5formateado = MD5s(i)) Then
                MD5ok = True
                Exit Function
            End If
        Next i
        MD5ok = False
    Else
        MD5ok = True
    End If

End Function

Public Sub MD5sCarga()

    Dim LoopC As Integer

    MD5ClientesActivado = val(GetVar(IniPath & "Server.ini", "MD5Hush", "Activado"))

    If MD5ClientesActivado = 1 Then
        ReDim MD5s(val(GetVar(IniPath & "Server.ini", "MD5Hush", "MD5Aceptados")))
        For LoopC = 0 To UBound(MD5s)
            MD5s(LoopC) = GetVar(IniPath & "Server.ini", "MD5Hush", "MD5Aceptado" & (LoopC + 1))
            MD5s(LoopC) = txtOffset(hexMd52Asc(MD5s(LoopC)), 55)
        Next LoopC
    End If

End Sub

Public Sub BanIpAgrega(ByVal ip As String)

    BanIps.Add ip

    Call BanIpGuardar
End Sub

Public Function BanIpBuscar(ByVal ip As String) As Long

    Dim Dale As Boolean
    Dim LoopC As Long

    Dale = True
    LoopC = 1
    Do While LoopC <= BanIps.Count And Dale
        Dale = (BanIps.Item(LoopC) <> ip)
        LoopC = LoopC + 1
    Loop

    If Dale Then
        BanIpBuscar = 0
    Else
        BanIpBuscar = LoopC - 1
    End If
End Function

Public Function BanIpQuita(ByVal ip As String) As Boolean

    On Error Resume Next

    Dim N As Long

    N = BanIpBuscar(ip)
    If N > 0 Then
        BanIps.Remove N
        BanIpGuardar
        BanIpQuita = True
    Else
        BanIpQuita = False
    End If

End Function

Public Sub BanIpGuardar()

    Dim ArchivoBanIp As String
    Dim ArchN As Long
    Dim LoopC As Long

    ArchivoBanIp = App.Path & "\Dat\BanIps.dat"

    ArchN = FreeFile()
    Open ArchivoBanIp For Output As #ArchN

    For LoopC = 1 To BanIps.Count
        Print #ArchN, BanIps.Item(LoopC)
    Next LoopC

    Close #ArchN

End Sub

Public Sub BanIpCargar()

    Dim ArchN As Long
    Dim Tmp As String
    Dim ArchivoBanIp As String

    ArchivoBanIp = App.Path & "\Dat\BanIps.dat"

    Do While BanIps.Count > 0
        BanIps.Remove 1
    Loop

    ArchN = FreeFile()
    Open ArchivoBanIp For Input As #ArchN

    Do While Not EOF(ArchN)
        Line Input #ArchN, Tmp
        BanIps.Add Tmp
    Loop

    Close #ArchN

End Sub

Public Sub ActualizaEstadisticasWeb()

    Static Andando As Boolean
    Static Contador As Long
    Dim Tmp As Boolean

    Contador = Contador + 1

    If Contador >= 10 Then
        Contador = 0
        Tmp = EstadisticasWeb.EstadisticasAndando()

        If Andando = False And Tmp = True Then
            Call InicializaEstadisticas
        End If

        Andando = Tmp
    End If

End Sub

Public Function UserDarPrivilegioLevel(ByVal Name As String) As PlayerType

    If EsAdmin(Name) Then
        UserDarPrivilegioLevel = PlayerType.Admin
    ElseIf EsDios(Name) Then
        UserDarPrivilegioLevel = PlayerType.Dios
    ElseIf EsSemiDios(Name) Then
        UserDarPrivilegioLevel = PlayerType.SemiDios
    ElseIf EsConsejero(Name) Then
        UserDarPrivilegioLevel = PlayerType.Consejero
    Else
        UserDarPrivilegioLevel = PlayerType.User
    End If

End Function

Public Sub BanCharacter(ByVal bannerUserIndex As Integer, ByVal UserName As String, ByVal Reason As String)

    Dim tUser As Integer
    Dim userPriv As Byte
    Dim cantPenas As Byte
    Dim Rank As Integer

    If InStrB(UserName, "+") Then
        UserName = Replace(UserName, "+", " ")
    End If

    tUser = NameIndex(UserName)

    Rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero

    With UserList(bannerUserIndex)
        If tUser <= 0 Then
            Call WriteConsoleMsg(bannerUserIndex, "El usuario no está online.", FontTypeNames.FONTTYPE_TALK)

            If FileExist(CharPath & UserName & ".chr", vbNormal) Then
                userPriv = UserDarPrivilegioLevel(UserName)

                If (userPriv And Rank) > (.flags.Privilegios And Rank) Then
                    Call WriteConsoleMsg(bannerUserIndex, "No puedes banear a al alguien de mayor jerarquía.", FontTypeNames.FONTTYPE_INFO)
                Else
                    If GetVar(CharPath & UserName & ".chr", "FLAGS", "Ban") <> "0" Then
                        Call WriteConsoleMsg(bannerUserIndex, "El personaje ya se encuentra baneado.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call LogBanFromName(UserName, bannerUserIndex, Reason)
                        Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor> " & .Name & " ha baneado a " & UserName & ".", FontTypeNames.FONTTYPE_SERVER))

                        'ponemos el flag de ban a 1
                        Call WriteVar(CharPath & UserName & ".chr", "FLAGS", "Ban", "1")
                        'ponemos la pena
                        cantPenas = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
                        Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", cantPenas + 1)
                        Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & cantPenas + 1, LCase$(.Name) & ": BAN POR " & LCase$(Reason) & " " & Date & " " & time)

                        If (userPriv And Rank) = (.flags.Privilegios And Rank) Then
                            .flags.Ban = 1
                            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.Name & " banned by the server por bannear un Administrador.", FontTypeNames.FONTTYPE_FIGHT))
                            Call CloseSocket(bannerUserIndex)
                        End If

                        Call LogGM(.Name, "BAN a " & UserName)
                    End If
                End If
            Else
                Call WriteConsoleMsg(bannerUserIndex, "El pj " & UserName & " no existe.", FontTypeNames.FONTTYPE_INFO)
            End If
        Else
            If (UserList(tUser).flags.Privilegios And Rank) > (.flags.Privilegios And Rank) Then
                Call WriteConsoleMsg(bannerUserIndex, "No puedes banear a al alguien de mayor jerarquía.", FontTypeNames.FONTTYPE_INFO)
            End If

            Call LogBan(tUser, bannerUserIndex, Reason)
            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor> " & .Name & " ha baneado a " & UserList(tUser).Name & ".", FontTypeNames.FONTTYPE_SERVER))

            'Ponemos el flag de ban a 1
            UserList(tUser).flags.Ban = 1

            If (UserList(tUser).flags.Privilegios And Rank) = (.flags.Privilegios And Rank) Then
                .flags.Ban = 1
                Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.Name & " banned by the server por bannear un Administrador.", FontTypeNames.FONTTYPE_FIGHT))
                Call CloseSocket(bannerUserIndex)
            End If

            Call LogGM(.Name, "BAN a " & UserName)

            'ponemos el flag de ban a 1
            Call WriteVar(CharPath & UserName & ".chr", "FLAGS", "Ban", "1")
            'ponemos la pena
            cantPenas = val(GetVar(CharPath & UserName & ".chr", "PENAS", "Cant"))
            Call WriteVar(CharPath & UserName & ".chr", "PENAS", "Cant", cantPenas + 1)
            Call WriteVar(CharPath & UserName & ".chr", "PENAS", "P" & cantPenas + 1, LCase$(.Name) & ": BAN POR " & LCase$(Reason) & " " & Date & " " & time)

            Call CloseSocket(tUser)
        End If
    End With
End Sub

