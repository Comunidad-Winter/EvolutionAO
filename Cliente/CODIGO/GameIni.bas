Attribute VB_Name = "GameIni"
Option Explicit

Public Type tCabecera    'Cabecera de los con
    Desc As String * 255
    CRC As Long
    MagicWord As Long
End Type

Public Type tGameIni
    Puerto As Long
    Musica As Byte
    fX As Byte
    tip As Byte
    Password As String
    Name As String
    DirGraficos As String
    DirSonidos As String
    DirMusica As String
    DirMapas As String
    NumeroDeBMPs As Long
    NumeroMapas As Integer
End Type

Public Type tSetupMods
    bDinamic As Boolean
    byMemory As Byte
    bUseVideo As Boolean
    bNoMusic As Boolean
    bNoSound As Boolean
    bNoRes As Boolean
    bNoSoundEffects As Boolean
    sGraficos As String * 13
    bGuildNews As Boolean
    bDie As Boolean
    bKill As Boolean
    byMurderedLevel As Byte
    bActive As Boolean
    bGldMsgConsole As Boolean
    bCantMsgs As Byte
End Type

Public ClientSetup As tSetupMods

Public MiCabecera As tCabecera
Public Config_Inicio As tGameIni

Public Sub IniciarCabecera(ByRef Cabecera As tCabecera)
    Cabecera.Desc = "Evolution Online"
    Cabecera.CRC = Rnd * 100
    Cabecera.MagicWord = Rnd * 10
End Sub

Public Function LeerGameIni() As tGameIni
    Dim N As Integer
    Dim GameIni As tGameIni
    N = FreeFile    'Valor en 1
    Open App.path & "\init\Inicio.con" For Binary As #N    'Leemos archivo
    Get #N, , MiCabecera    'Obtenemos dato 'La , despues del fichero freefile hace que se guarde.
    Get #N, , GameIni    'Obtenemos dato

    Close #N    'El primero termina seteado en 1 (Parece array que se define parametros).
    LeerGameIni = GameIni
End Function

Public Sub EscribirGameIni(ByRef GameIniConfiguration As tGameIni)
    On Local Error Resume Next

    Dim N As Integer
    N = FreeFile
    Open App.path & "\init\Inicio.con" For Binary As #N
    Put #N, , MiCabecera
    Put #N, , GameIniConfiguration
    Close #N
End Sub

