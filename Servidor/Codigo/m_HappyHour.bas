Attribute VB_Name = "m_HappyHour"
Option Explicit

Private iniHappyHourActivado As Byte

Public HappyHour As Single
Public HappyHourActivated As Boolean

Private Type tHappyHour
    Multi As Single
    Hour As Integer
End Type

Private iDay As Byte
Private HappyHourDays(1 To 7) As tHappyHour

Public Sub CargarHappyHour()

    Dim Leer As clsIniManager
    Set Leer = New clsIniManager

    Call Leer.Initialize(DatPath & "HappyHour.dat")

    ' HappyHour
    iniHappyHourActivado = val(Leer.GetValue("HAPPYHOUR", "Activado"))

    ' Dias y porcentaje de exp
    Dim lTemp As Long
    Dim sTemp As String

    For lTemp = 1 To 7
        sTemp = Leer.GetValue("HAPPYHOUR", StringSinTildes(WeekdayName(lTemp)))
        HappyHourDays(lTemp).Hour = val(ReadField(1, sTemp, 45))
        HappyHourDays(lTemp).Multi = val(ReadField(2, sTemp, 45))
        If HappyHourDays(lTemp).Hour < 0 Or HappyHourDays(lTemp).Hour > 23 Then HappyHourDays(lTemp).Hour = 20    ' Hora de 0 a 23.
        If HappyHourDays(lTemp).Multi < 0 Then HappyHourDays(lTemp).Multi = 0
    Next

    Set Leer = Nothing

End Sub

Public Function StringSinTildes(ByRef str As String) As String

    Dim temp As String
    temp = str

    If InStr(1, str, "á") > 0 Then temp = Replace(temp, "á", "a")
    If InStr(1, str, "é") > 0 Then temp = Replace(temp, "é", "e")
    If InStr(1, str, "í") > 0 Then temp = Replace(temp, "í", "i")
    If InStr(1, str, "ó") > 0 Then temp = Replace(temp, "ó", "o")
    If InStr(1, str, "ú") > 0 Then temp = Replace(temp, "ú", "u")

    StringSinTildes = temp

End Function

Public Sub PasarMinutoHappy()

    If iniHappyHourActivado Then

        Dim tmpHappyHour As Double

        ' HappyHour
        iDay = Weekday(Date)
        tmpHappyHour = HappyHourDays(iDay).Multi

        If tmpHappyHour <> HappyHour Then
            If HappyHourActivated Then
                ' Reestablece la exp de los npcs
                If HappyHour <> 0 Then Call UpdateNpcsExp(1 / HappyHour)
            End If

            If tmpHappyHour = 1 Then       ' Desactiva
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("¡Ha concluido la Happy Hour!", FontTypeNames.FONTTYPE_DIOS))
                HappyHourActivated = False

            Else       ' Activa?
                If HappyHourDays(iDay).Hour = Hour(Now) And tmpHappyHour > 1 Then       'Es la hora pautada?
                    Call UpdateNpcsExp(tmpHappyHour)

                    If HappyHour <> 1 Then
                        Call SendData(SendTarget.ToAll, 0, _
                                      PrepareMessageConsoleMsg("Se activo el Happy Hour, a partir de ahora los NPCS dan el doble de experiencia", FontTypeNames.FONTTYPE_DIOS))
                    Else
                        Call SendData(SendTarget.ToAll, 0, _
                                      PrepareMessageConsoleMsg("Se activo el Happy Hour, a partir de ahora los NPCS dan el doble de experiencia", FontTypeNames.FONTTYPE_DIOS))
                    End If

                    HappyHourActivated = True
                Else
                    HappyHourActivated = False
                End If
            End If

            HappyHour = tmpHappyHour
        End If
    Else
        ' Si estaba activado, lo deshabilitamos
        If HappyHour <> 0 Then
            Call UpdateNpcsExp(1 / HappyHour)
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("¡Ha concluido la Happy Hour!", FontTypeNames.FONTTYPE_DIOS))
            HappyHourActivated = False
            HappyHour = 0
        End If
    End If

End Sub

Public Sub UpdateNpcsExp(ByVal Multiplicador As Single)

    Dim NpcIndex As Long

    For NpcIndex = 1 To LastNPC
        With Npclist(NpcIndex)
            If .GiveEXP <> 0 Then
                .GiveEXP = .GiveEXP * Multiplicador
                .flags.ExpCount = .flags.ExpCount * Multiplicador
            End If
        End With
    Next NpcIndex

End Sub
