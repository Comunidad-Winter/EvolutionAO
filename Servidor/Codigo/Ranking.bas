Attribute VB_Name = "Ranking"
Option Explicit

Private Const RUTA_RANKING = "\server.ini"
Private Const MAX_TOP As Byte = 9

Public Valor(1 To MAX_TOP) As Byte
Public Nick(1 To MAX_TOP) As String

Public Function SaveRanking(ByVal UserIndex As Integer)

    Dim MiLevel As Byte, Nombre As String, Leer As New clsIniManager

    With UserList(UserIndex)
        Nombre = .Name
        MiLevel = .Stats.ELV

        Dim LoopC As Long, loopX As Long

        For LoopC = 1 To MAX_TOP
            If StrComp(UCase$(Nick(LoopC)), UCase$(Nombre)) = 0 Then
                If LoopC <> 1 Then    ' Si es el primero no hace falta que reordene.
                    For loopX = 1 To LoopC - 1
                        If MiLevel > Valor(loopX) Then

                            If Valor(loopX) <> 0 Then
                                If (loopX - LoopC) > 1 Then
                                    Call ChangePosRanking(loopX)
                                Else
                                    Call ReorderPosRanking(loopX, LoopC)
                                End If
                            End If

                            Nick(loopX) = Nombre
                            Valor(loopX) = MiLevel

                            Call Leer.Initialize(IniPath & RUTA_RANKING)

                            Call Leer.ChangeValue("RANKING", "Nombre" & loopX, Nick(loopX))
                            Call Leer.ChangeValue("RANKING", "Nivel" & loopX, Valor(loopX))

                            Call Leer.DumpFile(IniPath & RUTA_RANKING)
                            Set Leer = Nothing
                            Exit Function
                        End If
                    Next loopX
                End If

                ' ++ Si sube de nivel y estaba en ranking se lo actualizamos
                If MiLevel > Valor(LoopC) Then
                    Valor(LoopC) = MiLevel

                    Call Leer.Initialize(IniPath & RUTA_RANKING)

                    Call Leer.ChangeValue("RANKING", "Nivel" & LoopC, Valor(LoopC))

                    Call Leer.DumpFile(IniPath & RUTA_RANKING)
                    Set Leer = Nothing
                End If

                Exit Function
            End If
        Next LoopC

        ' ++ Lo agregamos dependiendo segun el top que paso +1
        For LoopC = 1 To MAX_TOP

            If MiLevel > Valor(LoopC) Then
                If Valor(LoopC) <> 0 Then Call ChangePosRanking(LoopC)

                Nick(LoopC) = Nombre
                Valor(LoopC) = MiLevel

                Call Leer.Initialize(IniPath & RUTA_RANKING)

                Call Leer.ChangeValue("RANKING", "Nombre" & LoopC, Nick(LoopC))
                Call Leer.ChangeValue("RANKING", "Nivel" & LoopC, Valor(LoopC))

                Call Leer.DumpFile(IniPath & RUTA_RANKING)
                Set Leer = Nothing
                Exit Function
            End If

        Next LoopC

    End With

End Function

Private Sub ChangePosRanking(ByVal Top As Byte)

    If Top > 8 Then Exit Sub

    Dim LoopC As Long
    Dim NickTemp As String
    Dim ValorTemp As Byte    '0-255

    NickTemp = Nick(Top)
    ValorTemp = Valor(Top)

    Dim TopIndex As Byte
    TopIndex = Top + 1

    Nick(Top) = Nick(TopIndex)
    Valor(Top) = Valor(TopIndex)

    Nick(TopIndex) = NickTemp
    Valor(TopIndex) = ValorTemp

    Dim Leer As New clsIniManager
    Call Leer.Initialize(IniPath & RUTA_RANKING)

    Call Leer.ChangeValue("RANKING", "Nombre" & Top, Nick(Top))
    Call Leer.ChangeValue("RANKING", "Nivel" & Top, Valor(Top))

    Call Leer.ChangeValue("RANKING", "Nombre" & TopIndex, Nick(TopIndex))
    Call Leer.ChangeValue("RANKING", "Nivel" & TopIndex, Valor(TopIndex))

    Call Leer.DumpFile(IniPath & RUTA_RANKING)
    Set Leer = Nothing

End Sub

Private Sub ReorderPosRanking(ByVal Top As Byte, ByVal PosID As Byte)

    Dim LoopC As Long
    Dim NickTemp(1 To MAX_TOP) As String
    Dim ValorTemp(1 To MAX_TOP) As Byte

    ' ++ Reordena los que estan abajo del que quitamos
    For LoopC = PosID To MAX_TOP - 1
        Nick(LoopC) = Nick(LoopC + 1)
        Valor(LoopC) = Valor(LoopC + 1)
    Next LoopC

    ' ++ Guardamos 2 variables temporales
    For LoopC = 1 To MAX_TOP
        NickTemp(LoopC) = Nick(LoopC)
        ValorTemp(LoopC) = Valor(LoopC)
    Next LoopC

    Dim Leer As New clsIniManager
    Call Leer.Initialize(IniPath & RUTA_RANKING)

    ' ++ Reordenamos.
    For LoopC = Top To MAX_TOP - 1
        Nick(LoopC + 1) = NickTemp(LoopC)
        Valor(LoopC + 1) = ValorTemp(LoopC)

        Call Leer.ChangeValue("RANKING", "Nombre" & LoopC + 1, Nick(LoopC))
        Call Leer.ChangeValue("RANKING", "Nivel" & LoopC + 1, Valor(LoopC))
    Next LoopC

    Call Leer.DumpFile(IniPath & RUTA_RANKING)
    Set Leer = Nothing

End Sub

Public Sub CargarRanking()

    Dim Leer As New clsIniManager
    Call Leer.Initialize(IniPath & RUTA_RANKING)

    Dim LoopC As Long

    For LoopC = 1 To MAX_TOP
        Nick(LoopC) = CStr(Leer.GetValue("RANKING", "Nombre" & LoopC))
        Valor(LoopC) = val(Leer.GetValue("RANKING", "Nivel" & LoopC))
    Next LoopC

    Set Leer = Nothing

End Sub
