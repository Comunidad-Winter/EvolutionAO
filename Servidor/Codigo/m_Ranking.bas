Attribute VB_Name = "m_Ranking"
Option Explicit

Public Const MAX_TOP As Byte = 9
Private Const MAX_RANKINGS As Byte = 6

Private Type tRanking
    Value(1 To MAX_TOP) As Long
    Nombre(1 To MAX_TOP) As String
End Type

Public Ranking(1 To MAX_RANKINGS) As tRanking

Public Enum eRanking
    TopFrags = 1
    TopOro = 2
    TopLevel = 3
    TopRetos = 4
    TopCriminales = 5
    TopCiudadanos = 6
End Enum

Public Sub CargarRanking()

    Dim i As Long
    Dim X As Long
    Dim str As String

    For X = 1 To MAX_RANKINGS
        For i = 1 To MAX_TOP
            str = GetVar(DatPath & "Ranking.dat", GetNameRanking(X), "Top" & i)
            Ranking(X).Nombre(i) = ReadField(1, str, 45)
            Ranking(X).Value(i) = val(ReadField(2, str, 45))
        Next i
    Next X

End Sub

Private Function GetNameRanking(ByVal Ranking As eRanking) As String
    
    Select Case Ranking
        Case eRanking.TopFrags
            GetNameRanking = "FRAGS"
        Case eRanking.TopOro
            GetNameRanking = "ORO"
        Case eRanking.TopLevel
            GetNameRanking = "NIVEL"
        Case eRanking.TopRetos
            GetNameRanking = "RETOS"
        Case eRanking.TopCriminales
            GetNameRanking = "CRIMINALES"
        Case eRanking.TopCiudadanos
            GetNameRanking = "CIUDADANOS"
    End Select
    
End Function

Private Function RenameValue(ByVal UserIndex As Integer, ByVal Ranking As eRanking) As Long

    With UserList(UserIndex)

        Select Case Ranking

            Case eRanking.TopFrags
                RenameValue = .Stats.UsuariosMatados
                
            Case eRanking.TopOro
                RenameValue = .Stats.GLD

            Case eRanking.TopLevel
                RenameValue = .Stats.ELV

            Case eRanking.TopRetos
                RenameValue = .Stats.RetosGanados

            Case eRanking.TopCriminales
                RenameValue = .Faccion.CriminalesMatados

            Case eRanking.TopCiudadanos
                RenameValue = .Faccion.CiudadanosMatados

        End Select

    End With

End Function

Private Sub GuardarRanking(ByVal Rank As eRanking)

    Dim Leer As clsIniManager
    Set Leer = New clsIniManager
    Call Leer.Initialize(DatPath & "Ranking.Dat")

    Dim LoopC As Long

    For LoopC = 1 To MAX_TOP
        Call Leer.ChangeValue(GetNameRanking(Rank), "Top" & LoopC, Ranking(Rank).Nombre(LoopC) & "-" & Ranking(Rank).Value(LoopC))
    Next LoopC

    Call Leer.DumpFile(DatPath & "Ranking.Dat")
    Set Leer = Nothing

End Sub

Public Sub CheckRankingUser(ByVal UserIndex As Integer, ByVal Rank As eRanking)

    If EsGm(UserIndex) Then Exit Sub

    Dim X As Long
    Dim Y As Long
    Dim Z As Long
    Dim i As Long
    
    Dim Value As Long
    Dim Actualizacion As Byte
    Dim aux As String
    Dim Pos As Byte
    Dim TipoRank As String
    
    Value = RenameValue(UserIndex, Rank)

    With UserList(UserIndex)
        TipoRank = UCase$(.Name)

        For i = 1 To MAX_TOP
            If Ranking(Rank).Nombre(i) = TipoRank Then
                Pos = i
                Exit For
            End If
        Next i

        If Pos <> 0 Then
            If Value <> Ranking(Rank).Value(Pos) Then
                Ranking(Rank).Value(Pos) = Value
                'If Pos = 1 Then Exit Sub

                For Y = 1 To MAX_TOP
                    For Z = 1 To MAX_TOP - Y

                        If Ranking(Rank).Value(Z) < Ranking(Rank).Value(Z + 1) Then
                            aux = Ranking(Rank).Value(Z)
                            Ranking(Rank).Value(Z) = Ranking(Rank).Value(Z + 1)
                            Ranking(Rank).Value(Z + 1) = aux

                            aux = Ranking(Rank).Nombre(Z)
                            Ranking(Rank).Nombre(Z) = Ranking(Rank).Nombre(Z + 1)
                            Ranking(Rank).Nombre(Z + 1) = aux
                            Actualizacion = 1
                        End If
                    Next Z
                Next Y

                If Actualizacion <> 0 Then
                    Call GuardarRanking(Rank)
                End If
            End If

            Exit Sub
        End If

        For X = 1 To MAX_TOP
            If Value > Ranking(Rank).Value(X) Then
                Call ActualizarRanking(X, Rank, TipoRank, Value)
                Exit Sub
            End If
        Next X

    End With

End Sub

Private Sub ActualizarRanking(ByVal Top As Byte, ByVal Rank As eRanking, ByVal UserName As String, ByVal Value As Long)

    Dim Valor(1 To MAX_TOP) As Long
    Dim Nombre(1 To MAX_TOP) As String

    Dim LoopC As Long

    For LoopC = 1 To MAX_TOP
        Valor(LoopC) = Ranking(Rank).Value(LoopC)
        Nombre(LoopC) = Ranking(Rank).Nombre(LoopC)
    Next LoopC

    For LoopC = Top To MAX_TOP - 1
        Ranking(Rank).Value(LoopC + 1) = Valor(LoopC)
        Ranking(Rank).Nombre(LoopC + 1) = Nombre(LoopC)
    Next LoopC

    Ranking(Rank).Value(Top) = Value
    Ranking(Rank).Nombre(Top) = UserName

    Call GuardarRanking(Rank)

End Sub


