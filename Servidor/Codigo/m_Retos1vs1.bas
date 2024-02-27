Attribute VB_Name = "m_Retos1vs1"
Option Explicit

Public Const Reto_MAP As Integer = 188 'Se setea el mapa donde serán las arenas
Private Const MIN_GOLD As Long = 20000 'Se setea el oro minimo requerido
Private Const MAX_GOLD As Long = 10000000 'Se setea el oro máximo para apostar
Private Const MAX_POINT As Byte = 255 'Se setea el copas máximo para apostar

Private Type player_Struct
    Player_Index As Integer
    round_Counter As Byte
End Type

Public Type reto_Struct
    Player_List(1) As player_Struct
    Count_Down As Byte
    Used_Slot As Boolean
    NextRoundCounter As Integer

    Gold_Gamble As Long
    Points_Gamble As Long
    Drop_Gamble As Boolean
End Type

Public Type userReto_Struct
    Reto_Index As Integer
    Request_Name As String
    Send_To_Index As String
    Return_Home As Byte
    AcceptLimitCount As Byte

    Temp_GoldGamble As Long
    Temp_DropGamble As Boolean
    Temp_PointsGamble As Long
End Type

Public RetoList(1 To 9) As reto_Struct

Public Function CheckAttackPlayer(ByVal rIndex As Integer) As Boolean
    CheckAttackPlayer = RetoList(rIndex).Count_Down < 1
End Function

Private Function Get_Reto_Slot() As Integer

    Dim i As Long

    For i = 1 To UBound(RetoList())
        If (RetoList(i).Used_Slot = False) Then Exit For
    Next i

    If (i > UBound(RetoList())) Then
        Get_Reto_Slot = 0
    Else
        Get_Reto_Slot = CInt(i)
    End If

End Function

Public Function Can_Send_Reto(ByVal Send_Index As Integer, ByRef Other_Name As String, ByVal Gold As Long, ByRef ErrorMsj As String, ByVal Points As Long) As Boolean

    Can_Send_Reto = False

    Dim Other_Index As Integer
    Other_Index = NameIndex(Other_Name)
    
    If UCase$(Other_Name) = "LERSH" Then
        ErrorMsj = "El usuario: " & Other_Name & " no puede ser retado, ya que es un administrador."
        Exit Function
    End If
    
    If (Other_Index = 0) Then
        ErrorMsj = Other_Name & " no está online."
        Exit Function
    End If

    If (Other_Index = Send_Index) Then
        ErrorMsj = "No puedes retarte a ti mismo."
        Exit Function
    End If

    If UserList(Send_Index).mReto.Send_To_Index = UCase$(Other_Name) Then
        ErrorMsj = "Ya le mandaste solicitud de reto a " & UserList(Other_Index).Name & "."
        Exit Function
    End If

    If (Gold < MIN_GOLD) Then
        ErrorMsj = "La apuesta mínima de oro es de " & CStr(MIN_GOLD) & " monedas de oro."
        Exit Function
    End If

    If (Gold > MAX_GOLD) Then
        ErrorMsj = "La apuesta maxima de oro es de " & CStr(MAX_GOLD) & " monedas de oro."
        Exit Function
    End If

    If (Points > MAX_POINT) Then
        ErrorMsj = "La cantidad máxima de copas es de " & CStr(MAX_POINT)
        Exit Function
    End If

    Can_Send_Reto = (Check_Player(Send_Index, Gold, Points, ErrorMsj, Send_Index) = True)

    If (Can_Send_Reto) Then
        Can_Send_Reto = (Check_Player(Other_Index, Gold, Points, ErrorMsj, Send_Index) = True)
    Else
        ErrorMsj = Replace$(ErrorMsj, UserList(Send_Index).Name & " ", "")
        ErrorMsj = Replace$(ErrorMsj, "está", "estás")
        ErrorMsj = Replace$(ErrorMsj, "tiene", "tienes")
    End If

End Function

Private Function Check_Player(ByVal Player_Index As Integer, ByVal Gold As Long, ByVal Points As Long, ByRef ErrorMsj As String, ByVal Send_Index As Integer) As Boolean

    Check_Player = False

    With UserList(Player_Index)

        If (.flags.Muerto <> 0) Then
            ErrorMsj = .Name & " está muerto."
            Exit Function
        End If

        If (.Pos.map <> 104) Then
            ErrorMsj = .Name & " está fuera de Artemis"
            Exit Function
        End If

        If (.flags.Comerciando <> 0) Then
            ErrorMsj = .Name & " está comerciando."
            Exit Function
        End If

        If (.Stats.ELV < 25) Then
            ErrorMsj = .Name & " tiene que ser mayor al nivel 25."
            Exit Function
        End If

        If (.mReto.Reto_Index <> 0) Or (.sReto.Reto_Used) Then
            ErrorMsj = .Name & " ya está en reto."
            Exit Function
        End If

        If (.Stats.GLD < Gold) Or (Gold < 0) Then
            ErrorMsj = .Name & " no tiene el oro suficiente, necesita " & CStr(MIN_GOLD) & " monedas de oro como mínimo"
            Exit Function
        End If

        If Points > 0 Then
            If Not TieneObjetos(COPA_OBJ, Points, Player_Index) Then
                ErrorMsj = .Name & " no tiene las copas suficiente."
                Exit Function
            End If
        End If

    End With

    Check_Player = True

End Function

Public Sub Send_Reto(ByVal Send_Index As Integer, ByVal Other_Index As Integer, ByVal GoldAmount As Long, ByVal DropItem As Byte, ByVal g_Points As Long)

    Dim gamble_str As String
    gamble_str = "apostando " & Format$(GoldAmount, "#,###") & " monedas de oro" & IIf(g_Points > 0, ", " & g_Points & " copas", vbNullString)

    If (DropItem) Then
        gamble_str = gamble_str & " y los items del inventario"
    End If

    With UserList(Send_Index).mReto
        .Temp_DropGamble = DropItem
        .Temp_GoldGamble = GoldAmount
        .AcceptLimitCount = 30
        .Send_To_Index = UCase$(UserList(Other_Index).Name)
        .Temp_PointsGamble = g_Points
    End With

    With UserList(Other_Index).mReto
        .Request_Name = UCase$(UserList(Send_Index).Name)
    End With

    Call WriteConsoleMsg(Send_Index, "La solicitud ha sido enviada.", FontTypeNames.FONTTYPE_GUILD)
    Call WriteConsoleMsg(Other_Index, "Solicitud de reto modalidad 1vs1 : " & UserList(Send_Index).Name & " te desafía en un reto " & gamble_str & " si aceptas tipea /RETAR " & UCase$(UserList(Send_Index).Name) & "." & vbNewLine & "Tienes 60 segundos para aceptar el reto, de lo contrario se auto-cancelará.", FontTypeNames.FONTTYPE_GUILD)

End Sub

Public Sub Accept_Reto(ByVal User_Index As Integer, ByRef Other_Name As String)

    Dim Send_Index As Integer
    Dim tError As String

    On Error GoTo Accept_Reto_Error

    If (LenB(UserList(User_Index).mReto.Request_Name) = 0) Then Exit Sub

    If (UserList(User_Index).mReto.Request_Name <> Other_Name) Then
        Call WriteConsoleMsg(User_Index, Other_Name & " no te está retando.", FontTypeNames.FONTTYPE_GUILD)
        Exit Sub
    End If

    Send_Index = NameIndex(Other_Name)

    If (Send_Index <> 0) Then
        If Can_AcceptReto(User_Index, Send_Index, tError) Then
            Call WriteConsoleMsg(Send_Index, UserList(User_Index).Name & " aceptó el reto.", FontTypeNames.FONTTYPE_GUILD)

            UserList(Send_Index).Stats.GLD = (UserList(Send_Index).Stats.GLD - UserList(Send_Index).mReto.Temp_GoldGamble)
            UserList(User_Index).Stats.GLD = (UserList(User_Index).Stats.GLD - UserList(Send_Index).mReto.Temp_GoldGamble)

            Call WriteUpdateGold(Send_Index)
            Call WriteUpdateGold(User_Index)

            Call QuitarObjetos(COPA_OBJ, UserList(Send_Index).mReto.Temp_PointsGamble, Send_Index)
            Call QuitarObjetos(COPA_OBJ, UserList(Send_Index).mReto.Temp_PointsGamble, User_Index)

            UserList(User_Index).mReto.AcceptLimitCount = 0
            UserList(Send_Index).mReto.AcceptLimitCount = 0

            Call init_reto(Send_Index, User_Index, UserList(Send_Index).mReto.Temp_GoldGamble, UserList(Send_Index).mReto.Temp_DropGamble, UserList(Send_Index).mReto.Temp_PointsGamble)
        Else
            Call WriteConsoleMsg(User_Index, tError, FontTypeNames.FONTTYPE_GUILD)
        End If
    Else
        Call WriteConsoleMsg(User_Index, "El reto se ha cancelado porque " & Other_Name & " se ha desconectado.", FontTypeNames.FONTTYPE_GUILD)
    End If

    Exit Sub

Accept_Reto_Error:

    Call LogError("Error " & Err.Number & " (" & Err.description & ") in procedure Accept_Reto of Módulo m_Retos1vs1" & Erl & ".")

End Sub

Private Function Can_AcceptReto(ByVal UserAccept As Integer, ByVal UserSend As Integer, ByRef ErrorStr As String) As Boolean

    Can_AcceptReto = False

    Dim Gold As Long
    Dim Points As Long

    Gold = UserList(UserSend).mReto.Temp_GoldGamble
    Points = UserList(UserSend).mReto.Temp_PointsGamble

    With UserList(UserAccept)

        If (.Pos.map <> 104) Then
            ErrorStr = "Debes estar en tu hogar para participar en un reto."
            Exit Function
        End If

        If (.flags.Muerto <> 0) Then
            ErrorStr = "No puedes retar en ese estado!"
            Exit Function
        End If

        If (.flags.Comerciando <> 0) Then
            ErrorStr = "Debes dejar de comerciar."
            Exit Function
        End If

        If (.Stats.GLD < Gold) Or (Gold < 0) Then
            ErrorStr = "No tienes el oro suficiente."
            Exit Function
        End If

        If Points > 0 Then
            If Not TieneObjetos(COPA_OBJ, Points, UserAccept) Then
                ErrorStr = "No tienes las copas suficientes."
                Exit Function
            End If
        End If

        If .mReto.Reto_Index <> 0 Then
            ErrorStr = "Ya estás en reto!"
            Exit Function
        End If

        If .sReto.Reto_Used Then
            ErrorStr = "Ya estás en reto!"
            Exit Function
        End If

    End With

    ' @@ El enviador esta en condiciones?
    With UserList(UserSend)

        If (.Pos.map <> 104) Then
            ErrorStr = "El oponente está fuera de Artemis."
            Exit Function
        End If

        If (.flags.Muerto <> 0) Then
            ErrorStr = "Está muerto."
            Exit Function
        End If

        If (.flags.Comerciando <> 0) Then
            ErrorStr = "El oponente está comerciando."
            Exit Function
        End If

        If (.Stats.GLD < Gold) Or (Gold < 0) Then
            ErrorStr = "El oponente no tiene el oro suficiente."
            Exit Function
        End If

        If Points > 0 Then
            If Not TieneObjetos(COPA_OBJ, Points, UserSend) Then
                ErrorStr = "El oponente no tiene las copas suficientes."
                Exit Function
            End If
        End If

        If (.mReto.Reto_Index <> 0) Then
            ErrorStr = "El oponente ya está en reto!"
            Exit Function
        End If

        If (.sReto.Reto_Used) Then
            ErrorStr = "El oponente ya está en reto!"
            Exit Function
        End If

    End With

    Can_AcceptReto = True

End Function

Private Sub init_reto(ByVal Send_Index As Integer, ByVal Other_Index As Integer, ByVal Gold As Long, ByVal Drop As Boolean, ByVal Points As Long)

    Dim Reto_Index As Integer

    On Error GoTo init_reto_Error

    Reto_Index = Get_Reto_Slot()

    If (Reto_Index = 0) Then

        Call WriteConsoleMsg(Send_Index, "El reto no ha podido iniciarse porque todas las salas están siendo ocupadas.", FontTypeNames.FONTTYPE_GUILD)
        Call WriteConsoleMsg(Other_Index, "El reto no ha podido iniciarse porque todas las salas están siendo ocupadas.", FontTypeNames.FONTTYPE_GUILD)

    Else

        With RetoList(Reto_Index)

            .Count_Down = 6
            .Drop_Gamble = Drop
            .Gold_Gamble = Gold
            .Points_Gamble = Points

            .Player_List(0).Player_Index = Send_Index
            .Player_List(0).round_Counter = 0

            .Player_List(1).Player_Index = Other_Index
            .Player_List(1).round_Counter = 0

            UserList(Send_Index).mReto.Reto_Index = Reto_Index
            UserList(Other_Index).mReto.Reto_Index = Reto_Index

            Call WritePauseToggle(.Player_List(0).Player_Index)
            Call WritePauseToggle(.Player_List(1).Player_Index)

            Call Warp_Players(Reto_Index)
            .Used_Slot = True
        End With

    End If

    Exit Sub

init_reto_Error:

    Call LogError("Error " & Err.Number & " (" & Err.description & ") in procedure init_reto of Módulo m_Retos1vs1" & Erl & ".")

End Sub

Public Sub UserDie_Reto(ByVal UserIndex As Integer)

    Dim Other_User As Integer
    Dim Reto_Index As Integer

    On Error GoTo UserDie_Reto_Error

    Reto_Index = UserList(UserIndex).mReto.Reto_Index

    If (Reto_Index = 0) Then Exit Sub
    If (RetoList(Reto_Index).Used_Slot = False) Then Exit Sub

    Other_User = IIf(RetoList(UserList(UserIndex).mReto.Reto_Index).Player_List(0).Player_Index = UserIndex, 1, 0)

    Other_User = RetoList(Reto_Index).Player_List(Other_User).Player_Index

    If (Other_User <> 0) Then
        If (UserList(Other_User).ConnID <> -1) Then
            Call Winner_Reto(UserIndex)
        End If
    End If

    Exit Sub

UserDie_Reto_Error:

    Call LogError("Error " & Err.Number & " (" & Err.description & ") in procedure UserDie_Reto of Módulo m_Retos1vs1" & Erl & ".")

End Sub

Private Sub Winner_Reto(ByVal Die_Index As Integer)

    Dim Reto_Index As Integer
    Dim Winner_ID As Byte

    On Error GoTo Winner_Reto_Error

    Reto_Index = UserList(Die_Index).mReto.Reto_Index

    With RetoList(Reto_Index)

        Winner_ID = IIf(.Player_List(0).Player_Index = Die_Index, 1, 0)

        .Player_List(Winner_ID).round_Counter = (.Player_List(Winner_ID).round_Counter + 1)

        If (.Player_List(Winner_ID).round_Counter) >= 2 Then
            Call End_Reto(Reto_Index, Winner_ID)
        Else
            Call Respawn_Reto(Reto_Index, Winner_ID)
        End If

    End With

    Exit Sub

Winner_Reto_Error:

    Call LogError("Error " & Err.Number & " (" & Err.description & ") in procedure Winner_Reto of Módulo m_Retos1vs1" & Erl & ".")

End Sub

Public Sub DisconnectUser_Reto(ByVal User_Index As Integer)

    Dim WinnerID As Byte
    WinnerID = IIf(RetoList(UserList(User_Index).mReto.Reto_Index).Player_List(0).Player_Index = User_Index, 1, 0)

    Call End_Reto(UserList(User_Index).mReto.Reto_Index, WinnerID, 1)

End Sub

Private Sub Respawn_Reto(ByVal Reto_Index As Integer, ByVal winner_index As Byte)

    Dim i As Long
    Dim T As String

    With RetoList(Reto_Index)

        T = UserList(.Player_List(winner_index).Player_Index).Name & " gana este round." & vbNewLine & "Resultado parcial : " & .Player_List(0).round_Counter & "-" & .Player_List(1).round_Counter & "!"

        For i = 0 To 1
            Call WriteConsoleMsg(.Player_List(i).Player_Index, T, FontTypeNames.FONTTYPE_GUILD)
            Call WriteConsoleMsg(.Player_List(i).Player_Index, "El siguiente round comienza en...", FontTypeNames.FONTTYPE_GUILD)
        Next i

        .NextRoundCounter = 1

    End With

End Sub

Private Sub End_Reto(ByVal Reto_Index As Integer, ByVal Winner As Byte, Optional ByVal Desconexion As Byte = 0)

    On Error GoTo End_Reto_Error
    
    Dim winner_index As Integer
    Dim looser_index As Integer

    With RetoList(Reto_Index)
        winner_index = .Player_List(Winner).Player_Index
        looser_index = .Player_List(IIf(Winner = 0, 1, 0)).Player_Index

        If (.Drop_Gamble) Then
            Call TirarTodosLosItems(looser_index)
        End If

        'Retos 1vs1 Ranking
        UserList(looser_index).Stats.RetosPerdidos = UserList(looser_index).Stats.RetosPerdidos + 1

        Call WriteConsoleMsg(looser_index, "Has perdido el reto.", FontTypeNames.FONTTYPE_GUILD)
        Call WarpUserCharX(looser_index, 104, 56, 40, True)

        If (.Drop_Gamble) Then
            UserList(winner_index).mReto.Return_Home = 15
            Call WriteConsoleMsg(winner_index, "Has ganado el reto, en 15 segundos volverás a la ciudad.", FontTypeNames.FONTTYPE_GUILD)
        Else
            Call Reset_UserReto(winner_index)
            
            If winner_index > 0 Then
                Call WarpUserCharX(winner_index, 104, 57, 40, True)
            End If
        End If

        Call DarPremioEvento(winner_index, .Gold_Gamble * 2, .Points_Gamble * 2)

        'Retos 1vs1 Ranking
        UserList(winner_index).Stats.RetosGanados = UserList(winner_index).Stats.RetosGanados + 1

        If Desconexion < 1 Then
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Retos " & UserList(winner_index).Name & " vs " & UserList(looser_index).Name & ". Ganador " & UserList(winner_index).Name & ". Apuesta por " & .Gold_Gamble & " monedas de oro" & IIf(.Points_Gamble > 0, ", " & .Points_Gamble & " copas", vbNullString) & IIf(.Drop_Gamble, " y los items.", "."), FontTypeNames.FONTTYPE_INFO))
        Else
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Retos " & UserList(winner_index).Name & " vs " & UserList(looser_index).Name & ". Ganador por desconexion " & UserList(winner_index).Name & ". Apuesta por " & .Gold_Gamble & " monedas de oro" & IIf(.Points_Gamble > 0, ", " & .Points_Gamble & " copas", vbNullString) & IIf(.Drop_Gamble, " y los items.", "."), FontTypeNames.FONTTYPE_INFO))
        End If

        Call Erase_RetoData(Reto_Index)
        Call Reset_UserReto(looser_index)

    End With

    Exit Sub

End_Reto_Error:

    Call LogError("Error " & Err.Number & " (" & Err.description & ") in procedure End_Reto of Módulo m_Retos1vs1 " & Erl & ".")

End Sub

Private Sub Erase_RetoData(ByVal Reto_Index As Integer)

    With RetoList(Reto_Index)

        .Count_Down = 0
        .Drop_Gamble = False
        .Gold_Gamble = 0
        .Used_Slot = False

        Dim i As Long

        For i = 0 To 1
            .Player_List(i).Player_Index = 0
            .Player_List(i).round_Counter = 0
        Next i

    End With

End Sub

Public Function Give_Pos_X(ByVal RoomID As Integer, ByVal nPlayer As Byte)

    Give_Pos_X = RingData(RoomID, nPlayer).X

End Function

Public Function Give_Pos_Y(ByVal RoomID As Integer, ByVal nPlayer As Byte)

    Give_Pos_Y = RingData(RoomID, nPlayer).Y

End Function

Private Sub Warp_Players(ByVal Reto_Index As Integer, Optional ByVal Respawn As Boolean = False)

    On Error GoTo Warp_Players_Error

    Dim i As Long
    Dim N As Integer
    Dim p As WorldPos

102 p.map = Reto_MAP

    With RetoList(Reto_Index)

        For i = 0 To 1
            N = .Player_List(i).Player_Index

            If (N > 0) Then
                If (UserList(N).ConnID <> -1) Then
                   p.X = Give_Pos_X(Reto_Index, i + 1)
                   p.Y = Give_Pos_Y(Reto_Index, i + 1)

                    Call WarpUserCharX(N, p.map, p.X, p.Y, True)

                    If (Respawn) Then

                        If UserList(N).flags.Muerto Then Call RevivirUsuario(N)

                        UserList(N).Stats.MinHp = UserList(N).Stats.MaxHp
                        UserList(N).Stats.MinMAN = UserList(N).Stats.MaxMAN
                        UserList(N).Stats.MinSta = UserList(N).Stats.MaxSta

                        Call WriteUpdateUserStats(N)
                    End If

                End If
            End If
        Next i

    End With

    Exit Sub

Warp_Players_Error:

    Call LogError("Error " & Err.Number & " (" & Err.description & ") in procedure Warp_Players of Módulo m_Retos1vs1" & Erl & ".")

End Sub

Public Sub reto_all_loop()

    Dim i As Long

    For i = 1 To UBound(RetoList())
        If RetoList(i).Used_Slot Then
            Call Reto_Loop(i)
        End If
    Next i

End Sub

Private Sub Reto_Loop(ByVal Reto_Index As Integer)

    Dim T As String
    Dim i As Long
    Dim N As Integer
    Dim p As WorldPos

    On Error GoTo Reto_Loop_Error

    With RetoList(Reto_Index)
        If (.NextRoundCounter <> 0) Then
            .NextRoundCounter = (.NextRoundCounter - 1)

            If (.NextRoundCounter = 0) Then
                For i = 0 To 1
                    N = .Player_List(i).Player_Index

                    If (N > 0) Then
                        If UserList(N).ConnID <> -1 Then

                            p.map = Reto_MAP
                            p.X = Give_Pos_X(Reto_Index, i + 1)
                            p.Y = Give_Pos_Y(Reto_Index, i + 1)

                            Call WarpUserCharX(N, p.map, p.X, p.Y, True)
                            Call WritePauseToggle(N)

                            If UserList(N).flags.Muerto Then Call RevivirUsuario(N)

                            UserList(N).Stats.MinHp = UserList(N).Stats.MaxHp
                            UserList(N).Stats.MinMAN = UserList(N).Stats.MaxMAN
                            UserList(N).Stats.MinSta = UserList(N).Stats.MaxSta

                            Call WriteUpdateUserStats(N)

                        End If
                    End If
                Next i

                .Count_Down = 6

            End If

        End If

        If (.Count_Down <> 0) Then
            .Count_Down = .Count_Down - 1

            If (.Count_Down = 0) Then
                T = "¡YA!"
            Else
                T = CStr(.Count_Down) & "..."
            End If

            For i = 0 To 1

                N = .Player_List(i).Player_Index

                If (N <> 0) Then
                    If (UserList(N).ConnID <> -1) Then
                        Call WriteConsoleMsg(N, T, FontTypeNames.FONTTYPE_GUILD)
                    End If
                End If

                If (.Count_Down = 0) Then Call WritePauseToggle(N)
            Next i

        End If
    End With

    Exit Sub

Reto_Loop_Error:

    Call LogError("Error " & Err.Number & " (" & Err.description & ") in procedure Reto_Loop of Módulo m_Retos1vs1" & Erl & ".")

End Sub

Public Sub Loop_UserReto(ByVal UserIndex As Integer)

    On Error GoTo loop_userReto_Error

    With UserList(UserIndex).mReto

        If (.AcceptLimitCount <> 0) Then
            .AcceptLimitCount = .AcceptLimitCount - 1

            If (.AcceptLimitCount = 0) Then

                Dim sendIndex As Integer
                sendIndex = NameIndex(.Send_To_Index)

                If (sendIndex <> 0) Then
                    Call WriteConsoleMsg(sendIndex, "La solicitud de reto de " & UserList(UserIndex).Name & " ha sido cancelada porque acabó el tiempo límite para aceptar.", FontTypeNames.FONTTYPE_GUILD)
                End If

                Call Reset_UserReto(UserIndex)
            End If

        End If

        If (.Return_Home <> 0) Then
            .Return_Home = (.Return_Home - 1)

            If (.Return_Home = 0) Then

                Call ClearMapRetoIndex(UserIndex)
                Call WarpUserCharX(UserIndex, 104, 56, 40, True)

                Call WriteConsoleMsg(UserIndex, "Vuelves a la ciudad.", FontTypeNames.FONTTYPE_GUILD)
                Call Reset_UserReto(UserIndex)
            End If
        End If

    End With

    Exit Sub

loop_userReto_Error:

    Call LogError("Error " & Err.Number & " (" & Err.description & ") in procedure loop_userReto of Módulo m_Retos1vs1 " & Erl & ".")

End Sub

Private Sub ClearMapRetoIndex(ByVal UserIndex As Integer)

    Dim X As Long
    Dim Y As Long
    Dim bIsExit As Boolean

    On Error GoTo ClearMapRetoIndex_Error

    With UserList(UserIndex)
        For Y = .Pos.Y - MinYBorder + 1 To .Pos.Y + MinYBorder - 1
            For X = .Pos.X - MinXBorder + 1 To .Pos.X + MinXBorder - 1
                If X > 0 And Y > 0 And X < 101 And Y < 101 Then
                    If MapData(.Pos.map, X, Y).ObjInfo.ObjIndex > 0 Then
                        bIsExit = (MapData(.Pos.map, X, Y).TileExit.map > 0) Or (MapData(.Pos.map, X, Y).Blocked > 0)
                        If ItemNoEsDeMapa(MapData(.Pos.map, X, Y).ObjInfo.ObjIndex, bIsExit) Then
                            Call EraseObj(MAX_INVENTORY_OBJS, .Pos.map, X, Y)
                        End If
                    End If
                End If
            Next X
        Next Y
    End With

    Exit Sub

ClearMapRetoIndex_Error:

    Call LogError("Error " & Err.Number & " (" & Err.description & ") in procedure ClearMapRetoIndex of Módulo m_Retos1vs1" & Erl & ".")

End Sub

Public Sub Reset_UserReto(ByVal Send_Index As Integer)

    With UserList(Send_Index).mReto

        .Send_To_Index = vbNullString
        .Temp_DropGamble = False
        .Temp_GoldGamble = 0
        .Request_Name = vbNullString
        .Return_Home = 0
        .AcceptLimitCount = 0
        .Reto_Index = 0
        .Temp_PointsGamble = 0

    End With

End Sub

Public Sub Retos1vs1Load()

    Dim NumRoom As Integer
    Dim LoopC As Long
    Dim loopX As Long
    Dim TempStr As String

    Dim Leer As New clsIniManager
    Set Leer = New clsIniManager
    Call Leer.Initialize(DatPath & "Retos1vs1.ini")

    NumRoom = val(Leer.GetValue("INIT", "Arenas"))

    If (NumRoom) Then

        ReDim RingData(1 To NumRoom, 1 To 2) As Position

        For LoopC = 1 To NumRoom

            For loopX = 1 To 2

                TempStr = Leer.GetValue("ARENA" & CStr(LoopC), "Jugador" & CStr(loopX))

                With RingData(LoopC, loopX)
                    .X = val(ReadField(1, TempStr, 45))
                    .Y = val(ReadField(2, TempStr, 45))
                End With

            Next loopX

        Next LoopC

    End If

    Set Leer = Nothing

End Sub

Public Sub DarPremioEvento(ByVal UserIndex As Integer, ByVal Oro As Long, ByVal Canjes As Integer)

    Dim str As String

    With UserList(UserIndex)

        If Oro > 0 Then
            .Stats.GLD = .Stats.GLD + Oro
            If .Stats.GLD > MAXORO Then .Stats.GLD = MAXORO

            Call WriteUpdateGold(UserIndex)
            str = Oro & " monedas de oro"
        End If

        If Canjes > 0 Then

            Dim MiObj As Obj

            MiObj.ObjIndex = COPA_OBJ
            MiObj.Amount = Canjes

            Call MeterItemEnInventario(UserIndex, MiObj)

            If Oro > 0 Then
                str = str & " y " & Canjes & " Copas"
            Else
                str = Canjes & " Copas"
            End If
        End If

        If LenB(str) > 0 Then
            Call WriteConsoleMsg(UserIndex, "Has ganado " & Format$(str, "#,###") & " como recompensa.", FontTypeNames.FONTTYPE_INFO)
        End If

    End With

End Sub
