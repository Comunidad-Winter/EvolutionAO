Attribute VB_Name = "m_Retos2vs2"
Option Explicit

Public Const MAPA_ARENAS As Integer = 189 'Se setea el mapa donde serán las arenas
Private Const MIN_GOLD As Long = 20000 'Se setea el oro minimo requerido
Private Const MAX_GOLD As Long = 10000000 'Se setea el oro máximo para apostar
Private Const MAX_POINT As Byte = 255 'Se setea el copas máximo para apostar

Public Type RuleStruct
    Drop_Inv As Boolean
    Gold_Gamble As Long
    Points_Gamble As Long
    RespawnToggle As Boolean
End Type

Public Type TeamStruct
    User_Index(1) As Integer
    Round_Count As Byte
    Return_City As Byte
End Type

Public Type RetoStruct
    Team_Array(1) As TeamStruct
    General_Rules As RuleStruct
    Count_Down As Byte
    Used_Ring As Boolean
    NextRoundCount As Integer
End Type

Public Type UserStruct
    TempStruct As RetoStruct
    Accept_Count As Byte
    Reto_Index As Integer
    Nick_Sender As String
    Reto_Used As Boolean
    Return_City As Byte
    AcceptedOK As Boolean
    AcceptLimit As Integer
End Type

Public Type RetoPosStruct
    X As Integer
    Y As Integer
End Type

Public Reto_List() As RetoStruct
Public RetoPos() As RetoPosStruct

Public Sub Loop_reto()

    Dim LoopC As Long

    For LoopC = LBound(Reto_List()) To UBound(Reto_List())
        If Reto_List(LoopC).Used_Ring Then
            Call Loop_Reto_Index(LoopC)
        End If
    Next LoopC

End Sub

Public Function Can_Attack(ByVal AttackerIndex As Integer, ByVal VictimIndex As Integer) As Boolean

    Dim retoIndex As Integer
    Dim TeamIndex As Integer
    Dim tempIndex As Integer
    Dim teamLoop As Long

    Can_Attack = True

    retoIndex = UserList(AttackerIndex).sReto.Reto_Index

    TeamIndex = -1

    If Reto_List(retoIndex).Used_Ring Then

        For teamLoop = 0 To 1
            If Reto_List(retoIndex).Team_Array(teamLoop).User_Index(0) = AttackerIndex Or Reto_List(retoIndex).Team_Array(teamLoop).User_Index(1) = AttackerIndex Then
                TeamIndex = teamLoop
                Exit For
            End If
        Next teamLoop

        If TeamIndex <> -1 Then
            tempIndex = IIf(Reto_List(retoIndex).Team_Array(TeamIndex).User_Index(0) = AttackerIndex, 1, 0)

            If Reto_List(retoIndex).Team_Array(TeamIndex).User_Index(tempIndex) = VictimIndex Then
                Can_Attack = False
            End If

        End If

    End If

End Function

Private Sub Loop_Reto_Index(ByVal Reto_Index As Integer)

    '
    ' @ amishar.-

    Dim i As Long
    Dim j As Long
    Dim h As Integer
    Dim M As String

    With Reto_List(Reto_Index)

        If (.NextRoundCount <> 0) Then
            .NextRoundCount = .NextRoundCount - 1

            If (.NextRoundCount = 0) Then
                Call Warp_Teams(Reto_Index, True)
                .Count_Down = 6
            End If
        End If

        If (.Count_Down <> 0) Then
            .Count_Down = (.Count_Down - 1)

            If (.Count_Down > 0) Then
                M = CStr(.Count_Down) & "..."
            Else
                M = "¡YA!"
            End If

            For i = 0 To 1
                For j = 0 To 1
                    h = .Team_Array(i).User_Index(j)

                    If (h <> 0) Then
                        If UserList(h).ConnID <> -1 Then
                            Call WriteConsoleMsg(h, M, FontTypeNames.FONTTYPE_GUILD)
                            If (.Count_Down = 0) Then Call WritePauseToggle(h)
                        End If
                    End If
                Next j
            Next i

        End If

    End With

End Sub

Public Function Get_Reto_Index() As Integer

    Dim LoopC As Long

    For LoopC = LBound(Reto_List()) To UBound(Reto_List())
        If (Reto_List(LoopC).Used_Ring = False) Then
            Get_Reto_Index = CInt(LoopC)
            Exit Function
        End If
    Next LoopC

    Get_Reto_Index = -1

End Function

Public Sub set_reto_struct(ByVal UserIndex As Integer, _
                           ByVal My_Team As String, _
                           ByRef Enemy_Name As String, _
                           ByRef Team_Enemy As String, _
                           ByVal Drop As Boolean, _
                           ByVal Gold As Long, _
                           ByVal Points As Long, _
                           ByVal Resu As Boolean)

    With UserList(UserIndex).sReto

        .Accept_Count = 0

        With .TempStruct
            .Count_Down = 0
            .Used_Ring = False

            With .Team_Array(0)
                .User_Index(0) = UserIndex
                .User_Index(1) = NameIndex(My_Team)
            End With

            With .Team_Array(1)
                .User_Index(0) = NameIndex(Enemy_Name)
                .User_Index(1) = NameIndex(Team_Enemy)
            End With

            With .General_Rules
                .Drop_Inv = Drop
                .Gold_Gamble = Gold
                .Points_Gamble = Points
                .RespawnToggle = Resu
            End With

        End With

    End With

End Sub

Public Sub User_RetoLoop(ByVal User_Index As Integer)

    With UserList(User_Index).sReto

        If (.AcceptLimit <> 0) Then
            .AcceptLimit = .AcceptLimit - 1

            If (.AcceptLimit < 1) Then

                Call Message_Reto(.TempStruct, "El reto se ha autocancelado debido a que el tiempo para aceptar ha llegado a su límite.")

                Dim j As Long
                Dim i As Long
                Dim N As Integer
                Dim B As UserStruct

                Dim UserName As String
                UserName = UCase$(UserList(User_Index).Name)

                For j = 0 To 1
                    For i = 0 To 1
                        N = .TempStruct.Team_Array(j).User_Index(i)

                        If N > 0 Then
                            If StrComp(UCase$(UserList(N).sReto.Nick_Sender), UserName) = 0 Then
                                UserList(N).sReto.Nick_Sender = vbNullString
                                UserList(N).sReto.AcceptedOK = False
                            End If
                        End If

                    Next i
                Next j

                UserList(User_Index).sReto = B

            End If

        End If

        If (.Return_City <> 0) Then
            .Return_City = .Return_City - 1

            If (.Return_City = 0) Then

                Dim X As Long
                Dim Y As Long
                Dim bIsExit As Boolean

                With UserList(User_Index)
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

                Call WarpUserCharX(User_Index, 104, 57, 40, True)
                Call WriteConsoleMsg(User_Index, "Regresas a la ciudad.", FontTypeNames.FONTTYPE_GUILD)

                Dim rIndex As Integer
                rIndex = .Reto_Index

                .Nick_Sender = vbNullString
                .Reto_Index = 0

                Call Clear_Data(rIndex)

            End If

        End If

    End With

End Sub

Public Sub Erase_UserData(ByVal User_Index As Integer)

    With UserList(User_Index).sReto

        Dim dumpStruct As RetoStruct

        .Accept_Count = 0
        .Nick_Sender = vbNullString
        .Reto_Index = 0
        .Reto_Used = False
        .TempStruct = dumpStruct
        .Return_City = 0
        .AcceptedOK = False

    End With

End Sub

Public Function Can_Send_Reto(ByVal User_Index As Integer, _
                              ByRef fError As String) As Boolean

    Can_Send_Reto = False

    With UserList(User_Index)

        If (.Pos.map <> 104) Then
            fError = "Debes estar en Artemis."
            Exit Function
        End If

        If (.flags.Muerto <> 0) Then
            fError = "¡Estás muerto!"
            Exit Function
        End If

        If (.Counters.Pena <> 0) Then
            fError = "Estás en la cárcel"
            Exit Function
        End If

        If (.sReto.TempStruct.General_Rules.Gold_Gamble < MIN_GOLD) Then
            fError = "El mínimo de oro para retar es de 20.000 monedas de oro."
            Exit Function
        End If

        If (.sReto.TempStruct.General_Rules.Gold_Gamble > MAX_GOLD) Then
            fError = "El máximo de oro para apostar en el reto son 10.000.000 monedas de oro."
            Exit Function
        End If

        If (.sReto.TempStruct.General_Rules.Points_Gamble < 0) Then
            fError = "Cantidad de copas inválidas."
            Exit Function
        End If

        If (.sReto.TempStruct.General_Rules.Points_Gamble > MAX_POINT) Then
            fError = "El máximo de copas para retar es de " & CStr(MAX_POINT) & " copas."
            Exit Function
        End If

        If (.Stats.GLD < .sReto.TempStruct.General_Rules.Gold_Gamble) Then
            fError = "No tienes el oro necesario"
            Exit Function
        End If

        If Not TieneObjetos(COPA_OBJ, .sReto.TempStruct.General_Rules.Points_Gamble, User_Index) Then
            fError = "No tienes las copas que deseas apostar."
            Exit Function
        End If

        If (.mReto.Reto_Index <> 0) Or (.sReto.Reto_Used = True) Then
            fError = .Name & " ya está en reto."
            Exit Function
        End If

        If (.sReto.Nick_Sender <> vbNullString) Then
            fError = "Ya has mandado reto"
            Exit Function
        End If

        If (.Stats.ELV < 25) Then
            fError = "Debes ser mayor a nivel 25!"
            Exit Function
        End If

        With .sReto.TempStruct

            Can_Send_Reto = Check_User(.Team_Array(0).User_Index(1), fError, .General_Rules.Gold_Gamble, .General_Rules.Points_Gamble, User_Index)

            If (Can_Send_Reto) Then
                Can_Send_Reto = Check_User(.Team_Array(1).User_Index(0), fError, .General_Rules.Gold_Gamble, .General_Rules.Points_Gamble, User_Index)
            Else
                Exit Function
            End If

            If (Can_Send_Reto) Then
                Can_Send_Reto = Check_User(.Team_Array(1).User_Index(1), fError, .General_Rules.Gold_Gamble, .General_Rules.Points_Gamble, User_Index)
            Else
                Exit Function
            End If

        End With

    End With

End Function

Private Function Check_AcceptUser(ByVal User_Index As Integer, ByVal nGold As Long, ByVal PointsGamble As Long, ByRef fError As String) As Boolean

    Check_AcceptUser = False

    With UserList(User_Index)

        If .flags.Muerto <> 0 Then
            fError = "Estás muerto."
            Exit Function
        End If

        If .Pos.map <> 104 Then
            fError = "Debes estar en Artemis para aceptar el reto."
            Exit Function
        End If

        If .Stats.GLD < nGold Then
            fError = "No tienes el oro suficiente para aceptar el reto."
            Exit Function
        End If

        If (.mReto.Reto_Index <> 0) Or (.sReto.Reto_Used = True) Then
            fError = "Ya estás en reto."
            Exit Function
        End If

        If Not TieneObjetos(COPA_OBJ, PointsGamble, User_Index) Then
            fError = "No tienes las copas necesarias para aceptar el reto."
            Exit Function
        End If

        If .sReto.AcceptedOK Then
            fError = "¡Ya has aceptado!"
            Exit Function
        End If

        If Not (UCase$(.Name) <> .sReto.Nick_Sender) Then
            Call WriteConsoleMsg(User_Index, "No te puedes aceptar a ti mismo.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If

    End With

    Check_AcceptUser = True

End Function

Private Function Check_User(ByVal User_Index As Integer, _
                            ByRef fError As String, _
                            ByVal goldGamble As Long, _
                            ByVal PointsGamble As Long, _
                            ByVal Send_Index As Integer) As Boolean

    Check_User = False

    If (User_Index = 0) Then
        fError = "Algún usuario está offline."
        Exit Function
    End If

    With UserList(User_Index)

        If .sReto.Nick_Sender = UserList(Send_Index).Name Then
            fError = "Ya le mandase solicitud de reto a " & .Name & "."
            Exit Function
        End If

        If (.flags.Muerto <> 0) Then
            fError = .Name & " ¡Está muerto!"
            Exit Function
        End If

        If (.Counters.Pena <> 0) Then
            fError = .Name & " Está en la cárcel"
            Exit Function
        End If

        If (.Pos.map <> 104) Then
            fError = .Name & " está fuera de Artemis."
            Exit Function
        End If

        If (.mReto.Reto_Index <> 0) Or (.sReto.Reto_Used = True) Then
            fError = .Name & " ya está en reto."
            Exit Function
        End If

        If (.Stats.GLD < goldGamble) Then
            fError = .Name & " no tiene el oro necesario, como mínimo necesita tener " & MIN_GOLD & " monedas de oro"
            Exit Function
        End If

        If (PointsGamble < 0) Then
            fError = .Name & " no tienes las copas necesarias."
            Exit Function
        End If

        If Not TieneObjetos(COPA_OBJ, PointsGamble, User_Index) Then
            fError = .Name & " no tienes las copas necesarias."
            Exit Function
        End If

        If (.Stats.ELV < 25) Then
            fError = .Name & " debe ser mayor a nivel 25!"
            Exit Function
        End If

        Check_User = True

    End With

End Function

Public Function CheckRespawnPlayer(ByVal rIndex As Integer) As Boolean

    CheckRespawnPlayer = Reto_List(rIndex).General_Rules.RespawnToggle

End Function

Public Sub Send_Reto(ByVal User_Index As Integer)

    Dim i As Long
    Dim j As Long

    Dim team_str As String
    Dim gamble_str As String

    With UserList(User_Index).sReto

        team_str = UserList(.TempStruct.Team_Array(0).User_Index(0)).Name & " y " & UserList(.TempStruct.Team_Array(0).User_Index(1)).Name & " vs " & UserList(.TempStruct.Team_Array(1).User_Index(0)).Name & " y " & UserList(.TempStruct.Team_Array(1).User_Index(1)).Name
        gamble_str = " apostando " & Format$(.TempStruct.General_Rules.Gold_Gamble, "#,###") & " monedas de oro" & IIf(.TempStruct.General_Rules.Points_Gamble > 0, ", " & .TempStruct.General_Rules.Points_Gamble & " copas", vbNullString)

        If (.TempStruct.General_Rules.Drop_Inv) Then
            gamble_str = " y los items del inventario"
        End If

        For i = 0 To 1
            For j = 0 To 1
                UserList(.TempStruct.Team_Array(i).User_Index(j)).sReto.Nick_Sender = UCase$(UserList(User_Index).Name)

                If (.TempStruct.Team_Array(i).User_Index(j) <> User_Index) Then
                    Call WriteConsoleMsg(.TempStruct.Team_Array(i).User_Index(j), "Solicitud de reto modalidad 2vs2 : " & team_str & " " & gamble_str & " para aceptar tipea /RETAR " & UCase$(UserList(User_Index).Name) & ".", FontTypeNames.FONTTYPE_GUILD)
                End If
            Next j
        Next i

        Call WriteConsoleMsg(User_Index, "Se han enviado las solicitudes.", FontTypeNames.FONTTYPE_GUILD)
        .AcceptLimit = 60

    End With

End Sub

Public Sub Disconnect_Reto(ByVal User_Index As Integer)

    Dim Team_Index As Integer
    Dim Team_Winner As Byte
    Dim Reto_Index As Integer

    Reto_Index = UserList(User_Index).sReto.Reto_Index
    Team_Index = Find_Team(User_Index, Reto_Index)

    If (Team_Index <> -1) Then
        Team_Winner = IIf(Team_Index = 1, 0, 1)
        Call Finish_Reto(UserList(User_Index).sReto.Reto_Index, Team_Winner)
    End If

End Sub

Public Sub closeOtherReto(ByVal UserIndex As Integer)

    Dim j As Long
    Dim i As Long
    Dim N As Integer
    Dim c As Boolean

    N = NameIndex(UserList(UserIndex).sReto.Nick_Sender)

    If (N > 0) Then

        For i = 0 To 1
            For j = 0 To 1
                With UserList(N).sReto.TempStruct.Team_Array(i)
                    If (.User_Index(j) = UserIndex) Then
                        c = True
                        Exit For
                    End If
                End With
            Next j
        Next i

        If c Then

            Dim UserName As String
            UserName = UCase$(UserList(N).Name)

            For i = 0 To 1
                For j = 0 To 1

                    With UserList(N).sReto.TempStruct.Team_Array(i)

                        If (.User_Index(j) > 0) Then
                            If StrComp(UCase$(UserList(.User_Index(j)).sReto.Nick_Sender), UserName) = 0 Then
                                Call WriteConsoleMsg(.User_Index(j), "El reto solicitado por " & UserName & " ha sido cancelado debido a la desconexión de un participante.", FontTypeNames.FONTTYPE_GUILD)
                            End If
                        End If

                    End With

                Next j
            Next i

        End If

    End If

End Sub

Public Sub Accept_Reto(ByVal User_Index As Integer, ByVal requestName As String)

    Dim sendIndex As Integer
    sendIndex = NameIndex(requestName)

    If (sendIndex = 0) Or (UCase$(requestName) <> UserList(User_Index).sReto.Nick_Sender) Then
        Call WriteConsoleMsg(User_Index, requestName & " no te está retando!!", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If

    Dim S As String

    If Check_AcceptUser(User_Index, UserList(sendIndex).sReto.TempStruct.General_Rules.Gold_Gamble, UserList(sendIndex).sReto.TempStruct.General_Rules.Points_Gamble, S) = False Then
        Call WriteConsoleMsg(User_Index, S, FontTypeNames.FONTTYPE_GUILD)
        Exit Sub
    End If

    ' ++ Paso toda la wea entonces entras xd
    UserList(sendIndex).sReto.Accept_Count = (UserList(sendIndex).sReto.Accept_Count + 1)

    Call Message_Reto(UserList(sendIndex).sReto.TempStruct, UserList(User_Index).Name & " aceptó el reto.")

    If (UserList(sendIndex).sReto.Accept_Count = 3) Then
        Call init_reto(sendIndex)
        Call Message_Reto(UserList(sendIndex).sReto.TempStruct, "Todos los participantes han aceptado el reto.")
    End If

    UserList(User_Index).sReto.AcceptedOK = True

End Sub

Public Sub ClearInfoAhreVilla(ByVal UserIndex As Integer)

    Dim i As Long, j As Long, User As Integer

    With UserList(UserIndex)

        For j = 0 To 1
            For i = 0 To 1

                User = .sReto.TempStruct.Team_Array(j).User_Index(i)

                ' @ Cui
                If (UserList(User).ConnID <> -1) Then

                    With UserList(User)

                        .sReto.Accept_Count = 0
                        .sReto.AcceptedOK = 0
                        .sReto.AcceptLimit = 0
                        .sReto.Nick_Sender = vbNullString
                        .sReto.Reto_Index = 0
                        .sReto.Reto_Used = 0

                    End With

                End If
            Next i
        Next j

    End With

End Sub

Private Sub init_reto(ByVal UserSendIndex As Integer)

    Dim Reto_Index As Integer
    Reto_Index = Get_Reto_Index()

    If (Reto_Index = -1) Then
        Call Message_Reto(UserList(UserSendIndex).sReto.TempStruct, "Reto cancelado, todas las arenas están ocupadas.")
        Exit Sub
    End If

    Reto_List(Reto_Index).General_Rules.RespawnToggle = UserList(UserSendIndex).sReto.TempStruct.General_Rules.RespawnToggle

    UserList(UserSendIndex).sReto.AcceptLimit = 0

    Reto_List(Reto_Index) = UserList(UserSendIndex).sReto.TempStruct

    Dim i As Long
    Dim j As Long
    Dim N As Integer

    ' @@ Descontamos el oro a los wachines.
    For i = 0 To 1
        For j = 0 To 1

            N = Reto_List(Reto_Index).Team_Array(i).User_Index(j)

            If (N > 0) Then
                UserList(N).Stats.GLD = UserList(N).Stats.GLD - Reto_List(Reto_Index).General_Rules.Gold_Gamble
                Call WriteUpdateGold(N)

                Call QuitarObjetos(COPA_OBJ, Reto_List(Reto_Index).General_Rules.Points_Gamble, N)
                Call QuitarObjetos(COPA_OBJ, Reto_List(Reto_Index).General_Rules.Points_Gamble, N)
            End If

        Next j
    Next i

    Reto_List(Reto_Index).Used_Ring = True
    Reto_List(Reto_Index).Count_Down = 6

    Call Warp_Teams(Reto_Index)

End Sub

Private Sub Warp_Teams(ByVal Reto_Index As Integer, _
                       Optional ByVal respawnUser As Boolean = False)

    With Reto_List(Reto_Index)

        Dim LoopC As Long

        Dim mPosX As Byte

        Dim mPosY As Byte

        Dim nUser As Integer

        .Count_Down = 6

        For LoopC = 0 To 1
            nUser = .Team_Array(0).User_Index(LoopC)

            If (nUser <> 0) Then
                If (UserList(nUser).ConnID <> -1) Then
                    mPosX = Get_Pos_X(Reto_Index, 1, CInt(LoopC))
                    mPosY = Get_Pos_Y(Reto_Index, 1, CInt(LoopC))

                    UserList(nUser).sReto.Reto_Used = True
                    UserList(nUser).sReto.Reto_Index = Reto_Index

                    Call WarpUserCharX(nUser, MAPA_ARENAS, mPosX, mPosY, True)
                    Call WritePauseToggle(nUser)

                    If (respawnUser) Then
                        If (UserList(nUser).flags.Muerto) Then
                            Call RevivirUsuario(nUser)
                        End If

                        UserList(nUser).Stats.MinHp = UserList(nUser).Stats.MaxHp
                        UserList(nUser).Stats.MinMAN = UserList(nUser).Stats.MaxMAN
                        UserList(nUser).Stats.MinSta = UserList(nUser).Stats.MaxSta

                        Call WriteUpdateUserStats(nUser)

                    End If

                Else

                    UserList(nUser).sReto.AcceptedOK = False

                End If

            End If

        Next LoopC

        For LoopC = 0 To 1
            nUser = .Team_Array(1).User_Index(LoopC)

            If (nUser <> 0) Then
                If (UserList(nUser).ConnID <> -1) Then
                    mPosX = Get_Pos_X(Reto_Index, 2, CInt(LoopC))
                    mPosY = Get_Pos_Y(Reto_Index, 2, CInt(LoopC))

                    UserList(nUser).sReto.Reto_Used = True
                    UserList(nUser).sReto.Reto_Index = Reto_Index

                    Call WarpUserCharX(nUser, MAPA_ARENAS, mPosX, mPosY, True)
                    Call WritePauseToggle(nUser)

                    If (respawnUser) Then

                        If (UserList(nUser).flags.Muerto) Then Call RevivirUsuario(nUser)

                        UserList(nUser).Stats.MinHp = UserList(nUser).Stats.MaxHp
                        UserList(nUser).Stats.MinMAN = UserList(nUser).Stats.MaxMAN
                        UserList(nUser).Stats.MinSta = UserList(nUser).Stats.MaxSta

                        Call WriteUpdateUserStats(nUser)

                    End If

                Else

                    UserList(nUser).sReto.AcceptedOK = False

                End If

            End If

        Next LoopC

    End With

End Sub

Private Sub Message_Reto(ByRef RetoStr As RetoStruct, ByRef sMessage As String)

    Dim i As Long
    Dim j As Long
    Dim U As Integer

    With RetoStr

        For i = 0 To 1
            For j = 0 To 1
                U = .Team_Array(i).User_Index(j)
                If (U <> 0) Then
                    If (UserList(U).ConnID <> -1) Then
                        Call WriteConsoleMsg(U, sMessage, FontTypeNames.FONTTYPE_GUILD)
                    End If
                End If
            Next j
        Next i

    End With

End Sub

Public Sub User_Die_Reto(ByVal User_Index As Integer)

    Dim Team_Index As Integer
    Dim user_slot As Integer
    Dim Other_User As Integer
    Dim Reto_Index As Integer

    On Error GoTo User_Die_Reto_Error

    Reto_Index = UserList(User_Index).sReto.Reto_Index

    Team_Index = Find_Team(User_Index, Reto_Index)

    If (Team_Index <> -1) Then
        user_slot = Find_User(Team_Index, User_Index, Reto_Index)
    Else
        Exit Sub
    End If

    If (user_slot = -1) Then Exit Sub

    Other_User = IIf(user_slot = 0, 1, 0)
    Other_User = Reto_List(Reto_Index).Team_Array(Team_Index).User_Index(Other_User)

    'is dead?
    If (Other_User) Then
        If UserList(Other_User).flags.Muerto Then Call Team_Winner(Reto_Index, IIf(Team_Index = 0, 1, 0))
    Else
        Call Team_Winner(Reto_Index, IIf(Team_Index = 0, 1, 0))
    End If

    Exit Sub

User_Die_Reto_Error:

    Call LogError("Error " & Err.Number & " (" & Err.description & ") in procedure User_Die_Reto of Módulo m_Retos2vs2 " & Erl & ".")

End Sub

Public Function Find_Team(ByVal User_Index As Integer, _
                          ByVal Reto_Index As Integer) As Integer

    Dim i As Long
    Dim j As Long

    For i = 0 To 1
        For j = 0 To 1
            If Reto_List(Reto_Index).Team_Array(i).User_Index(j) = User_Index Then
                Find_Team = i
                Exit Function
            End If
        Next j
    Next i

    Find_Team = -1

End Function

Private Function Find_User(ByVal Team_Index As Integer, _
                           ByVal User_Index As Integer, _
                           ByVal Reto_Index As Integer) As Integer

    Dim i As Long

    For i = 0 To 1
        If Reto_List(Reto_Index).Team_Array(Team_Index).User_Index(i) = User_Index Then
            Find_User = i
            Exit Function
        End If
    Next i

    Find_User = -1

End Function

Private Sub Team_Winner(ByVal Reto_Index As Integer, ByVal Team_Winner As Byte)

    On Error GoTo Team_Winner_Error

    With Reto_List(Reto_Index)

        .Team_Array(Team_Winner).Round_Count = (.Team_Array(Team_Winner).Round_Count + 1)

        If (.Team_Array(Team_Winner).Round_Count = 2) Then
            Call Finish_Reto(Reto_Index, Team_Winner)
        Else
            Call Respawn_Reto(Reto_Index, Team_Winner)
        End If

    End With

    Exit Sub

Team_Winner_Error:

    Call LogError("Error " & Err.Number & " (" & Err.description & ") in procedure Team_Winner of Módulo m_Retos2vs2 " & Erl & ".")

End Sub

Private Sub Respawn_Reto(ByVal Reto_Index As Integer, ByVal Team_Winner As Integer)

    Dim loopX As Long
    Dim LoopC As Long
    Dim mStr As String
    Dim Index As Integer

    On Error GoTo Respawn_Reto_Error

   With Reto_List(Reto_Index)

        mStr = "El equipo " & CStr(Team_Winner + 1) & " gana este duelo." & vbNewLine & "Resultado parcial : " & CStr(.Team_Array(0).Round_Count) & "-" & CStr(.Team_Array(1).Round_Count)

        For loopX = 0 To 1
            For LoopC = 0 To 1
                Index = .Team_Array(loopX).User_Index(LoopC)
 
                If (Index <> 0) Then
                    If UserList(Index).ConnID <> -1 Then
                        Call WriteConsoleMsg(Index, mStr, FontTypeNames.FONTTYPE_GUILD)
                        Call WriteConsoleMsg(Index, "El siguiente round iniciará en 3 segundos.", FontTypeNames.FONTTYPE_GUILD)
                    End If
                End If

            Next LoopC
        Next loopX

        .NextRoundCount = 3

    End With

    Exit Sub

Respawn_Reto_Error:

    Call LogError("Error " & Err.Number & " (" & Err.description & ") in procedure Respawn_Reto of Módulo m_Retos2vs2" & Erl & ".")

End Sub

Private Sub Finish_Reto(ByVal Reto_Index As Integer, ByVal Team_Winner As Byte)

    With Reto_List(Reto_Index)

        On Error GoTo Errhandler

        Dim RetoMessage As String
        Dim Team_Looser As Byte
        Dim Temp_Index As Integer

        RetoMessage = Get_Reto_Message(Reto_Index)
        Team_Looser = IIf(Team_Winner = 0, 1, 0)

        RetoMessage = RetoMessage & "Retos 2vs2 " & UserList(.Team_Array(Team_Winner).User_Index(0)).Name & " y " & UserList(.Team_Array(Team_Winner).User_Index(1)).Name & " vs " & _
                      UserList(.Team_Array(Team_Looser).User_Index(0)).Name & " y " & UserList(.Team_Array(Team_Looser).User_Index(1)).Name & ". Ganador el equipo de " & _
                      UserList(.Team_Array(Team_Winner).User_Index(0)).Name & " y " & UserList(.Team_Array(Team_Winner).User_Index(1)).Name & ". Apuesta por " & .General_Rules.Gold_Gamble & " monedas de oro" & IIf(.General_Rules.Points_Gamble > 0, ", " & .General_Rules.Points_Gamble & " copas", vbNullString) & IIf(.General_Rules.Drop_Inv, " y los items.", ".")

        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(RetoMessage, FontTypeNames.FONTTYPE_INFO))

        Dim LoopC As Long
        Dim ByDrop As Boolean
        Dim ByGold As Long
        Dim ByPoints As Long

        ByDrop = (.General_Rules.Drop_Inv = True)
        ByGold = (.General_Rules.Gold_Gamble * 2)
        ByPoints = (.General_Rules.Points_Gamble * 2)

        With .Team_Array(Team_Looser)

        For LoopC = 0 To 1
            Temp_Index = .User_Index(LoopC)

            UserList(Temp_Index).sReto.Reto_Used = False
            UserList(Temp_Index).sReto.AcceptedOK = False

            If (ByDrop) Then Call TirarTodosLosItems(Temp_Index)

                Call WarpUserCharX(Temp_Index, 104, 56 + LoopC, 40, True) '104, 56, 39

                UserList(Temp_Index).sReto.Nick_Sender = vbNullString
                UserList(Temp_Index).sReto.Reto_Index = 0
                UserList(Temp_Index).Stats.RetosPerdidos = UserList(Temp_Index).Stats.RetosPerdidos + 1
            Next LoopC

        End With

        With .Team_Array(Team_Winner)

            For LoopC = 0 To 1
                Temp_Index = .User_Index(LoopC)

                UserList(Temp_Index).sReto.Reto_Used = False
                UserList(Temp_Index).sReto.AcceptedOK = False

                If (ByDrop) Then
                    UserList(Temp_Index).sReto.Return_City = 15
                    Call WriteConsoleMsg(Temp_Index, "Regresarás a tu hogar en 15 segundos.", FontTypeNames.FONTTYPE_GUILD)
                Else
                    Call WarpUserCharX(Temp_Index, 104, 56 + LoopC, 40, True) '104, 56, 39 mapa.
                End If

                Call DarPremioEvento(Temp_Index, ByGold, ByPoints)

                UserList(Temp_Index).sReto.Nick_Sender = vbNullString
                UserList(Temp_Index).sReto.Reto_Index = 0
                UserList(Temp_Index).Stats.RetosGanados = UserList(Temp_Index).Stats.RetosGanados + 1
            Next LoopC

        End With

        Call Clear_Data(Reto_Index)

    End With

    Exit Sub

Errhandler:

    Call LogError("Error en Finish_Reto de 2vs2 en " & Erl & ". Err " & Err.Number & " " & Err.description)

End Sub

Private Sub Clear_Data(ByVal Reto_Index As Integer)

    On Error GoTo Clear_data_Err

    With Reto_List(Reto_Index)
        .Count_Down = 0

        With .General_Rules
            .Drop_Inv = False
            .Gold_Gamble = 0
            .RespawnToggle = False
        End With

            .Used_Ring = False

        Dim i As Long

        For i = 0 To 1
            .Team_Array(i).User_Index(0) = 0
            .Team_Array(i).User_Index(1) = 0

            .Team_Array(i).Round_Count = 0
        Next i

    End With

    Exit Sub

Clear_data_Err:

    Call LogError("Error en Clear_data. Error: " & Err.Number & " - " & Err.description)

End Sub

Private Function Get_Reto_Message(ByVal Reto_Index As Integer) As String

    Dim TempStr As String
    Dim TempCO As String
    Dim TempUser As Integer

    With Reto_List(Reto_Index)

        TempStr = "Retos "

        With .Team_Array(0)

            TempUser = .User_Index(0)

            If (TempUser <> 0) Then
                If UserList(TempUser).ConnID <> -1 Then
                    TempStr = TempStr & UserList(TempUser).Name
                End If
            End If

            TempUser = .User_Index(1)

            If (TempUser <> 0) Then
                If UserList(TempUser).ConnID <> -1 Then
                    TempStr = TempStr & " y " & UserList(TempUser).Name
                End If
            End If

        End With

        With .Team_Array(1)
            TempUser = .User_Index(0)

            If (TempUser <> 0) Then
                If UserList(TempUser).ConnID <> -1 Then
                    TempStr = TempStr & " vs " & UserList(TempUser).Name
                End If
            End If

            TempUser = .User_Index(1)

            If (TempUser <> 0) Then
                If UserList(TempUser).ConnID <> -1 Then
                    TempStr = TempStr & " y " & UserList(TempUser).Name
                End If
            End If

        End With

        With .General_Rules

            TempStr = TempStr & " con apuesta de " & Format$(.Gold_Gamble, "#,###") & " monedas de oro"

            If (.Drop_Inv) Then
                TempStr = TempStr & " y los items del inventario"
            End If

        End With

    End With

    TempStr = TempStr & TempCO

End Function

Public Function Get_Pos_X(ByVal Ring_Index As Integer, _
                          ByVal Team_Index As Integer, _
                          ByVal User_Index As Integer) As Integer

    Dim EndPos As Integer
    EndPos = RetoPos(Ring_Index, Team_Index, User_Index + 1).X
    Get_Pos_X = EndPos

End Function

Public Function Get_Pos_Y(ByVal Ring_Index As Integer, _
                          ByVal Team_Index As Integer, _
                          ByVal User_Index As Integer) As Integer

    Dim EndPos As Integer
    EndPos = RetoPos(Ring_Index, Team_Index, User_Index + 1).Y
    Get_Pos_Y = EndPos

End Function

Public Sub Retos2vs2Load()

    Dim nArenas As Integer

    Dim Leer As New clsIniManager
    Set Leer = New clsIniManager

    Leer.Initialize DatPath & "Retos2vs2.ini"

    nArenas = val(Leer.GetValue("INIT", "Arenas"))

    If (nArenas = 0) Then Exit Sub

    ReDim m_Retos2vs2.RetoPos(1 To nArenas, 1 To 2, 1 To 2) As m_Retos2vs2.RetoPosStruct
    ReDim m_Retos2vs2.Reto_List(1 To nArenas) As m_Retos2vs2.RetoStruct

    Dim i As Long
    Dim j As Long
    Dim p As Long
    Dim S As String

    For i = 1 To nArenas
        For j = 1 To 2
            For p = 1 To 2
                S = Leer.GetValue("ARENA" & CStr(i), "Equipo" & CStr(j) & "Jugador" & CStr(p))
                m_Retos2vs2.RetoPos(i, j, p).X = val(ReadField(1, S, 45))
                m_Retos2vs2.RetoPos(i, j, p).Y = val(ReadField(2, S, 45))
            Next p
        Next j
    Next i

    Set Leer = Nothing

End Sub


