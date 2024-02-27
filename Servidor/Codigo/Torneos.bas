Attribute VB_Name = "Torneos"
Option Explicit
' Codigo: Torneos Automaticos 100%
' Autor: Joan Calder�n - SaturoS.
Public Torneo_Activo As Boolean
Public Torneo_Esperando As Boolean
Private Torneo_Rondas As Integer
Private Torneo_Luchadores() As Integer
 
Private Const mapatorneo As Integer = 197
' esquinas superior isquierda del ring
Private Const esquina1x As Integer = 28
Private Const esquina1y As Integer = 21
' esquina inferior derecha del ring
Private Const esquina2x As Integer = 53
Private Const esquina2y As Integer = 37
' Donde esperan los tios
Private Const esperax As Integer = 77
Private Const esperay As Integer = 80
' Mapa desconecta
Private Const mapa_fuera As Integer = 132
Private Const fueraesperay As Integer = 19
Private Const fueraesperax As Integer = 53
 ' estas son las pocisiones de las 2 esquinas de la zona de espera, en su mapa tienen que tener en la misma posicion las 2 esquinas.
Private Const X1 As Integer = 65
Private Const X2 As Integer = 89
Private Const Y1 As Integer = 78
Private Const Y2 As Integer = 88
 
Sub Torneoauto_Cancela()
On Error GoTo errorh:
    If (Not Torneo_Activo And Not Torneo_Esperando) Then Exit Sub
    Torneo_Activo = False
    Torneo_Esperando = False
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Torneo Autom�tico> Se cancel� por falta de participantes.", FontTypeNames.FONTTYPE_SERVER))
    Dim i As Integer
     For i = LBound(Torneo_Luchadores) To UBound(Torneo_Luchadores)
                If (Torneo_Luchadores(i) <> -1) Then
                 Dim NuevaPos As WorldPos
                  Dim FuturePos As WorldPos
                    FuturePos.map = mapa_fuera
                    FuturePos.X = fueraesperax: FuturePos.Y = fueraesperay
                    Call ClosestLegalPos(FuturePos, NuevaPos)
                    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(Torneo_Luchadores(i), NuevaPos.map, NuevaPos.X, NuevaPos.Y, True)
                      UserList(Torneo_Luchadores(i)).flags.automatico = False
                End If
        Next i
errorh:
End Sub
Sub Rondas_Cancela()
On Error GoTo errorh
    If (Not Torneo_Activo And Not Torneo_Esperando) Then Exit Sub
    Torneo_Activo = False
    Torneo_Esperando = False
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Torneo Autom�tico> Cancelado por un Game Master.", FontTypeNames.FONTTYPE_SERVER))
    Dim i As Integer
    For i = LBound(Torneo_Luchadores) To UBound(Torneo_Luchadores)
                If (Torneo_Luchadores(i) <> -1) Then
                        Dim NuevaPos As WorldPos
                  Dim FuturePos As WorldPos
                    FuturePos.map = mapa_fuera
                    FuturePos.X = fueraesperax: FuturePos.Y = fueraesperay
                    Call ClosestLegalPos(FuturePos, NuevaPos)
                    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(Torneo_Luchadores(i), NuevaPos.map, NuevaPos.X, NuevaPos.Y, True)
                    UserList(Torneo_Luchadores(i)).flags.automatico = False
                End If
        Next i
errorh:
End Sub
Sub Rondas_UsuarioMuere(ByVal UserIndex As Integer, Optional Real As Boolean = True, Optional CambioMapa As Boolean = False)
On Error GoTo rondas_usuariomuere_errorh
        Dim i As Integer, Pos As Integer, j As Integer
        Dim combate As Integer, LI1 As Integer, LI2 As Integer
        Dim UI1 As Integer, UI2 As Integer
If (Not Torneo_Activo) Then
                Exit Sub
            ElseIf (Torneo_Activo And Torneo_Esperando) Then
                For i = LBound(Torneo_Luchadores) To UBound(Torneo_Luchadores)
                    If (Torneo_Luchadores(i) = UserIndex) Then
                        Torneo_Luchadores(i) = -1
                        Call WarpUserChar(UserIndex, mapa_fuera, fueraesperay, fueraesperax, True)
                         UserList(UserIndex).flags.automatico = False
                        Exit Sub
                    End If
                Next i
                Exit Sub
            End If
 
        For Pos = LBound(Torneo_Luchadores) To UBound(Torneo_Luchadores)
                If (Torneo_Luchadores(Pos) = UserIndex) Then Exit For
        Next Pos
 
        ' si no lo ha encontrado
        If (Torneo_Luchadores(Pos) <> UserIndex) Then Exit Sub
       
 '  Ojo con esta parte, aqui es donde verifica si el usuario esta en la posicion de espera del torneo, en estas cordenadas tienen que fijarse al crear su Mapa de torneos.
 
If UserList(UserIndex).Pos.X >= X1 And UserList(UserIndex).Pos.X <= X2 And UserList(UserIndex).Pos.Y >= Y1 And UserList(UserIndex).Pos.Y <= Y2 Then
Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Torneo Autom�tico> " & UserList(UserIndex).Name & " se fue del torneo mientras esperaba pelear!", FontTypeNames.FONTTYPE_SERVER))
Call WarpUserChar(UserIndex, mapa_fuera, fueraesperax, fueraesperay, True)
UserList(UserIndex).flags.automatico = False
Torneo_Luchadores(Pos) = -1
Exit Sub
End If
 
        combate = 1 + (Pos - 1) \ 2
 
        'ponemos li1 y li2 (luchador index) de los que combatian
        LI1 = 2 * (combate - 1) + 1
        LI2 = LI1 + 1
 
        'se informa a la gente
        If (Real) Then
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Torneo Autom�tico> " & UserList(UserIndex).Name & " pierde el combate!", FontTypeNames.FONTTYPE_SERVER))
        Else
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Torneo Autom�tico> " & UserList(UserIndex).Name & " se fue del combate!", FontTypeNames.FONTTYPE_SERVER))
        End If
 
        'se le teleporta fuera si murio
        If (Real) Then
                Call WarpUserChar(UserIndex, mapa_fuera, fueraesperax, fueraesperay, True)
                 UserList(UserIndex).flags.automatico = False
        ElseIf (Not CambioMapa) Then
             
                 Call WarpUserChar(UserIndex, mapa_fuera, fueraesperax, fueraesperay, True)
                  UserList(UserIndex).flags.automatico = False
        End If
 
        'se le borra de la lista y se mueve el segundo a li1
        If (Torneo_Luchadores(LI1) = UserIndex) Then
                Torneo_Luchadores(LI1) = Torneo_Luchadores(LI2) 'cambiamos slot
                Torneo_Luchadores(LI2) = -1
        Else
                Torneo_Luchadores(LI2) = -1
        End If
 
    'si es la ultima ronda
    If (Torneo_Rondas = 1) Then
        Call WarpUserChar(Torneo_Luchadores(LI1), mapa_fuera, 51, 51, True)
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Torneo Autom�tico> " & UserList(UserIndex).Name & " pierde el combate!", FontTypeNames.FONTTYPE_SERVER))
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Torneo Autom�tico> Ganador del torneo: " & UserList(Torneo_Luchadores(LI1)).Name, FontTypeNames.FONTTYPE_SERVER))
        UserList(Torneo_Luchadores(LI1)).Stats.GLD = UserList(Torneo_Luchadores(LI1)).Stats.GLD + 1000000
        Call WriteUpdateGold(Torneo_Luchadores(LI1))
        UserList(Torneo_Luchadores(LI1)).flags.automatico = False
        Torneo_Activo = False
        
        Exit Sub
    Else
        'a su compa�ero se le teleporta dentro, condicional por seguridad
        Call WarpUserChar(Torneo_Luchadores(LI1), mapatorneo, esperax, esperay, True)
    End If
 
               
        'si es el ultimo combate de la ronda
        If (2 ^ Torneo_Rondas = 2 * combate) Then
                
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Torneo Autom�tico> Siguiente ronda!", FontTypeNames.FONTTYPE_SERVER))
                Torneo_Rondas = Torneo_Rondas - 1
 
        'antes de llamar a la proxima ronda hay q copiar a los tipos
        For i = 1 To 2 ^ Torneo_Rondas
                UI1 = Torneo_Luchadores(2 * (i - 1) + 1)
                UI2 = Torneo_Luchadores(2 * i)
                If (UI1 = -1) Then UI1 = UI2
                Torneo_Luchadores(i) = UI1
        Next i
ReDim Preserve Torneo_Luchadores(1 To 2 ^ Torneo_Rondas) As Integer
        Call Rondas_Combate(1)
        Exit Sub
        End If
 
        'vamos al siguiente combate
        Call Rondas_Combate(combate + 1)
rondas_usuariomuere_errorh:
 
End Sub
 
 
 
Sub Rondas_UsuarioDesconecta(ByVal UserIndex As Integer)
On Error GoTo errorh
Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Torneo Autom�tico> " & UserList(UserIndex).Name & " ha desconectado en torneo autom�tico, se le penaliza con 500K!!", FontTypeNames.FONTTYPE_SERVER))
 If UserList(UserIndex).Stats.GLD >= 500000 Then
UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - 500000
End If
Call Rondas_UsuarioMuere(UserIndex, False, False)
errorh:
End Sub
 
 
 
Sub Rondas_UsuarioCambiamapa(ByVal UserIndex As Integer)
On Error GoTo errorh
        Call Rondas_UsuarioMuere(UserIndex, False, True)
errorh:
End Sub
 
Sub torneos_auto(ByVal rondas As Integer)
On Error GoTo errorh
If (Torneo_Activo) Then
               
                Exit Sub
        End If
        
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Torneo Autom�tico> Esta empezando un nuevo torneo 1VS1 de " & val(2 ^ rondas) & " participantes!! para participar teclea /PARTICIPAR - (No cae inventario)", FontTypeNames.FONTTYPE_SERVER))
        Torneo_Rondas = rondas
        Torneo_Activo = True
        Torneo_Esperando = True
 
        ReDim Torneo_Luchadores(1 To 2 ^ rondas) As Integer
        Dim i As Integer
        For i = LBound(Torneo_Luchadores) To UBound(Torneo_Luchadores)
                Torneo_Luchadores(i) = -1
        Next i
errorh:
End Sub
 
Sub Torneos_Inicia(ByVal UserIndex As Integer, ByVal rondas As Integer)
On Error GoTo errorh
 
        If (Torneo_Activo) Then
                Call WriteConsoleMsg(UserIndex, "Ya hay un torneo en curso", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
        End If
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Torneo Autom�tico> Esta empezando un nuevo torneo 1VS1 de " & val(2 ^ rondas) & " participantes!! para participar teclea /PARTICIPAR - (No cae inventario)", FontTypeNames.FONTTYPE_SERVER))
 
        Torneo_Rondas = rondas
        Torneo_Activo = True
        Torneo_Esperando = True
 
        ReDim Torneo_Luchadores(1 To 2 ^ rondas) As Integer
        Dim i As Integer
        For i = LBound(Torneo_Luchadores) To UBound(Torneo_Luchadores)
                Torneo_Luchadores(i) = -1
        Next i
errorh:
End Sub
 
 
 
Sub Torneos_Entra(ByVal UserIndex As Integer)
On Error GoTo errorh
        Dim i As Integer
       
        If (Not Torneo_Activo) Then
                Call WriteConsoleMsg(UserIndex, "�No hay ning�n torneo en curso!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
        End If
       
        If (Not Torneo_Esperando) Then
                Call WriteConsoleMsg(UserIndex, "Cupos llenos.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
        End If
       
        For i = LBound(Torneo_Luchadores) To UBound(Torneo_Luchadores)
                If (Torneo_Luchadores(i) = UserIndex) Then
                        Call WriteConsoleMsg(UserIndex, "�Ya est�s dentro del torneo!", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                End If
        Next i
 
        For i = LBound(Torneo_Luchadores) To UBound(Torneo_Luchadores)
        If (Torneo_Luchadores(i) = -1) Then
                Torneo_Luchadores(i) = UserIndex
                 Dim NuevaPos As WorldPos
                  Dim FuturePos As WorldPos
                    FuturePos.map = mapatorneo
                    FuturePos.X = esperax: FuturePos.Y = esperay
                    Call ClosestLegalPos(FuturePos, NuevaPos)
                   
                    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then Call WarpUserChar(Torneo_Luchadores(i), NuevaPos.map, NuevaPos.X, NuevaPos.Y, True)
                 UserList(Torneo_Luchadores(i)).flags.automatico = True
                 
                Call WriteConsoleMsg(UserIndex, "�Has ingresado al torneo!", FontTypeNames.FONTTYPE_INFO)
               Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Torneo Autom�tico> El personaje " & UserList(UserIndex).Name & " ingres� al torneo.", FontTypeNames.FONTTYPE_SERVER))
                If (i = UBound(Torneo_Luchadores)) Then
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Torneo Autom�tico> Se da inicio al torneo.", FontTypeNames.FONTTYPE_SERVER))
                Torneo_Esperando = False
                Call Rondas_Combate(1)
     
                End If
                  Exit Sub
        End If
        Next i
errorh:
End Sub
 
 
Sub Rondas_Combate(combate As Integer)
On Error GoTo errorh
Dim UI1 As Integer, UI2 As Integer
    UI1 = Torneo_Luchadores(2 * (combate - 1) + 1)
    UI2 = Torneo_Luchadores(2 * combate)
   
    If (UI2 = -1) Then
        UI2 = Torneo_Luchadores(2 * (combate - 1) + 1)
        UI1 = Torneo_Luchadores(2 * combate)
    End If
   
    If (UI1 = -1) Then
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Torneo Autom�tico> Combate anulado por la desconexi�n de uno de los dos participantes.", FontTypeNames.FONTTYPE_SERVER))
        If (Torneo_Rondas = 1) Then
            If (UI2 <> -1) Then
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Torneo Autom�tico> Torneo terminado, ganador del torneo por eliminaci�n " & UserList(UI2).Name & ".", FontTypeNames.FONTTYPE_SERVER))
                UserList(UI2).flags.automatico = False
                ' dale_recompensa()
                Torneo_Activo = False
                Exit Sub
            End If
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Torneo Autom�tico> No hay ganador del evento por la desconexi�n de todos sus participantes.", FontTypeNames.FONTTYPE_SERVER))
            Exit Sub
        End If
        If (UI2 <> -1) Then _
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Torneo Autom�tico> El usuario " & UserList(UI2).Name & " pas� a la siguiente ronda.", FontTypeNames.FONTTYPE_SERVER))
            
        If (2 ^ Torneo_Rondas = 2 * combate) Then
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Torneo Autom�tico> Siguiente ronda.", FontTypeNames.FONTTYPE_SERVER))
            Torneo_Rondas = Torneo_Rondas - 1
            'antes de llamar a la proxima ronda hay q copiar a los tipos
            Dim i As Integer, j As Integer
            For i = 1 To 2 ^ Torneo_Rondas
                UI1 = Torneo_Luchadores(2 * (i - 1) + 1)
                UI2 = Torneo_Luchadores(2 * i)
                If (UI1 = -1) Then UI1 = UI2
                Torneo_Luchadores(i) = UI1
            Next i
            ReDim Preserve Torneo_Luchadores(1 To 2 ^ Torneo_Rondas) As Integer
            Call Rondas_Combate(1)
            Exit Sub
        End If
        Call Rondas_Combate(combate + 1)
        Exit Sub
    End If
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Torneo Autom�tico> " & UserList(UI1).Name & " vs. " & UserList(UI2).Name & ". Esquinas!!, Peleen!", FontTypeNames.FONTTYPE_SERVER))
    Call WarpUserChar(UI1, mapatorneo, esquina1x, esquina1y, True)
    Call WarpUserChar(UI2, mapatorneo, esquina2x, esquina2y, True)
errorh:
End Sub

