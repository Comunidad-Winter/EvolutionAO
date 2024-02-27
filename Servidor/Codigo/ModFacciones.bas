Attribute VB_Name = "ModFacciones"
'Argentum Online 0.12.2
'Copyright (C) 2002 Márquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez


Option Explicit


Public ArmaduraImperial1 As Integer
Public ArmaduraImperial2 As Integer
Public ArmaduraImperial3 As Integer
Public TunicaMagoImperial As Integer
Public TunicaMagoImperialEnanos As Integer
Public ArmaduraCaos1 As Integer
Public ArmaduraCaos2 As Integer
Public ArmaduraCaos3 As Integer
Public TunicaMagoCaos As Integer
Public TunicaMagoCaosEnanos As Integer

Public VestimentaImperialHumano As Integer
Public VestimentaImperialEnano As Integer
Public TunicaConspicuaHumano As Integer
Public TunicaConspicuaEnano As Integer
Public ArmaduraNobilisimaHumano As Integer
Public ArmaduraNobilisimaEnano As Integer
Public ArmaduraGranSacerdote As Integer

Public VestimentaLegionHumano As Integer
Public VestimentaLegionEnano As Integer
Public TunicaLobregaHumano As Integer
Public TunicaLobregaEnano As Integer
Public TunicaEgregiaHumano As Integer
Public TunicaEgregiaEnano As Integer
Public SacerdoteDemoniaco As Integer



Public Const NUM_RANGOS_FACCION As Integer = 15
Private Const NUM_DEF_FACCION_ARMOURS As Byte = 3

Public Enum eTipoDefArmors
    ieBaja
    ieMedia
    ieAlta
End Enum

Public Type tFaccionArmaduras
    Armada(NUM_DEF_FACCION_ARMOURS - 1) As Integer
    Caos(NUM_DEF_FACCION_ARMOURS - 1) As Integer
End Type

' Matriz que contiene las armaduras faccionarias segun raza, clase, faccion y defensa de armadura
Public ArmadurasFaccion(1 To NUMCLASES, 1 To NUMRAZAS) As tFaccionArmaduras

' Contiene la cantidad de exp otorgada cada vez que aumenta el rango
Public RecompensaFacciones(NUM_RANGOS_FACCION) As Long

Private Function GetArmourAmount(ByVal Rango As Integer, ByVal TipoDef As eTipoDefArmors) As Integer
    '***************************************************
    'Autor: ZaMa
    'Last Modification: 15/04/2010
    'Returns the amount of armours to give, depending on the specified rank
    '***************************************************

    Select Case TipoDef

        Case eTipoDefArmors.ieBaja
            GetArmourAmount = 20 / (Rango + 1)

        Case eTipoDefArmors.ieMedia
            GetArmourAmount = Rango * 2 / MaximoInt((Rango - 4), 1)

        Case eTipoDefArmors.ieAlta
            GetArmourAmount = Rango * 1.35

    End Select

End Function

Private Sub GiveFactionArmours(ByVal UserIndex As Integer, ByVal IsCaos As Boolean)
    '***************************************************
    'Autor: ZaMa
    'Last Modification: 15/04/2010
    'Gives faction armours to user
    '***************************************************

    Dim ObjArmour As Obj
    Dim Rango As Integer

    With UserList(UserIndex)

        Rango = val(IIf(IsCaos, .Faccion.RecompensasCaos, .Faccion.RecompensasReal)) + 1


        ' Entrego armaduras de defensa baja
        ObjArmour.Amount = GetArmourAmount(Rango, eTipoDefArmors.ieBaja)

        If IsCaos Then
            ObjArmour.ObjIndex = ArmadurasFaccion(.clase, .raza).Caos(eTipoDefArmors.ieBaja)
        Else
            ObjArmour.ObjIndex = ArmadurasFaccion(.clase, .raza).Armada(eTipoDefArmors.ieBaja)
        End If

        If Not MeterItemEnInventario(UserIndex, ObjArmour) Then
            Call TirarItemAlPiso(.Pos, ObjArmour)
        End If


        ' Entrego armaduras de defensa media
        ObjArmour.Amount = GetArmourAmount(Rango, eTipoDefArmors.ieMedia)

        If IsCaos Then
            ObjArmour.ObjIndex = ArmadurasFaccion(.clase, .raza).Caos(eTipoDefArmors.ieMedia)
        Else
            ObjArmour.ObjIndex = ArmadurasFaccion(.clase, .raza).Armada(eTipoDefArmors.ieMedia)
        End If

        If Not MeterItemEnInventario(UserIndex, ObjArmour) Then
            Call TirarItemAlPiso(.Pos, ObjArmour)
        End If


        ' Entrego armaduras de defensa alta
        ObjArmour.Amount = GetArmourAmount(Rango, eTipoDefArmors.ieAlta)

        If IsCaos Then
            ObjArmour.ObjIndex = ArmadurasFaccion(.clase, .raza).Caos(eTipoDefArmors.ieAlta)
        Else
            ObjArmour.ObjIndex = ArmadurasFaccion(.clase, .raza).Armada(eTipoDefArmors.ieAlta)
        End If

        If Not MeterItemEnInventario(UserIndex, ObjArmour) Then
            Call TirarItemAlPiso(.Pos, ObjArmour)
        End If

    End With

End Sub

Public Sub GiveExpReward(ByVal UserIndex As Integer, ByVal Rango As Long)
    '***************************************************
    'Autor: ZaMa
    'Last Modification: 15/04/2010
    'Gives reward exp to user
    '***************************************************

    Dim GivenExp As Long

    With UserList(UserIndex)

        GivenExp = RecompensaFacciones(Rango)

        .Stats.Exp = .Stats.Exp + GivenExp

        If .Stats.Exp > MAXEXP Then .Stats.Exp = MAXEXP

        Call WriteConsoleMsg(UserIndex, "Has sido recompensado con " & GivenExp & " puntos de experiencia.", FontTypeNames.FONTTYPE_FIGHT)

        Call CheckUserLevel(UserIndex)

    End With

End Sub

Public Sub EnlistarArmadaReal(ByVal UserIndex As Integer)
    '***************************************************
    'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
    'Last Modification: 15/04/2010
    'Handles the entrance of users to the "Armada Real"
    '15/03/2009: ZaMa - No se puede enlistar el fundador de un clan con alineación neutral.
    '27/11/2009: ZaMa - Ahora no se puede enlistar un miembro de un clan neutro, por ende saque la antifaccion.
    '15/04/2010: ZaMa - Cambio en recompensas iniciales.
    '***************************************************

    With UserList(UserIndex)
        If .Faccion.ArmadaReal = 1 Then
            Call WriteChatOverHead(UserIndex, "¡¡¡Ya perteneces a las tropas reales!!! Ve a combatir criminales.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub
        End If

        If .Faccion.FuerzasCaos = 1 Then
            Call WriteChatOverHead(UserIndex, "¡¡¡Maldito insolente!!! Vete de aquí seguidor de las sombras.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub
        End If

        If criminal(UserIndex) Then
            Call WriteChatOverHead(UserIndex, "¡¡¡No se permiten criminales en el ejército real!!!", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub
        End If

        If .Faccion.CriminalesMatados < 30 Then
            Call WriteChatOverHead(UserIndex, "Para unirte a nuestras fuerzas debes matar al menos 30 criminales, sólo has matado " & .Faccion.CriminalesMatados & ".", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub
        End If

        If .Stats.ELV < 25 Then
            Call WriteChatOverHead(UserIndex, "¡¡¡Para unirte a nuestras fuerzas debes ser al menos de nivel 25!!!", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub
        End If

        If .Faccion.CiudadanosMatados > 0 Then
            Call WriteChatOverHead(UserIndex, "¡Has asesinado gente inocente, no aceptamos asesinos en las tropas reales!", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub
        End If

        If .Faccion.Reenlistadas > 4 Then
            Call WriteChatOverHead(UserIndex, "¡Has sido expulsado de las fuerzas reales demasiadas veces!", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub
        End If

        'Enlistarse Armada con 10.000 noble.
        If .Reputacion.NobleRep < 10000 Then
            Call WriteChatOverHead(UserIndex, "Necesitas ser aún más noble para integrar el ejército real, sólo tienes " & .Reputacion.NobleRep & "/10.000 puntos de nobleza", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub
        End If

        If .GuildIndex > 0 Then
            If modGuilds.GuildAlignment(.GuildIndex) = "Neutral" Then
                Call WriteChatOverHead(UserIndex, "¡¡¡Perteneces a un clan neutro, sal de él si quieres unirte a nuestras fuerzas!!!", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                Exit Sub
            End If
        End If

        .Faccion.ArmadaReal = 1
        .Faccion.Reenlistadas = .Faccion.Reenlistadas + 1

        Call WriteChatOverHead(UserIndex, "¡¡¡Bienvenido al ejército real!!! Aquí tienes tus vestimentas. Cumple bien tu labor exterminando criminales y me encargaré de recompensarte.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)

        ' TODO: Dejo esta variable por ahora, pero con chequear las reenlistadas deberia ser suficiente :S
        If .Faccion.RecibioArmaduraReal = 0 Then

            Call GiveFactionArmours(UserIndex, False)
            Call GiveExpReward(UserIndex, 0)

            .Faccion.RecibioArmaduraReal = 1
            .Faccion.NivelIngreso = .Stats.ELV
            .Faccion.FechaIngreso = Date
            'Esto por ahora es inútil, siempre va a ser cero, pero bueno, despues va a servir.
            .Faccion.MatadosIngreso = .Faccion.CiudadanosMatados

            .Faccion.RecibioExpInicialReal = 1
            .Faccion.RecompensasReal = 0
            .Faccion.NextRecompensa = 10    ' acordate que este numero te manda al CASE. sisi



        End If

        If .flags.Navegando Then Call RefreshCharStatus(UserIndex)    'Actualizamos la barca si esta navegando (NicoNZ)

        Call LogEjercitoReal(.Name & " ingresó el " & Date & " cuando era nivel " & .Stats.ELV)
    End With

End Sub

Public Sub RecompensaArmadaReal(ByVal UserIndex As Integer)
    '***************************************************
    'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
    'Last Modification: 15/04/2010
    'Handles the way of gaining new ranks in the "Armada Real"
    '15/04/2010: ZaMa - Agrego recompensas de oro y armaduras
    '***************************************************
    Dim Crimis As Long
    Dim Lvl As Byte
    Dim NextRecom As Long
    Dim Nobleza As Long

    With UserList(UserIndex)
        Lvl = .Stats.ELV
        Crimis = .Faccion.CriminalesMatados
        NextRecom = .Faccion.NextRecompensa
        Nobleza = .Reputacion.NobleRep

        If Crimis < NextRecom Then
            Call WriteChatOverHead(UserIndex, "Mata " & NextRecom - Crimis & " criminales más para recibir la próxima recompensa.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub
        End If

        Select Case NextRecom
            Case 10:    'cuantos matados para el seg rango? y asi con cada rango mira. intenta ahora /recom :D
                If .Faccion.CriminalesMatados < 50 Then
                    Call WriteChatOverHead(UserIndex, "Debes eliminar 50 ciudadanos o mas para subir de Rango.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub
                End If
                .Faccion.RecompensasReal = 1
                .Faccion.NextRecompensa = 20

            Case 20:
                If .Faccion.CriminalesMatados < 70 Then
                    Call WriteChatOverHead(UserIndex, "Debes eliminar 70 ciudadanos o mas para subir de Rango.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub
                End If
                .Faccion.RecompensasReal = 2
                .Faccion.NextRecompensa = 30

            Case 30:
                If .Faccion.CriminalesMatados < 100 Then
                    Call WriteChatOverHead(UserIndex, "Debes eliminar 100 ciudadanos o mas para subir de Rango.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub
                End If
                .Faccion.RecompensasReal = 3
                .Faccion.NextRecompensa = 40

            Case 40:
                If .Faccion.CriminalesMatados < 150 Then
                    Call WriteChatOverHead(UserIndex, "Debes eliminar 150 ciudadanos o mas para subir de Rango.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub
                End If
                .Faccion.RecompensasReal = 4
                .Faccion.NextRecompensa = 50

            Case 50:
                If .Faccion.CriminalesMatados < 200 Then
                    Call WriteChatOverHead(UserIndex, "Debes eliminar 200 ciudadanos o mas para subir de Rango.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub
                End If
                .Faccion.RecompensasReal = 5
                .Faccion.NextRecompensa = 60

            Case 60:
                If Lvl < 27 Then
                    Call WriteChatOverHead(UserIndex, "Mataste suficientes criminales, pero te faltan " & 27 - Lvl & " niveles para poder recibir la próxima recompensa.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub
                End If
                If .Faccion.CriminalesMatados < 250 Then
                    Call WriteChatOverHead(UserIndex, "Debes eliminar 250 ciudadanos o mas para subir de Rango.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub
                End If
                .Faccion.RecompensasReal = 6
                .Faccion.NextRecompensa = 70

            Case 70:
                If .Faccion.CriminalesMatados < 300 Then
                    Call WriteChatOverHead(UserIndex, "Debes eliminar 300 ciudadanos o mas para subir de Rango.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub
                End If
                .Faccion.RecompensasReal = 7
                .Faccion.NextRecompensa = 80

            Case 80:
                If .Faccion.CriminalesMatados < 350 Then
                    Call WriteChatOverHead(UserIndex, "Debes eliminar 350 ciudadanos o mas para subir de Rango.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub
                End If
                .Faccion.RecompensasReal = 8
                .Faccion.NextRecompensa = 90

            Case 90:
                'aca iba xD
                If Lvl < 30 Then
                    Call WriteChatOverHead(UserIndex, "Mataste suficientes criminales, pero te faltan " & 30 - Lvl & " niveles para poder recibir la próxima recompensa.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub
                End If

                If .Faccion.CriminalesMatados < 400 Then
                    Call WriteChatOverHead(UserIndex, "Debes eliminar 400 ciudadanos o mas para subir de Rango.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub
                End If
                .Faccion.RecompensasReal = 9
                .Faccion.NextRecompensa = 100

            Case 100:
                If Nobleza < 2000000 Then
                    Call WriteChatOverHead(UserIndex, "Mataste suficientes criminales, pero te faltan " & 2000000 - Nobleza & " puntos de nobleza para poder recibir la próxima recompensa.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub
                End If
                If .Faccion.CriminalesMatados < 450 Then
                    Call WriteChatOverHead(UserIndex, "Debes eliminar 450 ciudadanos o mas para subir de Rango.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub
                End If
                .Faccion.RecompensasReal = 10
                .Faccion.NextRecompensa = 110

            Case 110:
                If Nobleza < 30000 Then
                    Call WriteChatOverHead(UserIndex, "Mataste suficientes criminales, pero te faltan " & 30000 - Nobleza & " puntos de nobleza para poder recibir la próxima recompensa.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub
                End If
                If .Faccion.CriminalesMatados < 500 Then
                    Call WriteChatOverHead(UserIndex, "Debes eliminar 500 ciudadanos o mas para subir de Rango.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub
                End If
                .Faccion.RecompensasReal = 11
                .Faccion.NextRecompensa = 120

            Case 120:
                If Lvl < 35 Then
                    Call WriteChatOverHead(UserIndex, "Mataste suficientes criminales, pero te faltan " & 35 - Lvl & " niveles para poder recibir la próxima recompensa.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub
                End If
                If Nobleza < 40000 Then
                    Call WriteChatOverHead(UserIndex, "Mataste suficientes criminales, pero te faltan " & 40000 - Nobleza & " puntos de nobleza para poder recibir la próxima recompensa.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub
                End If
                If .Faccion.CriminalesMatados < 700 Then
                    Call WriteChatOverHead(UserIndex, "Debes eliminar 700 ciudadanos o mas para subir de Rango.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub
                End If
                .Faccion.RecompensasReal = 12
                .Faccion.NextRecompensa = 130

            Case 130:
                If Lvl < 36 Then
                    Call WriteChatOverHead(UserIndex, "Mataste suficientes criminales, pero te faltan " & 36 - Lvl & " niveles para poder recibir la próxima recompensa.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub
                End If
                If Nobleza < 50000 Then
                    Call WriteChatOverHead(UserIndex, "Mataste suficientes criminales, pero te faltan " & 50000 - Nobleza & " puntos de nobleza para poder recibir la próxima recompensa.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub
                End If
                If .Faccion.CriminalesMatados < 900 Then
                    Call WriteChatOverHead(UserIndex, "Debes eliminar 900 ciudadanos o mas para subir de Rango.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub
                End If
                .Faccion.RecompensasReal = 13
                .Faccion.NextRecompensa = 140

            Case 140:
                If Lvl < 37 Then
                    Call WriteChatOverHead(UserIndex, "Mataste suficientes criminales, pero te faltan " & 37 - Lvl & " niveles para poder recibir la próxima recompensa.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub
                End If
                If Nobleza < 60000 Then
                    Call WriteChatOverHead(UserIndex, "Mataste suficientes criminales, pero te faltan " & 60000 - Nobleza & " puntos de nobleza para poder recibir la próxima recompensa.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub
                End If
                If .Faccion.CriminalesMatados < 1200 Then
                    Call WriteChatOverHead(UserIndex, "Debes eliminar 1200 ciudadanos o mas para subir de Rango.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub
                End If
                .Faccion.RecompensasReal = 14
                .Faccion.NextRecompensa = 150
                'termine amigo, pero me parece muy limado en asesinos que te pide 3000?, eso hay q pulir igual ja ok
                'ahi quedo andando

            Case 150:
                Call WriteChatOverHead(UserIndex, "Eres uno de mis mejores soldados. Mataste " & Crimis & " criminales, sigue así. Ya no tengo más recompensa para darte que mi agradecimiento. ¡Felicidades!", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                Exit Sub

            Case Else:
                Exit Sub
        End Select
        'che en que parte esta lo que te pide matados?, esa es mi consulta, dejame testear algo.
        Call WriteChatOverHead(UserIndex, "¡¡¡Aquí tienes tu recompensa " & TituloReal(UserIndex) & "!!!", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)

        ' Recompensas de armaduras y exp
        Call GiveFactionArmours(UserIndex, False)
        Call GiveExpReward(UserIndex, .Faccion.RecompensasReal)

    End With

End Sub

Public Sub ExpulsarFaccionReal(ByVal UserIndex As Integer, Optional Expulsado As Boolean = True)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    With UserList(UserIndex)
        .Faccion.ArmadaReal = 0
        'Call PerderItemsFaccionarios(UserIndex)
        If Expulsado Then
            Call WriteConsoleMsg(UserIndex, "¡¡¡Has sido expulsado del ejército real!!!", FontTypeNames.FONTTYPE_FIGHT)
        Else
            Call WriteConsoleMsg(UserIndex, "¡¡¡Te has retirado del ejército real!!!", FontTypeNames.FONTTYPE_FIGHT)
        End If

        If .Invent.ArmourEqpObjIndex <> 0 Then
            'Desequipamos la armadura real si está equipada
            If ObjData(.Invent.ArmourEqpObjIndex).Real = 1 Then Call Desequipar(UserIndex, .Invent.ArmourEqpSlot)
        End If

        If .Invent.EscudoEqpObjIndex <> 0 Then
            'Desequipamos el escudo de caos si está equipado
            If ObjData(.Invent.EscudoEqpObjIndex).Real = 1 Then Call Desequipar(UserIndex, .Invent.EscudoEqpObjIndex)
        End If

        If .flags.Navegando Then Call RefreshCharStatus(UserIndex)    'Actualizamos la barca si esta navegando (NicoNZ)
    End With

End Sub

Public Sub ExpulsarFaccionCaos(ByVal UserIndex As Integer, Optional Expulsado As Boolean = True)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    With UserList(UserIndex)
        .Faccion.FuerzasCaos = 0
        'Call PerderItemsFaccionarios(UserIndex)
        If Expulsado Then
            Call WriteConsoleMsg(UserIndex, "¡¡¡Has sido expulsado de la Legión Oscura!!!", FontTypeNames.FONTTYPE_FIGHT)
        Else
            Call WriteConsoleMsg(UserIndex, "¡¡¡Te has retirado de la Legión Oscura!!!", FontTypeNames.FONTTYPE_FIGHT)
        End If

        If .Invent.ArmourEqpObjIndex <> 0 Then
            'Desequipamos la armadura de caos si está equipada
            If ObjData(.Invent.ArmourEqpObjIndex).Caos = 1 Then Call Desequipar(UserIndex, .Invent.ArmourEqpSlot)
        End If

        If .Invent.EscudoEqpObjIndex <> 0 Then
            'Desequipamos el escudo de caos si está equipado
            If ObjData(.Invent.EscudoEqpObjIndex).Caos = 1 Then Call Desequipar(UserIndex, .Invent.EscudoEqpObjIndex)
        End If

        If .flags.Navegando Then Call RefreshCharStatus(UserIndex)    'Actualizamos la barca si esta navegando (NicoNZ)
    End With

End Sub

Public Function TituloReal(ByVal UserIndex As Integer) As String
    '***************************************************
    'Autor: Unknown
    'Last Modification: 23/01/2007 Pablo (ToxicWaste)
    'Handles the titles of the members of the "Armada Real"
    '***************************************************

    Select Case UserList(UserIndex).Faccion.RecompensasReal
            'Rango 1: Aprendiz (30 Criminales)
            'Rango 2: Escudero (70 Criminales)
            'Rango 3: Soldado (130 Criminales)
            'Rango 4: Sargento (210 Criminales)
            'Rango 5: Caballero (320 Criminales)
            'Rango 6: Comandante (460 Criminales)
            'Rango 7: Capitán (640 Criminales + > lvl 27)
            'Rango 8: Senescal (870 Criminales)
            'Rango 9: Mariscal (1160 Criminales)
            'Rango 10: Condestable (2000 Criminales + > lvl 30)
            'Rangos de Honor de la Armada Real: (Consejo de Bander)
            'Rango 11: Ejecutor Imperial (2500 Criminales + 2.000.000 Nobleza)
            'Rango 12: Protector del Reino (3000 Criminales + 3.000.000 Nobleza)
            'Rango 13: Avatar de la Justicia (3500 Criminales + 4.000.000 Nobleza + > lvl 35)
            'Rango 14: Guardián del Bien (4000 Criminales + 5.000.000 Nobleza + > lvl 36)
            'Rango 15: Campeón de la Luz (5000 Criminales + 6.000.000 Nobleza + > lvl 37)
            'aca estan los nombres xD
        Case 0
            TituloReal = "Aprendiz"
        Case 1
            TituloReal = "Escudero"
        Case 2
            TituloReal = "Soldado"
        Case 3
            TituloReal = "Sargento"
        Case 4
            TituloReal = "Teniente"
        Case 5
            TituloReal = "Comandante"
        Case 6
            TituloReal = "Capitán"
        Case 7
            TituloReal = "Senescal"
        Case 8
            TituloReal = "Mariscal"
        Case 9
            TituloReal = "Condestable"
        Case 10
            TituloReal = "Ejecutor Imperial"
        Case 11
            TituloReal = "Protector del Reino"
        Case 12
            TituloReal = "Avatar de la Justicia"
        Case 13
            TituloReal = "Guardián del Bien"
        Case Else
            TituloReal = "Campeón de la Luz"
    End Select


End Function

Public Sub EnlistarCaos(ByVal UserIndex As Integer)
    '***************************************************
    'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
    'Last Modification: 27/11/2009
    '15/03/2009: ZaMa - No se puede enlistar el fundador de un clan con alineación neutral.
    '27/11/2009: ZaMa - Ahora no se puede enlistar un miembro de un clan neutro, por ende saque la antifaccion.
    'Handles the entrance of users to the "Legión Oscura"
    '***************************************************

    With UserList(UserIndex)
        If Not criminal(UserIndex) Then
            Call WriteChatOverHead(UserIndex, "¡¡¡Lárgate de aquí, bufón!!!", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub
        End If

        If .Faccion.FuerzasCaos = 1 Then
            Call WriteChatOverHead(UserIndex, "¡¡¡Ya perteneces a la legión oscura!!!", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub
        End If

        If .Faccion.ArmadaReal = 1 Then
            Call WriteChatOverHead(UserIndex, "Las sombras reinarán en Argentum. ¡¡¡Fuera de aquí insecto real!!!", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub
        End If

        '[Barrin 17-12-03] Si era miembro de la Armada Real no se puede enlistar
        If .Faccion.RecibioExpInicialReal = 1 Then    'Tomamos el valor de ahí: ¿Recibio la experiencia para entrar?
            Call WriteChatOverHead(UserIndex, "No permitiré que ningún insecto real ingrese a mis tropas.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub
        End If
        '[/Barrin]

        If Not criminal(UserIndex) Then
            Call WriteChatOverHead(UserIndex, "¡¡Ja ja ja!! Tú no eres bienvenido aquí asqueroso ciudadano.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub
        End If

        If .Faccion.CiudadanosMatados < 45 Then
            Call WriteChatOverHead(UserIndex, "Para unirte a nuestras fuerzas debes matar al menos 45 ciudadanos, sólo has matado " & .Faccion.CiudadanosMatados & ".", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub
        End If

        If .Stats.ELV < 25 Then
            Call WriteChatOverHead(UserIndex, "¡¡¡Para unirte a nuestras fuerzas debes ser al menos nivel 25!!!", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub
        End If

        If .GuildIndex > 0 Then
            If modGuilds.GuildAlignment(.GuildIndex) = "Neutral" Then
                Call WriteChatOverHead(UserIndex, "¡¡¡Perteneces a un clan neutro, sal de él si quieres unirte a nuestras fuerzas!!!", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                Exit Sub
            End If
        End If


        If .Faccion.Reenlistadas > 4 Then
            If .Faccion.Reenlistadas = 200 Then
                Call WriteChatOverHead(UserIndex, "Has sido expulsado de las fuerzas oscuras y durante tu rebeldía has atacado a mi ejército. ¡Vete de aquí!", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Else
                Call WriteChatOverHead(UserIndex, "¡Has sido expulsado de las fuerzas oscuras demasiadas veces!", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            End If
            Exit Sub
        End If

        .Faccion.Reenlistadas = .Faccion.Reenlistadas + 1
        .Faccion.FuerzasCaos = 1

        Call WriteChatOverHead(UserIndex, "¡¡¡Bienvenido al lado oscuro!!! Aquí tienes tus armaduras. Derrama sangre ciudadana y real, y serás recompensado, lo prometo.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)

        If .Faccion.RecibioArmaduraCaos = 0 Then

            Call GiveFactionArmours(UserIndex, True)
            Call GiveExpReward(UserIndex, 0)

            .Faccion.RecibioArmaduraCaos = 1
            .Faccion.NivelIngreso = .Stats.ELV
            .Faccion.FechaIngreso = Date

            .Faccion.RecibioExpInicialCaos = 1
            .Faccion.RecompensasCaos = 0
            .Faccion.NextRecompensa = 10    'que raro, esta bien aver... ahora proba
        End If

        If .flags.Navegando Then Call RefreshCharStatus(UserIndex)    'Actualizamos la barca si esta navegando (NicoNZ)

        Call LogEjercitoCaos(.Name & " ingresó el " & Date & " cuando era nivel " & .Stats.ELV)
    End With

End Sub

Public Sub RecompensaCaos(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Pablo (ToxicWaste) & Unknown (orginal version)
    'Last Modification: 15/04/2010
    'Handles the way of gaining new ranks in the "Legión Oscura"
    '15/04/2010: ZaMa - Agrego recompensas de oro y armaduras
    '***************************************************
    Dim Ciudas As Long
    Dim Lvl As Byte
    Dim NextRecom As Long

    With UserList(UserIndex)
        Lvl = .Stats.ELV
        Ciudas = .Faccion.CiudadanosMatados
        NextRecom = .Faccion.NextRecompensa

        If Ciudas < NextRecom Then
            Call WriteChatOverHead(UserIndex, "Mata " & NextRecom - Ciudas & " cuidadanos más para recibir la próxima recompensa.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
            Exit Sub
        End If

        Select Case NextRecom
            Case 10:
                If .Faccion.CiudadanosMatados < 70 Then
                    Call WriteChatOverHead(UserIndex, "Debes eliminar 70 ciudadanos o mas para subir de Rango.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub
                End If
                .Faccion.RecompensasCaos = 1
                .Faccion.NextRecompensa = 20

            Case 20:
                If .Faccion.CiudadanosMatados < 100 Then
                    Call WriteChatOverHead(UserIndex, "Debes eliminar 100 ciudadanos o mas para subir de Rango.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub
                End If
                .Faccion.RecompensasCaos = 2
                .Faccion.NextRecompensa = 30

            Case 30:
                If .Faccion.CiudadanosMatados < 140 Then
                    Call WriteChatOverHead(UserIndex, "Debes eliminar 140 ciudadanos o mas para subir de Rango.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub
                End If
                .Faccion.RecompensasCaos = 3
                .Faccion.NextRecompensa = 40

            Case 40:
                If .Faccion.CiudadanosMatados < 200 Then
                    Call WriteChatOverHead(UserIndex, "Debes eliminar 200 ciudadanos o mas para subir de Rango.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub
                End If
                .Faccion.RecompensasCaos = 4
                .Faccion.NextRecompensa = 50

            Case 50:
                If .Faccion.CiudadanosMatados < 300 Then
                    Call WriteChatOverHead(UserIndex, "Debes eliminar 300 ciudadanos o mas para subir de Rango.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub
                End If
                .Faccion.RecompensasCaos = 5
                .Faccion.NextRecompensa = 60

            Case 60:
                If Lvl < 27 Then
                    Call WriteChatOverHead(UserIndex, "Mataste suficientes ciudadanos, pero te faltan " & 27 - Lvl & " niveles para poder recibir la próxima recompensa.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub
                End If
                If .Faccion.CiudadanosMatados < 400 Then
                    Call WriteChatOverHead(UserIndex, "Debes eliminar 400 ciudadanos o mas para subir de Rango.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub
                End If
                .Faccion.RecompensasCaos = 6
                .Faccion.NextRecompensa = 70

            Case 70:
                If .Faccion.CiudadanosMatados < 500 Then
                    Call WriteChatOverHead(UserIndex, "Debes eliminar 500 ciudadanos o mas para subir de Rango.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub
                End If
                .Faccion.RecompensasCaos = 7
                .Faccion.NextRecompensa = 80

            Case 80:
                If .Faccion.CiudadanosMatados < 600 Then
                    Call WriteChatOverHead(UserIndex, "Debes eliminar 600 ciudadanos o mas para subir de Rango.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub
                End If
                .Faccion.RecompensasCaos = 8
                .Faccion.NextRecompensa = 90

            Case 90:
                If Lvl < 30 Then
                    Call WriteChatOverHead(UserIndex, "Mataste suficientes ciudadanos, pero te faltan " & 30 - Lvl & " niveles para poder recibir la próxima recompensa.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub
                End If
                If .Faccion.CiudadanosMatados < 700 Then
                    Call WriteChatOverHead(UserIndex, "Debes eliminar 700 ciudadanos o mas para subir de Rango.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub
                End If
                .Faccion.RecompensasCaos = 9
                .Faccion.NextRecompensa = 100

            Case 100:
                If Lvl < 31 Then
                    Call WriteChatOverHead(UserIndex, "Mataste suficientes ciudadanos, pero te faltan " & 31 - Lvl & " niveles para poder recibir la próxima recompensa.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub
                End If
                If .Faccion.CiudadanosMatados < 800 Then
                    Call WriteChatOverHead(UserIndex, "Debes eliminar 800 ciudadanos o mas para subir de Rango.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub
                End If
                .Faccion.RecompensasCaos = 10
                .Faccion.NextRecompensa = 110

            Case 110:
                If Lvl < 33 Then
                    Call WriteChatOverHead(UserIndex, "Mataste suficientes ciudadanos, pero te faltan " & 33 - Lvl & " niveles para poder recibir la próxima recompensa.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub
                End If
                If .Faccion.CiudadanosMatados < 900 Then
                    Call WriteChatOverHead(UserIndex, "Debes eliminar 900 ciudadanos o mas para subir de Rango.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub
                End If
                .Faccion.RecompensasCaos = 11
                .Faccion.NextRecompensa = 120

            Case 120:
                If Lvl < 35 Then
                    Call WriteChatOverHead(UserIndex, "Mataste suficientes ciudadanos, pero te faltan " & 35 - Lvl & " niveles para poder recibir la próxima recompensa.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub
                End If
                If .Faccion.CiudadanosMatados < 1000 Then
                    Call WriteChatOverHead(UserIndex, "Debes eliminar 1000 ciudadanos o mas para subir de Rango.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub
                End If
                .Faccion.RecompensasCaos = 12
                .Faccion.NextRecompensa = 130

            Case 130:
                If Lvl < 36 Then
                    Call WriteChatOverHead(UserIndex, "Mataste suficientes ciudadanos, pero te faltan " & 36 - Lvl & " niveles para poder recibir la próxima recompensa.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub
                End If
                If .Faccion.CiudadanosMatados < 1200 Then
                    Call WriteChatOverHead(UserIndex, "Debes eliminar 1200 ciudadanos o mas para subir de Rango.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub
                End If
                .Faccion.RecompensasCaos = 13
                .Faccion.NextRecompensa = 140

            Case 140:
                If Lvl < 37 Then
                    Call WriteChatOverHead(UserIndex, "Mataste suficientes ciudadanos, pero te faltan " & 37 - Lvl & " niveles para poder recibir la próxima recompensa.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub
                End If
                If .Faccion.CiudadanosMatados < 1500 Then
                    Call WriteChatOverHead(UserIndex, "Debes eliminar 1500 ciudadanos o mas para subir de Rango.", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                    Exit Sub
                End If
                .Faccion.RecompensasCaos = 14
                .Faccion.NextRecompensa = 150

            Case 150:
                Call WriteChatOverHead(UserIndex, "Eres uno de mis mejores soldados. Mataste " & Ciudas & " ciudadanos . Tu única recompensa será la sangre derramada. ¡¡Continúa así!!", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                Exit Sub

            Case Else:
                Exit Sub

        End Select

        Call WriteChatOverHead(UserIndex, "¡¡¡Bien hecho " & TituloCaos(UserIndex) & ", aquí tienes tu recompensa!!!", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)

        ' Recompensas de armaduras y exp
        Call GiveFactionArmours(UserIndex, True)
        Call GiveExpReward(UserIndex, .Faccion.RecompensasCaos)

    End With

End Sub

Public Function TituloCaos(ByVal UserIndex As Integer) As String
    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 23/01/2007 Pablo (ToxicWaste)
    'Handles the titles of the members of the "Legión Oscura"
    '***************************************************
    'Rango 1: Acólito (70)
    'Rango 2: Alma Corrupta (160)
    'Rango 3: Paria (300)
    'Rango 4: Condenado (490)
    'Rango 5: Esbirro (740)
    'Rango 6: Sanguinario (1100)
    'Rango 7: Corruptor (1500 + lvl 27)
    'Rango 8: Heraldo Impio (2010)
    'Rango 9: Caballero de la Oscuridad (2700)
    'Rango 10: Señor del Miedo (4600 + lvl 30)
    'Rango 11: Ejecutor Infernal (5800 + lvl 31)
    'Rango 12: Protector del Averno (6990 + lvl 33)
    'Rango 13: Avatar de la Destrucción (8100 + lvl 35)
    'Rango 14: Guardián del Mal (9300 + lvl 36)
    'Rango 15: Campeón de la Oscuridad (11500 + lvl 37)

    Select Case UserList(UserIndex).Faccion.RecompensasCaos
        Case 0
            TituloCaos = "Acólito"
        Case 1
            TituloCaos = "Alma Corrupta"
        Case 2
            TituloCaos = "Paria"
        Case 3
            TituloCaos = "Condenado"
        Case 4
            TituloCaos = "Esbirro"
        Case 5
            TituloCaos = "Sanguinario"
        Case 6
            TituloCaos = "Corruptor"
        Case 7
            TituloCaos = "Heraldo Impío"
        Case 8
            TituloCaos = "Caballero de la Oscuridad"
        Case 9
            TituloCaos = "Señor del Miedo"
        Case 10
            TituloCaos = "Ejecutor Infernal"
        Case 11
            TituloCaos = "Protector del Averno"
        Case 12
            TituloCaos = "Avatar de la Destrucción"
        Case 13
            TituloCaos = "Guardián del Mal"
        Case Else
            TituloCaos = "Campeón de la Oscuridad"
    End Select

End Function


