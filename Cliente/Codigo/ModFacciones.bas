Attribute VB_Name = "ModFacciones"
'Imperium AO Clon version 1.0
'modFacciones reescrito de 0 por mannakia basandome
'en los procesimientos de alkon 12.2
'Copyright (C) 2002 Márquez Pablo Ignacio


Option Explicit
Public Sub EnlistarCaos(ByVal UserIndex As Integer)
Dim Matados As Integer
Matados = (UserList(UserIndex).Faccion.RenegadosMatados + UserList(UserIndex).Faccion.ArmadaMatados + UserList(UserIndex).Faccion.CiudadanosMatados + UserList(UserIndex).Faccion.MilicianosMatados + UserList(UserIndex).Faccion.RepublicanosMatados)

If UserList(UserIndex).Faccion.FuerzasCaos Then
    Call WriteChatOverHead(UserIndex, "Ya perteneces a la horda del caos Traeme Mas almas.", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

If Matados < 30 Then
    Call WriteChatOverHead(UserIndex, "Para unirte a nuestras fuerzas debes matar al menos 30 enemigos, solo has matado " & Matados, str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

If UserList(UserIndex).Stats.ELV < 40 Then
    Call WriteChatOverHead(UserIndex, "Para unirte a nuestras fuerzas debes ser al menos Nivel 40.", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

UserList(UserIndex).Faccion.FuerzasCaos = 1
UserList(UserIndex).Faccion.Renegado = 0
UserList(UserIndex).Faccion.Rango = 1
Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharStatus(UserList(UserIndex).Char.CharIndex, UserTypeColor(UserIndex)))
 
'------- Ropa -------
Dim MiObj As Obj
Dim bajos As Byte
MiObj.amount = 1
    
If UserList(UserIndex).Raza = Enano Or UserList(UserIndex).Raza = Gnomo Then
    bajos = 1
End If
    
Select Case UserList(UserIndex).Clase
    Case eClass.Clerigo
        MiObj.ObjIndex = 1500 + bajos
    Case eClass.Mago
        MiObj.ObjIndex = 1502 + bajos
    Case eClass.Guerrero
        MiObj.ObjIndex = 1504 + bajos
    Case eClass.Asesino
        MiObj.ObjIndex = 1506 + bajos
    Case eClass.Bardo
        MiObj.ObjIndex = 1508 + bajos
    Case eClass.Druida
        MiObj.ObjIndex = 1510 + bajos
    Case eClass.Gladiador
        MiObj.ObjIndex = 1512 + bajos
    Case eClass.Paladin
        MiObj.ObjIndex = 1514 + bajos
    Case eClass.Cazador
        MiObj.ObjIndex = 1516 + bajos
    Case eClass.Mercenario
        MiObj.ObjIndex = 1518 + bajos
    Case eClass.nigromante
        MiObj.ObjIndex = 1520 + bajos
    Case eClass.nigromante
        MiObj.ObjIndex = 1520 + bajos
End Select

If Not MeterItemEnInventario(UserIndex, MiObj) Then
    Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
End If
'------- Ropa -------

Call WriteChatOverHead(UserIndex, "¡¡¡Bienvenido a la Horda del Caos!!!, aqui tienes tus vestimentas. Cumple bien tu labor exterminando Criminales y me encargaré de recompensarte.", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)

End Sub
Public Sub EnlistarMilicia(ByVal UserIndex As Integer)

Dim Matados As Integer
Matados = (UserList(UserIndex).Faccion.RenegadosMatados + UserList(UserIndex).Faccion.ArmadaMatados + UserList(UserIndex).Faccion.CiudadanosMatados)

If UserList(UserIndex).Faccion.Milicia = 1 Then
    Call WriteChatOverHead(UserIndex, "Ya perteneces a las tropas milicianas Ve a combatir enemigos", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

If UserList(UserIndex).Faccion.Republicano = 0 Then
    Call WriteChatOverHead(UserIndex, "No aceptamos otros seguidores de otras Facciones en la Milicia Republicana.", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

If UserList(UserIndex).Faccion.FuerzasCaos = 1 Or UserList(UserIndex).Faccion.ArmadaReal = 1 Then
    Call WriteChatOverHead(UserIndex, "Sal de aqui, Asqueroso enemigo.", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

If Matados < 10 Then
    Call WriteChatOverHead(UserIndex, "Para unirte a nuestras fuerzas debes matar al menos 10 enemigos, solo has matado " & Matados, str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

If UserList(UserIndex).Stats.ELV < 25 Then
    Call WriteChatOverHead(UserIndex, "Para unirte a nuestras fuerzas debes ser al menos de nivel 25", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If
 
With UserList(UserIndex)
    If .GuildIndex > 0 Then
        If modGuilds.GuildFounder(.GuildIndex) = .name Then
            If modGuilds.GuildAlignment(.GuildIndex) = "Neutro" Then
                Call WriteChatOverHead(UserIndex, "Eres el fundador de un clan neutro", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                Exit Sub
            End If
        End If
    End If
End With

UserList(UserIndex).Faccion.Milicia = 1
UserList(UserIndex).Faccion.Republicano = 0
UserList(UserIndex).Faccion.Rango = 1

Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharStatus(UserList(UserIndex).Char.CharIndex, UserTypeColor(UserIndex)))
'------- Ropa -------
Dim MiObj As Obj
Dim bajos As Byte
MiObj.amount = 1
    
If UserList(UserIndex).Raza = Enano Or UserList(UserIndex).Raza = Gnomo Then
    MiObj.ObjIndex = 1587
Else
    MiObj.ObjIndex = 1588
End If

If Not MeterItemEnInventario(UserIndex, MiObj) Then
    Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
End If
'------- Ropa -------

Call WriteChatOverHead(UserIndex, "Bienvenido a la Milicia Republicana, aqui tienes tu Armadura. Cumple bien tu labor exterminando Criminales y me encargaré de recompensarte.", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
 
If UserList(UserIndex).GuildIndex > 0 Then
    If modGuilds.GuildAlignment(UserList(UserIndex).GuildIndex) = "Neutro" Then
        Call modGuilds.m_EcharMiembroDeClan(-1, UserList(UserIndex).name)
        Call WriteConsoleMsg(1, UserIndex, "Has sido expulsado del clan por tu nueva facción.", FontTypeNames.FONTTYPE_GUILD)
    End If
End If

If UserList(UserIndex).flags.Navegando Then Call RefreshCharStatus(UserIndex) 'Actualizamos la barca si esta navegando (NicoNZ)

End Sub
Public Sub EnlistarArmadaReal(ByVal UserIndex As Integer)
If UserList(UserIndex).Faccion.ArmadaReal = 1 Then
    Call WriteChatOverHead(UserIndex, "¡Ya perteneces a las tropas reales! Ve a combatir criminales", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

If UserList(UserIndex).Faccion.Ciudadano <> 1 Then
    Call WriteChatOverHead(UserIndex, "No aceptamos otros seguidores de otras Facciones en la armada imperial.", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

If UserList(UserIndex).Faccion.FuerzasCaos = 1 Or UserList(UserIndex).Faccion.Milicia = 1 Then
    Call WriteChatOverHead(UserIndex, "Sal de aqui, Asqueroso enemigo.", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

If (UserList(UserIndex).Faccion.RenegadosMatados + UserList(UserIndex).Faccion.RepublicanosMatados + UserList(UserIndex).Faccion.MilicianosMatados + UserList(UserIndex).Faccion.CaosMatados) < 15 Then
    Call WriteChatOverHead(UserIndex, "Para unirte a nuestras fuerzas debes matar al menos 15 enemigos, solo has matado " & (UserList(UserIndex).Faccion.RenegadosMatados + UserList(UserIndex).Faccion.RepublicanosMatados + UserList(UserIndex).Faccion.MilicianosMatados + UserList(UserIndex).Faccion.CaosMatados), str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

If UserList(UserIndex).Stats.ELV < 25 Then
    Call WriteChatOverHead(UserIndex, "Para unirte a nuestras fuerzas debes ser al menos de nivel 25", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If
 
With UserList(UserIndex)
    If .GuildIndex > 0 Then
        If modGuilds.GuildFounder(.GuildIndex) = .name Then
            If modGuilds.GuildAlignment(.GuildIndex) = "Neutro" Then
                Call WriteChatOverHead(UserIndex, "Eres el fundador de un clan neutro", str(Npclist(.flags.TargetNPC).Char.CharIndex), vbWhite)
                Exit Sub
            End If
        End If
    End If
End With

UserList(UserIndex).Faccion.ArmadaReal = 1
UserList(UserIndex).Faccion.Ciudadano = 0
UserList(UserIndex).Faccion.Rango = 1
Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharStatus(UserList(UserIndex).Char.CharIndex, UserTypeColor(UserIndex)))
Dim MiObj As Obj
Dim bajos As Byte
MiObj.amount = 1
    
If UserList(UserIndex).Raza = Enano Or UserList(UserIndex).Raza = Gnomo Then
    bajos = 1
End If
    
Select Case UserList(UserIndex).Clase
    Case eClass.Clerigo
        MiObj.ObjIndex = 1544 + bajos
    Case eClass.Mago
        MiObj.ObjIndex = 1546 + bajos
    Case eClass.Guerrero
        MiObj.ObjIndex = 1548 + bajos
    Case eClass.Asesino
        MiObj.ObjIndex = 1550 + bajos
    Case eClass.Bardo
        MiObj.ObjIndex = 1552 + bajos
    Case eClass.Druida
        MiObj.ObjIndex = 1554 + bajos
    Case eClass.Gladiador
        MiObj.ObjIndex = 1556 + bajos
    Case eClass.Paladin
        MiObj.ObjIndex = 1558 + bajos
    Case eClass.Cazador
        MiObj.ObjIndex = 1560 + bajos
    Case eClass.Mercenario
        MiObj.ObjIndex = 1562 + bajos
    Case eClass.nigromante
        MiObj.ObjIndex = 1564 + bajos
End Select

If Not MeterItemEnInventario(UserIndex, MiObj) Then
    Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
End If

Call WriteChatOverHead(UserIndex, "Bienvenido al Ejército Imperial, aqui tienes tus vestimentas. Cumple bien tu labor exterminando Criminales y me encargaré de recompensarte.", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
 
If UserList(UserIndex).GuildIndex > 0 Then
    If modGuilds.GuildAlignment(UserList(UserIndex).GuildIndex) = "Neutro" Then
        Call modGuilds.m_EcharMiembroDeClan(-1, UserList(UserIndex).name)
        Call WriteConsoleMsg(1, UserIndex, "Has sido expulsado del clan por tu nueva facción.", FontTypeNames.FONTTYPE_GUILD)
    End If
End If

If UserList(UserIndex).flags.Navegando Then Call RefreshCharStatus(UserIndex) 'Actualizamos la barca si esta navegando (NicoNZ)

End Sub

Public Sub RecompensaArmadaReal(ByVal UserIndex As Integer)
Dim Matados As Long

If UserList(UserIndex).Faccion.Rango = 10 Then
    Exit Sub
End If

Matados = UserList(UserIndex).Faccion.RenegadosMatados + UserList(UserIndex).Faccion.CaosMatados + UserList(UserIndex).Faccion.MilicianosMatados + UserList(UserIndex).Faccion.RepublicanosMatados

If Matados < matadosArmada(UserList(UserIndex).Faccion.Rango) Then
    Call WriteChatOverHead(UserIndex, "Mata " & matadosArmada(UserList(UserIndex).Faccion.Rango) - Matados & " Criminales más para recibir la próxima Recompensa", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

UserList(UserIndex).Faccion.Rango = UserList(UserIndex).Faccion.Rango + 1
If UserList(UserIndex).Faccion.Rango >= 6 Then ' Segunda jeraquia xD
    Dim MiObj As Obj
    MiObj.amount = 1
    Dim bajos As Byte
    
    If UserList(UserIndex).Raza = Enano Or UserList(UserIndex).Raza = Gnomo Then
        bajos = 1
    End If
        
    Select Case UserList(UserIndex).Clase
        Case eClass.Clerigo
            MiObj.ObjIndex = 1566 + bajos
        Case eClass.Mago
            MiObj.ObjIndex = 1568 + bajos
        Case eClass.Guerrero
            MiObj.ObjIndex = 1570 + bajos
        Case eClass.Asesino
            MiObj.ObjIndex = 1572 + bajos
        Case eClass.Bardo
            MiObj.ObjIndex = 1574 + bajos
        Case eClass.Druida
            MiObj.ObjIndex = 1576 + bajos
        Case eClass.Gladiador
            MiObj.ObjIndex = 1578 + bajos
        Case eClass.Paladin
            MiObj.ObjIndex = 1580 + bajos
        Case eClass.Cazador
            MiObj.ObjIndex = 1582 + bajos
        Case eClass.Mercenario
            MiObj.ObjIndex = 1584 + bajos
        Case eClass.nigromante
            MiObj.ObjIndex = 1586 + bajos
        Case eClass.nigromante
            MiObj.ObjIndex = 1586 + bajos
    End Select

    If Not MeterItemEnInventario(UserIndex, MiObj) Then
        Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
End If

End Sub
Public Sub RecompensaMilicia(ByVal UserIndex As Integer)
Dim Matados As Long

If UserList(UserIndex).Faccion.Rango = 7 Then
    Exit Sub
End If

Matados = (UserList(UserIndex).Faccion.RenegadosMatados + UserList(UserIndex).Faccion.ArmadaMatados + UserList(UserIndex).Faccion.CiudadanosMatados)
If Matados < matadosArmada(UserList(UserIndex).Faccion.Rango) Then
    Call WriteChatOverHead(UserIndex, "Mata " & matadosArmada(UserList(UserIndex).Faccion.Rango) - Matados & " Criminales más para recibir la próxima Recompensa", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

UserList(UserIndex).Faccion.Rango = UserList(UserIndex).Faccion.Rango + 1
If UserList(UserIndex).Faccion.Rango >= 4 Then
    Dim MiObj As Obj
    MiObj.amount = 1
    Dim bajos As Byte
    
    If UserList(UserIndex).Raza = Enano Or UserList(UserIndex).Raza = Gnomo Then
        bajos = 1
    End If
        
    Select Case UserList(UserIndex).Clase
        Case eClass.Clerigo, eClass.Mago, eClass.Bardo, eClass.Druida, eClass.nigromante
            MiObj.ObjIndex = 1592 + bajos
        Case eClass.Guerrero, eClass.Gladiador, eClass.Cazador, eClass.Mercenario, eClass.Paladin, eClass.Asesino
            MiObj.ObjIndex = 1590 + bajos
    End Select

    If Not MeterItemEnInventario(UserIndex, MiObj) Then
        Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
End If
End Sub
Public Sub RecompensaCaos(ByVal UserIndex As Integer)
Dim Matados As Long

If UserList(UserIndex).Faccion.Rango = 10 Then
    Exit Sub
End If

Matados = UserList(UserIndex).Faccion.RenegadosMatados + UserList(UserIndex).Faccion.CaosMatados + UserList(UserIndex).Faccion.MilicianosMatados + UserList(UserIndex).Faccion.RepublicanosMatados

If Matados < matadosCaos(UserList(UserIndex).Faccion.Rango) Then
    Call WriteChatOverHead(UserIndex, "Mata " & matadosCaos(UserList(UserIndex).Faccion.Rango) - Matados & " enemigos más para recibir la próxima Recompensa", str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex), vbWhite)
    Exit Sub
End If

UserList(UserIndex).Faccion.Rango = UserList(UserIndex).Faccion.Rango + 1
If UserList(UserIndex).Faccion.Rango >= 6 Then ' Segunda jeraquia xD
    Dim MiObj As Obj
    MiObj.amount = 1
    Dim bajos As Byte
    
    If UserList(UserIndex).Raza = Enano Or UserList(UserIndex).Raza = Gnomo Then
        bajos = 1
    End If
        
    Select Case UserList(UserIndex).Clase
        Case eClass.Clerigo
            MiObj.ObjIndex = 1522 + bajos
        Case eClass.Mago
            MiObj.ObjIndex = 1524 + bajos
        Case eClass.Guerrero
            MiObj.ObjIndex = 1526 + bajos
        Case eClass.Asesino
            MiObj.ObjIndex = 1528 + bajos
        Case eClass.Bardo
            MiObj.ObjIndex = 1530 + bajos
        Case eClass.Druida
            MiObj.ObjIndex = 1532 + bajos
        Case eClass.Gladiador
            MiObj.ObjIndex = 1534 + bajos
        Case eClass.Paladin
            MiObj.ObjIndex = 1536 + bajos
        Case eClass.Cazador
            MiObj.ObjIndex = 1538 + bajos
        Case eClass.Mercenario
            MiObj.ObjIndex = 1540 + bajos
        Case eClass.nigromante
            MiObj.ObjIndex = 1542 + bajos
    End Select

    If Not MeterItemEnInventario(UserIndex, MiObj) Then
        Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
End If
End Sub
Public Sub ExpulsarFaccionReal(ByVal UserIndex As Integer, Optional Expulsado As Boolean = True)

    UserList(UserIndex).Faccion.ArmadaReal = 0
    UserList(UserIndex).Faccion.Ciudadano = 1
    UserList(UserIndex).Faccion.Rango = 0

    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharStatus(UserList(UserIndex).Char.CharIndex, UserTypeColor(UserIndex)))

    If Expulsado Then
        Call WriteConsoleMsg(1, UserIndex, "¡Has sido expulsado de las tropas reales!.", FontTypeNames.FONTTYPE_FIGHT)
    Else
        Call WriteConsoleMsg(1, UserIndex, "Te has retirado de las tropas reales.", FontTypeNames.FONTTYPE_FIGHT)
    End If
    
    If UserList(UserIndex).Invent.ArmourEqpObjIndex Then
        'Desequipamos la armadura real si está equipada
        If ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Real = 1 Then Call Desequipar(UserIndex, UserList(UserIndex).Invent.ArmourEqpSlot)
    End If
    
    If UserList(UserIndex).Invent.EscudoEqpObjIndex Then
        If ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).Real = 1 Then Call Desequipar(UserIndex, UserList(UserIndex).Invent.EscudoEqpObjIndex)
    End If
    
    Call QuitarItemsFaccionarios(UserIndex)
    
    If UserList(UserIndex).flags.Navegando Then Call RefreshCharStatus(UserIndex) 'Actualizamos la barca si esta navegando (NicoNZ)

End Sub
Public Sub ExpulsarFaccionCaos(ByVal UserIndex As Integer, Optional ByVal Expulsar As Boolean)
    UserList(UserIndex).Faccion.FuerzasCaos = 0
    UserList(UserIndex).Faccion.Renegado = 1
    UserList(UserIndex).Faccion.Rango = 0
    
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharStatus(UserList(UserIndex).Char.CharIndex, UserTypeColor(UserIndex)))

    If Expulsar Then
        Call WriteConsoleMsg(1, UserIndex, "¡Has sido expulsado de las fuerza del caos!.", FontTypeNames.FONTTYPE_FIGHT)
    Else
        Call WriteConsoleMsg(1, UserIndex, "¡Te has retirado de las fuerza del caos!.", FontTypeNames.FONTTYPE_FIGHT)
    End If
    
    If UserList(UserIndex).Invent.ArmourEqpObjIndex Then
        'Desequipamos la armadura real si está equipada
        If ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Caos = 1 Then Call Desequipar(UserIndex, UserList(UserIndex).Invent.ArmourEqpSlot)
    End If
    
    If UserList(UserIndex).Invent.EscudoEqpObjIndex Then
        If ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).Caos = 1 Then Call Desequipar(UserIndex, UserList(UserIndex).Invent.EscudoEqpObjIndex)
    End If
    
    Call QuitarItemsFaccionarios(UserIndex)
    
    If UserList(UserIndex).flags.Navegando Then Call RefreshCharStatus(UserIndex) 'Actualizamos la barca si esta navegando (NicoNZ)

End Sub
Public Sub ExpulsarFaccionMilicia(ByVal UserIndex As Integer, Optional ByVal Expulsar As Boolean)
    UserList(UserIndex).Faccion.Milicia = 0
    UserList(UserIndex).Faccion.Republicano = 1
    UserList(UserIndex).Faccion.Rango = 0

    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharStatus(UserList(UserIndex).Char.CharIndex, UserTypeColor(UserIndex)))

    If Expulsar Then
        Call WriteConsoleMsg(1, UserIndex, "¡Has sido expulsado de las tropas republicanas.", FontTypeNames.FONTTYPE_FIGHT)
    Else
        Call WriteConsoleMsg(1, UserIndex, "¡Te has retirado de las tropas republicanas.!.", FontTypeNames.FONTTYPE_FIGHT)
    End If
    
    If UserList(UserIndex).Invent.ArmourEqpObjIndex Then
        'Desequipamos la armadura real si está equipada
        If ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Milicia = 1 Then Call Desequipar(UserIndex, UserList(UserIndex).Invent.ArmourEqpSlot)
    End If
    
    If UserList(UserIndex).Invent.EscudoEqpObjIndex Then
        If ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).Milicia = 1 Then Call Desequipar(UserIndex, UserList(UserIndex).Invent.EscudoEqpObjIndex)
    End If
    
    Call QuitarItemsFaccionarios(UserIndex)
    
    If UserList(UserIndex).flags.Navegando Then Call RefreshCharStatus(UserIndex) 'Actualizamos la barca si esta navegando (NicoNZ)

End Sub
Public Function TituloCaos(ByVal UserIndex As Integer) As String
    Select Case UserList(UserIndex).Faccion.Rango
        Case 1
            TituloCaos = "Lancero del Caos"
        Case 2
            TituloCaos = "Guerrero del Caos"
        Case 3
            TituloCaos = "Teniente del Caos"
        Case 4
            TituloCaos = "Comandante del Caos"
        Case 5
            TituloCaos = "General del Caos"
        Case 6
            TituloCaos = "Elite del Caos"
        Case 7
            TituloCaos = "Asolador de las Sombras"
        Case 8
            TituloCaos = "Caballero Negro"
        Case 9
            TituloCaos = "Segador Infernal"
        Case 10
            TituloCaos = "Emperador de las Tinieblas"
    End Select
End Function
Public Function TituloReal(ByVal UserIndex As Integer) As String
    Select Case UserList(UserIndex).Faccion.Rango
        Case 1
            TituloReal = "Soldado Real"
        Case 2
            TituloReal = "Sargento Real"
        Case 3
            TituloReal = "Teniente Real"
        Case 4
            TituloReal = "Comandante Real"
        Case 5
            TituloReal = "General Real"
        Case 6
            TituloReal = "Elite Real"
        Case 7
            TituloReal = "Guardian del Bien"
        Case 8
            TituloReal = "Caballero Imperial"
        Case 9
            TituloReal = "Justiciero"
        Case 10
            TituloReal = "Rey Imperial"
    End Select
End Function
Public Function TituloMilicia(ByVal UserIndex As Integer) As String
    Select Case UserList(UserIndex).Faccion.Rango
        Case 1
            TituloMilicia = "Milicia de Reserva"
        Case 2
            TituloMilicia = "Miliciano"
        Case 3
            TituloMilicia = "Miliciano Elite"
        Case 4
            TituloMilicia = "Soldado de la República"
        Case 5
            TituloMilicia = "Soldado Raso"
        Case 6
            TituloMilicia = "Soldado Elite"
        Case 7
            TituloMilicia = "Comandante Republicano"
    End Select
End Function
Public Function matadosArmada(ByVal Rango As Byte) As Integer
    Select Case Rango
        Case 1
            matadosArmada = 20
        Case 2
            matadosArmada = 25
        Case 3
            matadosArmada = 30
        Case 4
            matadosArmada = 35
        Case 5
            matadosArmada = 40
        Case 6
            matadosArmada = 50
        Case 7
            matadosArmada = 85
        Case 8
            matadosArmada = 95
        Case 9
            matadosArmada = 105
    End Select
End Function
Public Function matadosCaos(ByVal Rango As Byte) As Integer
    Select Case Rango
        Case 1
            matadosCaos = 20
        Case 2
            matadosCaos = 30
        Case 3
            matadosCaos = 40
        Case 4
            matadosCaos = 50
        Case 5
            matadosCaos = 60
        Case 6
            matadosCaos = 70
        Case 7
            matadosCaos = 80
        Case 8
            matadosCaos = 90
        Case 9
            matadosCaos = 100
    End Select
End Function
Public Function matadosMilicia(ByVal Rango As Byte) As Integer
    Select Case Rango
        Case 1
            matadosMilicia = 15
        Case 2
            matadosMilicia = 20
        Case 3
            matadosMilicia = 25
        Case 4
            matadosMilicia = 30
        Case 5
            matadosMilicia = 60
        Case 6
            matadosMilicia = 70
    End Select
End Function
Public Sub QuitarItemsFaccionarios(ByVal UserIndex As Integer)
    Dim i As Byte
    Dim ObjIndex As Integer
    For i = 1 To MAX_INVENTORY_SLOTS
        ObjIndex = UserList(UserIndex).Invent.Object(i).ObjIndex
        If ObjIndex <> 0 Then
            If ObjData(ObjIndex).Caos = 1 Or ObjData(ObjIndex).Real = 1 Or ObjData(ObjIndex).Milicia = 1 Then
                QuitarUserInvItem UserIndex, i, UserList(UserIndex).Invent.Object(i).amount
                UpdateUserInv False, UserIndex, i
            End If
        End If
    Next i
            
End Sub
