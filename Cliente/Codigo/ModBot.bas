Attribute VB_Name = "ModBot"
Option Explicit
Public BotUser As Boolean

Public Sub CargarHechizosBot()
On Error GoTo Err
    Dim Leer As New clsIniReader
    Dim NumConjuros As Integer
    Dim i As Integer
    
    Leer.Initialize DatPath & "Bots.dat"
    
    NumConjuros = val(Leer.GetValue("INIT", "Nums"))
    ReDim BotHechizos(1 To NumConjuros) As tHechizosBot
    
    For i = 1 To NumConjuros
        BotHechizos(i).PalabrasMagicas = Leer.GetValue("BH" & i, "Palabras")
        BotHechizos(i).TargetMsg = Leer.GetValue("BH" & i, "TargetMsg")
        
        BotHechizos(i).ManaRequerido = Leer.GetValue("BH" & i, "ManaRequerido")
        
        BotHechizos(i).FXgrh = val(Leer.GetValue("BH" & i, "FX"))
        BotHechizos(i).WAV = val(Leer.GetValue("BH" & i, "Wav"))
        
        BotHechizos(i).Particle = val(Leer.GetValue("BH" & i, "Particle"))
        BotHechizos(i).life = val(Leer.GetValue("BH" & i, "PartLife"))
        
        BotHechizos(i).SubeHP = val(Leer.GetValue("BH" & i, "SubeHP"))
        BotHechizos(i).MinHP = val(Leer.GetValue("BH" & i, "MinHP"))
        BotHechizos(i).MaxHP = val(Leer.GetValue("BH" & i, "MaxHP"))
        
        BotHechizos(i).SubeAgilidad = val(Leer.GetValue("BH" & i, "SubeAgilidad"))
        BotHechizos(i).MinAgilidad = val(Leer.GetValue("BH" & i, "MinAgilidad"))
        BotHechizos(i).MaxAgilidad = val(Leer.GetValue("BH" & i, "MaxAgilidad"))
        
        BotHechizos(i).SubeFuerza = val(Leer.GetValue("BH" & i, "SubeFuerza"))
        BotHechizos(i).MinFuerza = val(Leer.GetValue("BH" & i, "MinFuerza"))
        BotHechizos(i).MaxFuerza = val(Leer.GetValue("BH" & i, "MaxFuerza"))
        
        BotHechizos(i).Paraliza = val(Leer.GetValue("BH" & i, "Paraliza"))
        BotHechizos(i).Inmoviliza = val(Leer.GetValue("BH" & i, "Inmoviliza"))
        BotHechizos(i).Envenena = val(Leer.GetValue("BH" & i, "Envenena"))
        BotHechizos(i).Maldicion = val(Leer.GetValue("BH" & i, "Maldicion"))
        BotHechizos(i).Bendicion = val(Leer.GetValue("BH" & i, "Bendicion"))
        BotHechizos(i).Estupidez = val(Leer.GetValue("BH" & i, "Estupidez"))
        BotHechizos(i).Ceguera = val(Leer.GetValue("BH" & i, "Ceguera"))
        BotHechizos(i).Revivir = val(Leer.GetValue("BH" & i, "Revivir"))

        BotHechizos(i).RemoverParalisis = val(Leer.GetValue("BH" & i, "RemoverParalisis"))
        BotHechizos(i).CuraVeneno = val(Leer.GetValue("BH" & i, "CuraVeneno"))
        BotHechizos(i).RemoverMaldicion = val(Leer.GetValue("BH" & i, "RemoverMaldicion"))
        BotHechizos(i).RemueveInvisibilidadParcial = val(Leer.GetValue("BH" & i, "RemueveInvisibilidadParcial"))

    Next i

Exit Sub
Err:
    Call LogError("CargarHechizosBot Desc:" & Err.description & " Number:" & Err.Number)
    
End Sub
Sub BotAI(ByVal BotIndex As Integer)
'On Error GoTo ErrorHandler

Dim UsaMano As Boolean
Dim UsaHechizo As Boolean
Dim NpcIndex As Integer
Dim tHeading As eHeading

    With Bots(BotIndex)
        
        NpcIndex = .NpcIndex
        
        'Checkeamos que no sea un NULL
        If .MinHP <= 0 Then
            Call DeleteBot(BotIndex)
            Call MuereNpc(NpcIndex, 0)
            Exit Sub
        End If
        
        If Npclist(NpcIndex).Pos.X = 0 And Npclist(NpcIndex).Pos.Y = 0 Then
            Call DeleteBot(BotIndex)
            Call MuereNpc(NpcIndex, 0)
            Exit Sub
        End If
        
        UsaMano = BotPotea(BotIndex)
        
        If Npclist(NpcIndex).flags.Paralizado And LanzaClase(.Clase) > 0 Then
            If IntervaloPuedeHechizo(BotIndex) Then
                UsaHechizo = BotRemueve(BotIndex)
                Dim Direccion As Byte
                .RandomDire = CByte(RandomNumber(eHeading.NORTH, eHeading.WEST))
                .RandomDire = Direccion
                Exit Sub
            End If
        End If
        
        If Not Npclist(NpcIndex).flags.Inmovilizado = 1 And Not Npclist(NpcIndex).flags.Paralizado = 1 Then
            If .TargetNPC Then
                If Npclist(.TargetNPC).Pos.X = 0 Or Npclist(.TargetNPC).Pos.Y = 0 Then
                    .TargetNPC = 0
                    Exit Sub
                End If
                Call AiTargetNpc(BotIndex, UsaMano, UsaHechizo)
            Else
                Call BotBuscarNPCCercano(NpcIndex)
            End If
        Else
            If UsaMano = False Then
                If Distancia(Npclist(.NpcIndex).Pos, Npclist(.TargetNPC).Pos) = 1 Then
                    tHeading = FindDirection(Npclist(.NpcIndex).Pos, UserList(.TargetUser).Pos)
                    If Not Npclist(.NpcIndex).Char.heading = tHeading Then
                        ChangeBotDireccion BotIndex, tHeading
                    Else
                        Call AtacaBotToNpc(BotIndex)
                    End If
                End If
            End If
        End If
    End With
        
'Exit Sub

'ErrorHandler:
 '   Call LogError("BotAI " & Npclist(NpcIndex).name & " Err:" & Err.description)
End Sub
Function BotRemueve(ByVal BotIndex As Integer) As Boolean
With Bots(BotIndex)
    If Npclist(.NpcIndex).flags.Paralizado Then
        If (.MinMAN - 400) >= 0 And Npclist(.NpcIndex).Contadores.Paralisis < 485 And RandomNumber(1, 5) > 2 Then
            Npclist(.NpcIndex).flags.Inmovilizado = 0
            Npclist(.NpcIndex).flags.Paralizado = 0
            Npclist(.NpcIndex).Contadores.Paralisis = 0
            Call SendData(SendTarget.ToNPCArea, .NpcIndex, PrepareMessageChatOverHead("AN HOAX VORP", Npclist(.NpcIndex).Char.CharIndex, RGB(128, 128, 0)))
            Call SendData(SendTarget.ToNPCArea, .NpcIndex, PrepareMessagePlayWave(110, Npclist(.NpcIndex).Pos.X, Npclist(.NpcIndex).Pos.Y))
            Call SendData(SendTarget.ToNPCArea, .NpcIndex, PrepareMessageCreateCharParticle(Npclist(.NpcIndex).Char.CharIndex, 45, 750))
            .MinMAN = .MinMAN - 400
            BotRemueve = True
            Exit Function
        End If
    End If
End With
BotRemueve = False
End Function
Function BotPotea(ByVal BotIndex As Integer) As Boolean
Dim UsaMano As Boolean

With Bots(BotIndex)
    If .MinHP < .MaxHP Then
        .MinHP = .MinHP + 30
        If .MinHP > .MaxHP Then .MinHP = .MaxHP
        Call SendData(SendTarget.ToNPCArea, .NpcIndex, PrepareMessagePlayWave(SND_BEBER, Npclist(.NpcIndex).Pos.X, Npclist(.NpcIndex).Pos.Y))
        UsaMano = True
    End If
    
    If .TargetNPC Then
        If Npclist(.TargetNPC).flags.Paralizado = 1 Then
            If Distancia(Npclist(.TargetNPC).Pos, Npclist(.NpcIndex).Pos) = 1 Then
                If AtacaClase(.Clase) < 3 Then
                    Exit Function
                End If
            End If
        End If
    End If
    
    If .MinMAN < .MaxMAN And Not UsaMano Then
        .MinMAN = .MinMAN + Porcentaje(.MaxMAN, 5)
        If .MinMAN > .MaxMAN Then .MinMAN = .MaxMAN
        Call SendData(SendTarget.ToNPCArea, .NpcIndex, PrepareMessagePlayWave(SND_BEBER, Npclist(.NpcIndex).Pos.X, Npclist(.NpcIndex).Pos.Y))
        UsaMano = True
    End If
End With

BotPotea = UsaMano

End Function
Sub BotBuscarUserCercano(ByVal NpcIndex As Integer)
Dim UI As Integer
Dim tHeading As eHeading
Dim i As Long
With Npclist(NpcIndex)
    'Buscamos un usuario
'    For i = 1 To ModAreas.ConnGroups(.Pos.map).CountEntrys
'        UI = ModAreas.ConnGroups(.Pos.map).UserEntrys(i)
'        'Is it in it's range of vision??
'        If Abs(UserList(UI).Pos.X - .Pos.X) <= RANGO_VISION_X Then
'            If Abs(UserList(UI).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
'                If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.invisible = 0 And UserList(UI).flags.Oculto = 0 Then ' And Not UserList(UI).flags.Privilegios > User Then
'                    .TargetUser = UI
'                    tHeading = FindDirection(.Pos, UserList(UI).Pos)
'                    Call MoveNPCChar(NpcIndex, tHeading)
'                    Exit Sub
'                End If
'            End If
'        End If
'    Next i
End With
End Sub
Sub BotBuscarNPCCercano(ByVal NpcIndex As Integer)
Dim UI As Integer
Dim X As Integer, Y As Integer
Dim tHeading As eHeading
Dim i As Long
With Npclist(NpcIndex)
    For Y = (.Pos.Y - 3) To (.Pos.Y + 3)
        For X = (.Pos.X - 4) To (.Pos.X + 4)
            If MapData(.Pos.map, X, Y).NpcIndex Then
                'If Not MapData(.Pos.map, X, Y).NpcIndex = NpcIndex And Npclist(MapData(.Pos.map, X, Y).NpcIndex).Bot <> 0 Then
                    'If Bots(Npclist(MapData(.Pos.map, X, Y).NpcIndex).Bot).Active > 0 Then
                '        Bots(.Bot).TargetNPC = MapData(.Pos.map, X, Y).NpcIndex
                '        tHeading = FindDirection(.Pos, Npclist(MapData(.Pos.map, X, Y).NpcIndex).Pos)
                '        Call MoveNPCChar(NpcIndex, tHeading)
                '        Exit Sub
                    'End If
                'End If
            End If
        Next X
    Next Y
End With
End Sub
Sub LanzarSpellBotToUser(ByVal NpcIndex As Integer, ByVal UserVic As Integer)
Dim uh As Byte
Dim UhParalizar As Byte

With Npclist(NpcIndex)
    If .flags.LanzaSpells Then
        'Elegimos 1
        uh = RandomNumber(1, .flags.LanzaSpells)
        'Si es paralizar y ya está paralizado elegimos otro
        If BotHechizos(uh).Paraliza = 1 And UserList(UserVic).flags.Paralizado Then
            UhParalizar = uh
            Do While uh = UhParalizar
                uh = RandomNumber(1, .flags.LanzaSpells)
                DoEvents
            Loop
        End If
        
        Dim daño As Integer
        
        If BotHechizos(uh).SubeHP = 1 Then
        
            daño = RandomNumber(BotHechizos(uh).MinHP, BotHechizos(uh).MaxHP)
            Call SendData(SendTarget.ToPCArea, UserVic, PrepareMessagePlayWave(BotHechizos(uh).WAV, UserList(UserVic).Pos.X, UserList(UserVic).Pos.Y))
            If BotHechizos(uh).FXgrh <> 0 Then Call SendData(SendTarget.ToPCArea, UserVic, PrepareMessageCreateFX(UserList(UserVic).Char.CharIndex, BotHechizos(uh).FXgrh, BotHechizos(uh).loops))
            If BotHechizos(uh).Particle <> 0 Then Call SendData(SendTarget.ToPCArea, UserVic, PrepareMessageCreateCharParticle(UserList(UserVic).Char.CharIndex, BotHechizos(uh).Particle, BotHechizos(uh).life))
            
            UserList(UserVic).Stats.MinHP = UserList(UserVic).Stats.MinHP + daño
            If UserList(UserVic).Stats.MinHP > UserList(UserVic).Stats.MaxHP Then UserList(UserVic).Stats.MinHP = UserList(UserVic).Stats.MaxHP
            
            Call WriteConsoleMsg(UserVic, Npclist(NpcIndex).name & " te ha quitado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
            Call WriteUpdateUserStats(UserVic)
        
        ElseIf BotHechizos(uh).SubeHP = 2 Then
            
                daño = RandomNumber(BotHechizos(uh).MinHP, BotHechizos(uh).MaxHP)
                
                If UserList(UserVic).Invent.CascoEqpObjIndex > 0 Then
                    daño = daño - RandomNumber(ObjData(UserList(UserVic).Invent.CascoEqpObjIndex).DefensaMagicaMin, ObjData(UserList(UserVic).Invent.CascoEqpObjIndex).DefensaMagicaMax)
                End If
                
                If UserList(UserVic).Invent.AnilloEqpObjIndex > 0 Then
                    daño = daño - RandomNumber(ObjData(UserList(UserVic).Invent.AnilloEqpObjIndex).DefensaMagicaMin, ObjData(UserList(UserVic).Invent.AnilloEqpObjIndex).DefensaMagicaMax)
                End If
                
                If daño < 0 Then daño = 0
                
                Call SendData(SendTarget.ToPCArea, UserVic, PrepareMessagePlayWave(BotHechizos(uh).WAV, UserList(UserVic).Pos.X, UserList(UserVic).Pos.Y))
                If BotHechizos(uh).FXgrh <> 0 Then Call SendData(SendTarget.ToPCArea, UserVic, PrepareMessageCreateFX(UserList(UserVic).Char.CharIndex, BotHechizos(uh).FXgrh, BotHechizos(uh).loops))
                If BotHechizos(uh).Particle <> 0 Then Call SendData(SendTarget.ToPCArea, UserVic, PrepareMessageCreateCharParticle(UserList(UserVic).Char.CharIndex, BotHechizos(uh).Particle, BotHechizos(uh).life))
            
                UserList(UserVic).Stats.MinHP = UserList(UserVic).Stats.MinHP - daño
                
                Call WriteConsoleMsg(UserVic, Npclist(NpcIndex).name & " te ha quitado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
                Call WriteUpdateUserStats(UserVic)
                
                'Muere
                If UserList(UserVic).Stats.MinHP < 1 Then
                    UserList(UserVic).Stats.MinHP = 0
                    Call UserDie(UserVic)
                End If
            
        End If
        
        If BotHechizos(uh).Paraliza = 1 Or BotHechizos(uh).Inmoviliza = 1 Then
            If UserList(UserVic).flags.Paralizado = 0 Then
                Call SendData(SendTarget.ToPCArea, UserVic, PrepareMessagePlayWave(BotHechizos(uh).WAV, UserList(UserVic).Pos.X, UserList(UserVic).Pos.Y))
                If BotHechizos(uh).FXgrh <> 0 Then Call SendData(SendTarget.ToPCArea, UserVic, PrepareMessageCreateFX(UserList(UserVic).Char.CharIndex, BotHechizos(uh).FXgrh, BotHechizos(uh).loops))
                If BotHechizos(uh).Particle <> 0 Then Call SendData(SendTarget.ToPCArea, UserVic, PrepareMessageCreateCharParticle(UserList(UserVic).Char.CharIndex, BotHechizos(uh).Particle, BotHechizos(uh).life))
                  
                If UserList(UserVic).Invent.AnilloEqpObjIndex = SUPERANILLO Then
                    Call WriteConsoleMsg(UserVic, " Tu anillo rechaza los efectos del hechizo.", FontTypeNames.FONTTYPE_FIGHT)
                    Exit Sub
                End If
                
                If BotHechizos(uh).Inmoviliza = 1 Then
                    UserList(UserVic).flags.Inmovilizado = 1
                End If
                  
                UserList(UserVic).flags.Paralizado = 1
                UserList(UserVic).Counters.Paralisis = IntervaloParalizado
                  
                Call WriteParalizeOK(UserVic)
            End If
        End If
        
        If BotHechizos(uh).Estupidez = 1 Then
             If UserList(UserVic).flags.Estupidez = 0 Then
                  Call SendData(SendTarget.ToPCArea, UserVic, PrepareMessagePlayWave(BotHechizos(uh).WAV, UserList(UserVic).Pos.X, UserList(UserVic).Pos.Y))
                  If BotHechizos(uh).FXgrh <> 0 Then Call SendData(SendTarget.ToPCArea, UserVic, PrepareMessageCreateFX(UserList(UserVic).Char.CharIndex, BotHechizos(uh).FXgrh, BotHechizos(uh).loops))
                  If BotHechizos(uh).Particle <> 0 Then Call SendData(SendTarget.ToPCArea, UserVic, PrepareMessageCreateCharParticle(UserList(UserVic).Char.CharIndex, BotHechizos(uh).Particle, BotHechizos(uh).life))
                  
                    If UserList(UserVic).Invent.AnilloEqpObjIndex = SUPERANILLO Then
                        Call WriteConsoleMsg(UserVic, " Tu anillo rechaza los efectos del hechizo.", FontTypeNames.FONTTYPE_FIGHT)
                        Exit Sub
                    End If
                  
                  UserList(UserVic).flags.Estupidez = 1
                  UserList(UserVic).Counters.Ceguera = IntervaloInvisible
                          
                Call WriteDumb(UserVic)
             End If
        End If
    End If
End With
    
End Sub
Sub LanzarSpellBotToNPC(ByVal BotIndex As Integer, ByVal botVic As Integer)
Dim uh As Byte
Dim UhParalizar As Byte

With Bots(BotIndex)
    If Npclist(.NpcIndex).flags.LanzaSpells Then
        'Elegimos 1
        uh = Npclist(.NpcIndex).Spells(RandomNumber(1, Npclist(.NpcIndex).flags.LanzaSpells))

        If .MinMAN < BotHechizos(uh).ManaRequerido Then   'And Npclist(botVic).flags.Paralizado Then
            Exit Sub
        End If
        
        Dim daño As Integer
        
        If BotHechizos(uh).SubeHP = 1 Then
        
            daño = RandomNumber(BotHechizos(uh).MinHP, BotHechizos(uh).MaxHP)
            
            Call SendData(SendTarget.ToBotArea, botVic, PrepareMessagePlayWave(BotHechizos(uh).WAV, Bots(botVic).Pos.X, Npclist(botVic).Pos.Y))
            
            If BotHechizos(uh).FXgrh <> 0 Then Call SendData(SendTarget.ToNPCArea, botVic, PrepareMessageCreateFX(Npclist(botVic).Char.CharIndex, BotHechizos(uh).FXgrh, BotHechizos(uh).loops))
            If BotHechizos(uh).Particle <> 0 Then Call SendData(SendTarget.ToNPCArea, botVic, PrepareMessageCreateCharParticle(Npclist(botVic).Char.CharIndex, BotHechizos(uh).Particle, BotHechizos(uh).life))
            
            Bots(botVic).MinHP = Bots(botVic).MinHP + daño
            If Bots(botVic).MinHP > Bots(botVic).MaxHP Then Bots(botVic).MinHP = Bots(botVic).MaxHP
            
            'Quitamos el Mana
            .MinMAN = .MinMAN = BotHechizos(uh).ManaRequerido
            
            'Decimos Palabras Magicas
            Call SendData(SendTarget.ToNPCArea, .NpcIndex, PrepareMessageChatOverHead(BotHechizos(uh).PalabrasMagicas, Npclist(.NpcIndex).Char.CharIndex, RGB(128, 128, 0)))
        ElseIf BotHechizos(uh).SubeHP = 2 Then
            
            daño = RandomNumber(BotHechizos(uh).MinHP, BotHechizos(uh).MaxHP)
                
            If Bots(botVic).casco > 0 Then
                daño = daño - RandomNumber(ObjData(Bots(botVic).casco).DefensaMagicaMin, ObjData(Bots(botVic).casco).DefensaMagicaMax)
            End If
                
            'If Npclist(botVic).Invent.AnilloEqpObjIndex > 0 Then
            '    daño = daño - RandomNumber(ObjData(Npclist(botVic).Invent.AnilloEqpObjIndex).DefensaMagicaMin, ObjData(Npclist(botVic).Invent.AnilloEqpObjIndex).DefensaMagicaMax)
            'End If
                
            If daño < 0 Then daño = 0
                
            Call SendData(SendTarget.ToNPCArea, botVic, PrepareMessagePlayWave(BotHechizos(uh).WAV, Npclist(botVic).Pos.X, Npclist(botVic).Pos.Y))
            If BotHechizos(uh).FXgrh <> 0 Then Call SendData(SendTarget.ToNPCArea, botVic, PrepareMessageCreateFX(Npclist(botVic).Char.CharIndex, BotHechizos(uh).FXgrh, BotHechizos(uh).loops))
            If BotHechizos(uh).Particle <> 0 Then Call SendData(SendTarget.ToNPCArea, botVic, PrepareMessageCreateCharParticle(Npclist(botVic).Char.CharIndex, BotHechizos(uh).Particle, BotHechizos(uh).life))
            
            Bots(botVic).MinHP = Bots(botVic).MinHP - daño
                
            If Bots(botVic).MinHP < 1 Then
                Bots(botVic).MinHP = 0
                Call MuereNpc(botVic, 0)
            End If
                
            'Quitamos el Mana
            .MinMAN = .MinMAN - BotHechizos(uh).ManaRequerido
            'Decimos Palabras Magicas
            Call SendData(SendTarget.ToNPCArea, .NpcIndex, PrepareMessageChatOverHead(BotHechizos(uh).PalabrasMagicas, Npclist(.NpcIndex).Char.CharIndex, RGB(128, 128, 0)))
        End If
        
        If BotHechizos(uh).Paraliza = 1 Or BotHechizos(uh).Inmoviliza = 1 Then
            If Npclist(botVic).flags.Paralizado = 0 Then
                Call SendData(SendTarget.ToNPCArea, botVic, PrepareMessagePlayWave(BotHechizos(uh).WAV, Npclist(botVic).Pos.X, Npclist(botVic).Pos.Y))
                If BotHechizos(uh).FXgrh <> 0 Then Call SendData(SendTarget.ToNPCArea, botVic, PrepareMessageCreateFX(Npclist(botVic).Char.CharIndex, BotHechizos(uh).FXgrh, BotHechizos(uh).loops))
                If BotHechizos(uh).Particle <> 0 Then Call SendData(SendTarget.ToNPCArea, botVic, PrepareMessageCreateCharParticle(Npclist(botVic).Char.CharIndex, BotHechizos(uh).Particle, BotHechizos(uh).life))
                  
                If Npclist(botVic).Invent.AnilloEqpObjIndex = SUPERANILLO Then
                    Exit Sub
                End If
                
                Npclist(botVic).flags.Inmovilizado = 0
                
                If BotHechizos(uh).Inmoviliza = 1 Then
                    Npclist(botVic).flags.Inmovilizado = 1
                End If
                  
                Npclist(botVic).flags.Paralizado = 1
                Npclist(botVic).Contadores.Paralisis = IntervaloParalizado
                'Quitamos el Mana
                .MinMAN = .MinMAN - BotHechizos(uh).ManaRequerido
                'Decimos Palabras Magicas
                Call SendData(SendTarget.ToNPCArea, .NpcIndex, PrepareMessageChatOverHead(BotHechizos(uh).PalabrasMagicas, Npclist(.NpcIndex).Char.CharIndex, RGB(128, 128, 0)))
            End If
        End If
        
    End If
End With
    
End Sub
Sub AiTargetUser(ByVal BotIndex As Integer, ByVal UsaMano As Boolean, ByVal UsaHechizo As Boolean)
Dim tHeading As eHeading
Dim S As String
Dim AlLado As Boolean

With Bots(BotIndex)
'    tHeading = FindDirection(Npclist(.NpcIndex).Pos, UserList(.TargetUser).Pos)
'    If IntervaloPuedeAtacar(.NpcIndex) Then
'        Select Case Npclist(.NpcIndex).Char.heading
'            Case eHeading.NORTH
'                If MapData(Npclist(.NpcIndex).Pos.map, Npclist(.NpcIndex).Pos.X, Npclist(.NpcIndex).Pos.Y - 1).UserIndex = .TargetUser Then
'                    Call AtacaBotToNpc(NpcIndex)
'                End If
'
'            Case eHeading.EAST
'                If MapData(Npclist(.NpcIndex).Pos.map, Npclist(.NpcIndex).Pos.X + 1, Npclist(.NpcIndex).Pos.Y).UserIndex = .TargetUser Then
'                    Call AtacaBotToUser(NpcIndex)
'                End If
'
'            Case eHeading.SOUTH
'                If MapData(Npclist(.NpcIndex).Pos.map, Npclist(.NpcIndex).Pos.X, Npclist(.NpcIndex).Pos.Y + 1).UserIndex = .TargetUser Then
'                    Call AtacaBotToUser(NpcIndex)
'                End If
'
'            Case eHeading.WEST
'                If MapData(Npclist(.NpcIndex).Pos.map, Npclist(.NpcIndex).Pos.X - 1, Npclist(.NpcIndex).Pos.Y).UserIndex = .TargetUser Then
'                    Call AtacaBotToUser(NpcIndex)
'                End If
'        End Select
'    End If
'
'    If IntervaloPuedeHechizo(NpcIndex) Then
'        If RandomNumber(1, 2) = 1 Then
'            If Not UsaHechizo And .MaxMAN > 0 Then
'                Call LanzarSpellBotToUser(NpcIndex, .TargetUser)
'            End If
'        End If
'    End If
'
'    If UserList(.TargetUser).flags.Paralizado Then
'        If Distancia(.Pos, UserList(.TargetUser).Pos) = 1 Then
'            If RandomNumber(1, 5) = 2 Then
'                Call MoveNPCChar(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
'            Else
'                ChangeNPCChar NpcIndex, .Char.body, .Char.Head, tHeading
'            End If
'        Else
'            Call MoveNPCChar(NpcIndex, tHeading)
'        End If
'        Exit Sub
'    End If
'
'    If Distancia(.Pos, UserList(.TargetUser).Pos) > 4 Then
'        Call MoveNPCChar(NpcIndex, tHeading)
'    Else
'        Dim Direccion As Byte
'        Direccion = CByte(RandomNumber(eHeading.NORTH, eHeading.WEST))
'        If Not Direccion = .RandomDire Then
''            .RandomDire = Direccion
 '       Else
 '           .RandomDire = CByte(RandomNumber(eHeading.NORTH, eHeading.WEST))
 '           .RandomDire = Direccion
 '       End If
 '       Call MoveNPCChar(NpcIndex, Direccion)
 '   End If
End With
End Sub
Sub AiTargetNpc(ByVal BotIndex As Integer, ByVal UsaMano As Boolean, ByVal UsaHechizo As Boolean)
Dim tHeading As eHeading
Dim S As String
Dim AlLado As Boolean
With Bots(BotIndex)
    
    tHeading = FindDirection(Npclist(.NpcIndex).Pos, Npclist(.TargetNPC).Pos)
    If IntervaloPuedeAtacar(BotIndex) And Not UsaMano Then
        Select Case Npclist(.NpcIndex).Char.heading
            Case eHeading.NORTH
                If MapData(Npclist(.NpcIndex).Pos.map, Npclist(.NpcIndex).Pos.X, Npclist(.NpcIndex).Pos.Y - 1).NpcIndex = .TargetNPC Then
                    Call AtacaBotToNpc(BotIndex)
                End If
        
            Case eHeading.EAST
                If MapData(Npclist(.NpcIndex).Pos.map, Npclist(.NpcIndex).Pos.X + 1, Npclist(.NpcIndex).Pos.Y).NpcIndex = .TargetNPC Then
                    If Bots(MapData(Bots(.NpcIndex).Pos.map, Bots(.NpcIndex).Pos.X + 1, Bots(.NpcIndex).Pos.Y).BotIndex).Active Then
                        Call AtacaBotToNpc(BotIndex)
                    Else
                        Call MuereNpc(MapData(Npclist(.NpcIndex).Pos.map, Npclist(.NpcIndex).Pos.X + 1, Npclist(.NpcIndex).Pos.Y).NpcIndex, 0)
                    End If
                End If
                    
            Case eHeading.SOUTH
                If MapData(Npclist(.NpcIndex).Pos.map, Npclist(.NpcIndex).Pos.X, Npclist(.NpcIndex).Pos.Y + 1).NpcIndex = .TargetNPC Then
                    Call AtacaBotToNpc(BotIndex)
                End If
                        
            Case eHeading.WEST
                If MapData(Npclist(.NpcIndex).Pos.map, Npclist(.NpcIndex).Pos.X - 1, Npclist(.NpcIndex).Pos.Y).NpcIndex = .TargetNPC Then
                    Call AtacaBotToNpc(BotIndex)
                End If
        End Select
    End If
    
    If IntervaloPuedeHechizo(BotIndex) And Not UsaHechizo Then
        If RandomNumber(1, 2) = 1 Then
            If Not UsaHechizo And .MaxMAN > 0 Then
                Call LanzarSpellBotToNPC(BotIndex, .TargetNPC)
            End If
        End If
    End If
    'listo
          
    If Npclist(.TargetNPC).flags.Paralizado Then
        If Distancia(Npclist(.NpcIndex).Pos, Npclist(.TargetNPC).Pos) = 1 Then
            If Not Npclist(.NpcIndex).Char.heading = tHeading Then ChangeBotDireccion BotIndex, tHeading
        Else
            Call MoveNPCChar(.NpcIndex, tHeading)
        End If
        Exit Sub
    End If
                
    If Distancia(Npclist(.NpcIndex).Pos, Npclist(.TargetNPC).Pos) > 4 Then
        Call MoveNPCChar(.NpcIndex, tHeading)
    Else
        Dim Direccion As Byte
        Direccion = CByte(RandomNumber(eHeading.NORTH, eHeading.WEST))
        If Not Direccion = .RandomDire Then
            .RandomDire = Direccion
        Else
            .RandomDire = CByte(RandomNumber(eHeading.NORTH, eHeading.WEST))
            .RandomDire = Direccion
        End If
        Call MoveNPCChar(.NpcIndex, Direccion)
    End If
End With
End Sub
Sub AtacaBotToNpc(ByVal BotIndex As Integer)
    Dim Victima As Integer
    Dim Atacante As Integer
    Dim VictimaB As Integer
    Dim AtacanteB As Integer
    
    With Bots(BotIndex)
    
        VictimaB = Npclist(.TargetNPC).Bot
        AtacanteB = BotIndex
        
        Victima = .TargetNPC
        Atacante = .NpcIndex
        
        If Npclist(.NpcIndex).flags.Snd1 > 0 Then
            Call SendData(SendTarget.ToNPCArea, Atacante, PrepareMessagePlayWave(Npclist(.NpcIndex).flags.Snd1, Npclist(.NpcIndex).Pos.X, Npclist(.NpcIndex).Pos.Y))
        End If
        
        If BotImpactoBot(AtacanteB, VictimaB) Then
            If Npclist(Victima).flags.Snd2 > 0 Then
                Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(Npclist(Victima).flags.Snd2, Npclist(Victima).Pos.X, Npclist(Victima).Pos.Y))
            Else
                Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(SND_IMPACTO2, Npclist(Victima).Pos.X, Npclist(Victima).Pos.Y))
            End If
        
            Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(SND_IMPACTO, Npclist(Victima).Pos.X, Npclist(Victima).Pos.Y))
            
            Call BotDañoBot(AtacanteB, VictimaB)
        Else
            Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(SND_SWING, Npclist(Victima).Pos.X, Npclist(Victima).Pos.Y))
        End If
    End With
End Sub
Sub AtacaBotToUser(ByVal NpcIndex As Integer)
    Dim Victima As Integer
    Dim Atacante As Integer
    
    With Npclist(NpcIndex)
    
        'Victima = .TargetUser
        'Atacante = NpcIndex
       '
       ' If .flags.Snd1 > 0 Then
       '     Call SendData(SendTarget.ToNPCArea, Atacante, PrepareMessagePlayWave(.flags.Snd1, .Pos.X, .Pos.Y))
       ' End If
       '
       ' If NpcImpacto(Atacante, Victima) Then
       '     Call SendData(SendTarget.ToPCArea, Victima, PrepareMessagePlayWave(SND_IMPACTO2, UserList(Victima).Pos.X, UserList(Victima).Pos.Y))
       '     Call SendData(SendTarget.ToPCArea, Victima, PrepareMessagePlayWave(SND_IMPACTO, UserList(Victima).Pos.X, UserList(Victima).Pos.Y))
       '
       '     Call NpcDaño(Atacante, Victima)
       ' Else
       '     Call SendData(SendTarget.ToPCArea, Victima, PrepareMessagePlayWave(SND_SWING, UserList(Victima).Pos.X, UserList(Victima).Pos.Y))
       ' End If
    End With
End Sub
Private Function IntervaloPuedeHechizo(ByVal BotIndex As Integer) As Boolean
Dim TActual As Long

If RandomNumber(1, 10) > 2 Then
    IntervaloPuedeHechizo = False
    Exit Function
End If

TActual = GetTickCount() And &H7FFFFFFF

If TActual - Bots(BotIndex).IntervaloHechizo >= IntervaloUserPuedeCastear Then
    Bots(BotIndex).IntervaloHechizo = TActual
    IntervaloPuedeHechizo = True
Else
    IntervaloPuedeHechizo = False
End If

End Function
Private Function IntervaloPuedeAtacar(ByVal BotIndex As Integer) As Boolean
Dim TActual As Long

TActual = GetTickCount() And &H7FFFFFFF

If TActual - Bots(BotIndex).IntervaloAtaque >= IntervaloUserPuedeAtacar Then
    Bots(BotIndex).IntervaloAtaque = TActual
    IntervaloPuedeAtacar = True
Else
    IntervaloPuedeAtacar = False
End If

End Function

Public Function ChangeBotDireccion(ByVal BotIndex As Integer, ByVal heading As Byte)
Dim ArmaAnim    As Integer
Dim BodyAnim    As Integer
Dim ShieldAnim  As Integer
Dim CascoAnim   As Integer
Dim Head        As Integer
Dim CharIndex   As Integer

With Bots(BotIndex)
    CascoAnim = 0
    BodyAnim = 0
    ArmaAnim = 0
    ShieldAnim = 0
    
    If .Arma > 0 Then
        ArmaAnim = ObjData(.Arma).WeaponAnim
    End If
    
    If .Armadura > 0 Then
        BodyAnim = ObjData(.Armadura).Ropaje
    End If
    
    If .Escudo > 0 Then
        ShieldAnim = ObjData(.Escudo).ShieldAnim
    End If
    
    If .casco > 0 Then
        CascoAnim = ObjData(.casco).CascoAnim
    End If
    
    Head = Npclist(.NpcIndex).Char.Head
    CharIndex = Npclist(.NpcIndex).Char.CharIndex
    
    Call SendData(SendTarget.ToNPCArea, .NpcIndex, _
            PrepareMessageCharacterChange(BodyAnim, _
            Head, heading, CharIndex, _
            ArmaAnim, ShieldAnim, _
            0, 0, CascoAnim))
End With

End Function

