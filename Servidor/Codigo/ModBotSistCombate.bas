Attribute VB_Name = "ModBotSistCombate"
'**************************************************************
'**************************************************************
'**************************************************************
'**************************************************************
'**************************************************************
Public Function BotImpactoBot(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer) As Boolean
    Dim ProbRechazo As Long
    Dim Rechazo As Boolean
    Dim ProbExito As Long
    Dim PoderAtaque As Long
    Dim PoderEvasion As Long
    Dim PoderEvasionEscudo As Long
    Dim Arma As Integer
    Dim SkillTacticas As Long
    Dim SkillDefensa As Long
    
    SkillTacticas = Bots(VictimaIndex).skills(eSkill.Tacticas)
    SkillDefensa = Bots(VictimaIndex).skills(eSkill.Defensa)
    
    Arma = Bots(AtacanteIndex).Arma
    
    'Calculamos el poder de evasion...
    UserPoderEvasion = BotPoderEvasion(VictimaIndex)
    
    If Bots(VictimaIndex).Escudo > 0 Then
        PoderEvasionEscudo = BotPoderEvasionEscudo(VictimaIndex)
        PoderEvasion = PoderEvasion + PoderEvasionEscudo
    Else
        PoderEvasionEscudo = 0
    End If
    
    'Esta usando un arma ???
    If Bots(AtacanteIndex).Arma > 0 Then
        If ObjData(Arma).proyectil = 1 Then
            PoderAtaque = BotPoderAtaqueProyectil(AtacanteIndex)
        Else
            PoderAtaque = BotPoderAtaqueArma(AtacanteIndex)
        End If
    Else
        PoderAtaque = BotPoderAtaqueWrestling(AtacanteIndex)
    End If
    
    ' Chances are rounded
    ProbExito = MaximoInt(10, MinimoInt(90, 50 + (PoderAtaque - PoderEvasion) * 0.4))
    
    BotImpactoBot = (RandomNumber(1, 100) <= ProbExito)
    
    If Bots(VictimaIndex).Escudo > 0 Then
        'Fallo ???
        If Not BotImpactoBot Then
            ' Chances are rounded
            ProbRechazo = MaximoInt(10, MinimoInt(90, 100 * SkillDefensa / (SkillDefensa + SkillTacticas)))
            Rechazo = (RandomNumber(1, 100) <= ProbRechazo)
            If Rechazo = True Then
                'Se rechazo el ataque con el escudo
                Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessagePlayWave(SND_ESCUDO, Npclist(Bots(VictimaIndex).NpcIndex).Pos.X, Npclist(Bots(VictimaIndex).NpcIndex).Pos.Y))
                
                'Call SubirSkill(VictimaIndex, Defensa)
            End If
        End If
    End If
    
End Function
Public Function BotPoderEvasionEscudo(ByVal BotIndex As Integer) As Long
    PoderEvasionEscudo = (Bots(BotIndex).skills(eSkill.Defensa) * ModClase(Bots(BotIndex).Clase).Evasion) / 2
End Function
Public Function BotPoderEvasion(ByVal BotIndex As Integer) As Long
    Dim lTemp As Long
    With Bots(BotIndex)
        lTemp = (.skills(eSkill.Tacticas) + _
          .skills(eSkill.Tacticas) / 33 * .Agilidad) * ModClase(.Clase + 1).Evasion
       
        PoderEvasion = (lTemp + (2.5 * MaximoInt(.Nivel - 12, 0)))
    End With
End Function
Public Function BotPoderAtaqueArma(ByVal BotIndex As Integer) As Long
    Dim PoderAtaqueTemp As Long
    
    With Bots(BotIndex)
        If .skills(eSkill.Armas) < 31 Then
            PoderAtaqueTemp = .skills(eSkill.Armas) * ModClase(.Clase).AtaqueArmas
        ElseIf .skills(eSkill.Armas) < 61 Then
            PoderAtaqueTemp = (.skills(eSkill.Armas) + .Agilidad) * ModClase(.Clase).AtaqueArmas
        ElseIf .skills(eSkill.Armas) < 91 Then
            PoderAtaqueTemp = (.skills(eSkill.Armas) + 2 * .Agilidad) * ModClase(.Clase).AtaqueArmas
        Else
           PoderAtaqueTemp = (.skills(eSkill.Armas) + 3 * .Agilidad) * ModClase(.Clase).AtaqueArmas
        End If
        
        PoderAtaqueArma = (PoderAtaqueTemp + (2.5 * MaximoInt(50 - 12, 0)))
    End With
End Function
Public Function BotPoderAtaqueProyectil(ByVal BotIndex As Integer) As Long
    Dim PoderAtaqueTemp As Long
    
    With Bots(BotIndex)
        If .skills(eSkill.Proyectiles) < 31 Then
            PoderAtaqueTemp = .skills(eSkill.Proyectiles) * ModClase(.Clase).AtaqueProyectiles
        ElseIf .skills(eSkill.Proyectiles) < 61 Then
            PoderAtaqueTemp = (.skills(eSkill.Proyectiles) + .Agilidad) * ModClase(.Clase).AtaqueProyectiles
        ElseIf .skills(eSkill.Proyectiles) < 91 Then
            PoderAtaqueTemp = (.skills(eSkill.Proyectiles) + 2 * .Agilidad) * ModClase(.Clase).AtaqueProyectiles
        Else
            PoderAtaqueTemp = (.skills(eSkill.Proyectiles) + 3 * .Agilidad) * ModClase(.Clase).AtaqueProyectiles
        End If
        
        PoderAtaqueProyectil = (PoderAtaqueTemp + (2.5 * MaximoInt(.Nivel - 12, 0)))
    End With
End Function

Public Function BotPoderAtaqueWrestling(ByVal BotIndex As Integer) As Long
    Dim PoderAtaqueTemp As Long
    
    With Bots(BotIndex)
        If .skills(eSkill.Wrestling) < 31 Then
            PoderAtaqueTemp = .skills(eSkill.Wrestling) * ModClase(.Clase).AtaqueArmas
        ElseIf .skills(eSkill.Wrestling) < 61 Then
            PoderAtaqueTemp = (.skills(eSkill.Wrestling) + .Agilidad) * ModClase(.Clase).AtaqueArmas
        ElseIf .skills(eSkill.Wrestling) < 91 Then
            PoderAtaqueTemp = (.skills(eSkill.Wrestling) + 2 * .Agilidad) * ModClase(.Clase).AtaqueArmas
        Else
            PoderAtaqueTemp = (.skills(eSkill.Wrestling) + 3 * .Agilidad) * ModClase(.Clase).AtaqueArmas
        End If
        
        PoderAtaqueWrestling = (PoderAtaqueTemp + (2.5 * MaximoInt(.Nivel - 12, 0)))
    End With
End Function
'**************************************************************
'**************************************************************
'**************************************************************
'**************************************************************

Public Sub BotDañoBot(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer)
    Dim daño As Long
    Dim Lugar As Integer
    Dim absorbido As Long
    Dim defbarco As Integer
    Dim Obj As ObjData
    Dim Resist As Byte
    
    daño = CalcularDaño(AtacanteIndex)
    
    With Bots(AtacanteIndex)
        
        defbarco = 0
        
        If .Arma > 0 Then
            Resist = ObjData(.Arma).Refuerzo
        End If
        
        Lugar = RandomNumber(PartesCuerpo.bCabeza, PartesCuerpo.bTorso)
        
        Select Case Lugar
            Case PartesCuerpo.bCabeza
                'Si tiene casco absorbe el golpe
                If Bots(VictimaIndex).casco > 0 Then
                    Obj = ObjData(Bots(VictimaIndex).casco)
                    absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
                    absorbido = absorbido + defbarco - Resist
                    daño = daño - absorbido
                    If daño < 0 Then daño = 1
                End If
            
            Case Else
                'Si tiene armadura absorbe el golpe
                If Bots(VictimaIndex).Armadura > 0 Then
                    Obj = ObjData(Bots(VictimaIndex).Armadura)
                    Dim Obj2 As ObjData
                    If Bots(VictimaIndex).Escudo Then
                        Obj2 = ObjData(Bots(VictimaIndex).Escudo)
                        absorbido = RandomNumber(Obj.MinDef + Obj2.MinDef, Obj.MaxDef + Obj2.MaxDef)
                    Else
                        absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
                    End If
                    absorbido = absorbido + defbarco - Resist
                    daño = daño - absorbido
                    If daño < 0 Then daño = 1
                End If
        End Select
        
        If daño > 145 Then
            Call SendData(SendTarget.ToNPCArea, Bots(AtacanteIndex).NpcIndex, PrepareMessageChatOverHead("¡" & daño & "!", Npclist(Bots(AtacanteIndex).NpcIndex).Char.CharIndex, vbRed))
        Else
            Call SendData(SendTarget.ToNPCArea, Bots(AtacanteIndex).NpcIndex, PrepareMessageChatOverHead(daño, Npclist(Bots(AtacanteIndex).NpcIndex).Char.CharIndex, vbRed))
        End If
        
        Bots(VictimaIndex).MinHP = Bots(VictimaIndex).MinHP - daño
                
        If Bots(VictimaIndex).MinHP <= 0 Then
            Call MuereNpc(Bots(VictimaIndex).NpcIndex, 0)
        End If
        
    End With
    
End Sub
Private Function CalcularDaño(ByVal BotIndex As Integer, Optional ByVal NpcIndex As Integer = 0) As Long
    Dim DañoArma As Long
    Dim DañoUsuario As Long
    Dim Arma As ObjData
    Dim ModifClase As Single
    Dim DañoMaxArma As Long
    
    With Bots(BotIndex)
        If .Arma > 0 Then
            Arma = ObjData(.Arma)
            ModifClase = ModClase(.Clase).DañoArmas
        Else
            ModifClase = ModClase(.Clase).DañoWrestling
            DañoArma = RandomNumber(1, 3) 'Hacemos que sea "tipo" una daga el ataque de Wrestling
            DañoMaxArma = 3
        End If
        
        DañoUsuario = RandomNumber(.MinHIT, .MaxHIT)
        
        CalcularDaño = (3 * DañoArma + ((DañoMaxArma / 5) * MaximoInt(0, .Fuerza - 15)) + DañoUsuario) * ModifClase

    End With
End Function
