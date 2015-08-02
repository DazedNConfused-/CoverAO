Attribute VB_Name = "ModBotStats"
Option Explicit
Public Type tBots
    Active As Byte
    
    Arma As Integer
    Armadura As Integer
    casco As Integer
    Escudo As Integer
    
    MinHIT As Byte
    MaxHIT As Byte
    
    MinMAN As Integer
    MaxMAN As Integer
    
    MinHP As Integer
    MaxHP As Integer

    Fuerza As Byte
    Agilidad As Byte
    
    NpcIndex As Integer
    
    TargetUser As Integer
    TargetNPC As Integer
    
    IntervaloAtaque As Long
    IntervaloHechizo As Long
    
    RandomDire As Byte
    
    skills(1 To NUMSKILLS) As Byte
    
    Nivel As Byte
    
    Clase As eClass
    
    Pos As WorldPos
End Type

Public Bots() As tBots
Public NumBots As Integer
Public Sub CrearBot(ByVal botNum As Byte, ByVal NpcIndex)
    
    Dim BotIndex As Integer
    BotIndex = NextOpenBot
    
    Dim Leer As New clsIniReader
    Leer.Initialize App.Path & "\Bots\Bot" & botNum & ".bts"
    
    With Bots(BotIndex)
        .Active = 1

        .Arma = val(Leer.GetValue("BOT", "Arma"))
        .Armadura = val(Leer.GetValue("BOT", "Armadura"))
        .casco = val(Leer.GetValue("BOT", "Casco"))
        .Escudo = val(Leer.GetValue("BOT", "Escudo"))
        .MinHIT = val(Leer.GetValue("BOT", "MinHIT"))
        .MaxHIT = val(Leer.GetValue("BOT", "MaxHIT"))
        
        .MaxMAN = val(Leer.GetValue("STATS", "Mana"))
        .MaxHP = val(Leer.GetValue("STATS", "Vida"))
        .MinHP = .MaxHP
        .Fuerza = val(Leer.GetValue("STATS", "Fuerza"))
        .Agilidad = val(Leer.GetValue("STATS", "Agilidad"))
        .Nivel = val(Leer.GetValue("STATS", "Nivel"))
        .Clase = val(Leer.GetValue("STATS", "Clase"))
        
        Dim i As Byte
        For i = 1 To NUMSKILLS
            .skills(i) = 100 ' Leer.GetValue("SKILLS", "SK" & i)
        Next i
    
    End With
    
End Sub
Public Function DeleteBot(ByVal BotIndex As Integer)
    Exit Function
    With Bots(BotIndex)
        .Active = 0
        .NpcIndex = 0
        
        .Arma = 0
        .Armadura = 0
        .casco = 0
        .Escudo = 0
        .MinHIT = 0
        .MaxHIT = 0
        
        .MaxMAN = 0
        .MaxHP = 0
        .MinHP = 0
        .Fuerza = 0
        .Agilidad = 0
        .Nivel = 0
        .Clase = 0
        
        Dim i As Byte
        For i = 1 To NUMSKILLS
            .skills(i) = 0
        Next i
        
    End With
End Function
Public Function NextOpenBot() As Integer
    Dim i As Integer
    Dim Active As Byte
    
    'If Not NumBots = 0 Then
    '    For i = 1 To NumBots
    '        Active = Bots(i).Active
    '        If Active = 0 Then
    '            NextOpenBot = i
    '            Exit Function
    '        End If
    '    Next i
    'End If
    
    NumBots = NumBots + 1
    ReDim Bots(1 To NumBots) As tBots
    NextOpenBot = NumBots
    
End Function

'TODO:Esto nose que hace aca..Codigo de Ao xd
Public Function LanzaClase(ByVal Clase As eClass) As Byte
    Select Case Clase
        Case eClass.Mage, eClass.Druid, eClass.Cleric
            LanzaClase = 1
        Case eClass.Assasin, eClass.Paladin
            LanzaClase = 2
        Case Else
            LanzaClase = 0
    End Select
End Function
Public Function AtacaClase(ByVal Clase As eClass) As Byte
    Select Case Clase
        Case eClass.Warrior, eClass.Pirat
            AtacaClase = 1
        Case eClass.Paladin, eClass.Assasin
            AtacaClase = 2
        Case eClass.Cleric, eClass.Druid
            AtacaClase = 3
        Case Else
            AtacaClase = 0
    End Select
End Function
Public Function MakeBotChar(ByVal BotIndex As Integer, ByVal X As Byte, ByVal Y As Byte, ByVal sndIndex As Integer, ByVal NpcIndex As Integer)
Dim ArmaAnim    As Integer
Dim BodyAnim    As Integer
Dim ShieldAnim  As Integer
Dim CascoAnim   As Integer
Dim Head        As Integer
Dim heading     As Byte
Dim CharIndex   As Integer

With Bots(BotIndex)
    .NpcIndex = NpcIndex
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
    heading = Npclist(.NpcIndex).Char.heading
    CharIndex = Npclist(.NpcIndex).Char.CharIndex
    
    Call WriteCharacterCreate(sndIndex, BodyAnim, _
                            Head, heading, CharIndex, X, Y, _
                            ArmaAnim, ShieldAnim, _
                            0, 0, CascoAnim, vbNullString, 0, 0)
End With
End Function
