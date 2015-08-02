Attribute VB_Name = "Extra"
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
Public Function ClaseToEnum(ByVal Clase As String) As eClass
Dim i As Byte
For i = 1 To NUMCLASES
    If UCase$(ListaClases(i)) = UCase$(Clase) Then
        ClaseToEnum = i
    End If
Next i
End Function
Public Function EsNewbie(ByVal UserIndex As Integer) As Boolean
    EsNewbie = UserList(UserIndex).Stats.ELV <= LimiteNewbie
End Function

Public Function esArmada(ByVal UserIndex As Integer) As Boolean
    esArmada = (UserList(UserIndex).Faccion.ArmadaReal = 1)
End Function
Public Function esCaos(ByVal UserIndex As Integer) As Boolean
    esCaos = (UserList(UserIndex).Faccion.FuerzasCaos = 1)
End Function
Public Function esMili(ByVal UserIndex As Integer) As Boolean
    esMili = (UserList(UserIndex).Faccion.Milicia = 1)
End Function
Public Function esFaccion(ByVal UserIndex As Integer) As Boolean
    esFaccion = (UserList(UserIndex).Faccion.ArmadaReal = 1 Or UserList(UserIndex).Faccion.FuerzasCaos = 1 Or UserList(UserIndex).Faccion.Milicia = 1)
End Function
Public Function criminal(ByVal UserIndex As Integer) As Boolean
    criminal = esRene(UserIndex)
End Function
Public Function esRene(ByVal UserIndex As Integer) As Boolean
    esRene = (UserList(UserIndex).Faccion.Renegado)
End Function
Public Function esCiuda(ByVal UserIndex As Integer) As Boolean
    esCiuda = (UserList(UserIndex).Faccion.Ciudadano)
End Function
Public Function esRepu(ByVal UserIndex As Integer) As Boolean
    esRepu = (UserList(UserIndex).Faccion.Republicano)
End Function

Public Function EsGM(ByVal UserIndex As Integer) As Boolean
'***************************************************
'Autor: Pablo (ToxicWaste)
'Last Modification: 23/01/2007
'***************************************************
    EsGM = (UserList(UserIndex).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero))
End Function

Public Sub DoTileEvents(ByVal UserIndex As Integer, ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer)
'***************************************************
'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
'Last Modification: 23/01/2007
'Handles the Map passage of Users. Allows the existance
'of exclusive maps for Newbies, Royal Army and Caos Legion members
'and enables GMs to enter every map without restriction.
'Uses: Mapinfo(map).Restringir = "NEWBIE" (newbies), "ARMADA", "CAOS", "FACCION" or "NO".
'***************************************************
    Dim nPos As WorldPos
    Dim FxFlag As Boolean
    
On Error GoTo Errhandler
    'Controla las salidas
    If InMapBounds(map, X, Y) Then
        With MapData(map, X, Y)
            If .ObjInfo.ObjIndex > 0 Then
                FxFlag = ObjData(.ObjInfo.ObjIndex).OBJType = eOBJType.otTeleport
            End If
            
            If .TileExit.map > 0 And .TileExit.map <= NumMaps Then
                '¿Es mapa de newbies?
                If UCase$(MapInfo(.TileExit.map).Restringir) = "NEWBIE" Then
                    '¿El usuario es un newbie?
                    If EsNewbie(UserIndex) Or EsGM(UserIndex) Then
                        If LegalPos(.TileExit.map, .TileExit.X, .TileExit.Y, PuedeAtravesarAgua(UserIndex)) Then
                            Call WarpUserChar(UserIndex, .TileExit.map, .TileExit.X, .TileExit.Y, FxFlag)
                        Else
                            Call ClosestLegalPos(.TileExit, nPos)
                            If nPos.X <> 0 And nPos.Y <> 0 Then
                                Call WarpUserChar(UserIndex, nPos.map, nPos.X, nPos.Y, FxFlag)
                            End If
                        End If
                    Else 'No es newbie
                        Call WriteConsoleMsg(1, UserIndex, "Mapa exclusivo para newbies.", FontTypeNames.FONTTYPE_INFO)
                        Call ClosestStablePos(UserList(UserIndex).Pos, nPos)
        
                        If nPos.X <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(UserIndex, nPos.map, nPos.X, nPos.Y, False)
                        End If
                    End If
                ElseIf UCase$(MapInfo(.TileExit.map).Restringir) = "ARMADA" Then '¿Es mapa de Armadas?
                    '¿El usuario es Armada?
                    If esArmada(UserIndex) Or EsGM(UserIndex) Then
                        If LegalPos(.TileExit.map, .TileExit.X, .TileExit.Y, PuedeAtravesarAgua(UserIndex)) Then
                            Call WarpUserChar(UserIndex, .TileExit.map, .TileExit.X, .TileExit.Y, FxFlag)
                        Else
                            Call ClosestLegalPos(.TileExit, nPos)
                            If nPos.X <> 0 And nPos.Y <> 0 Then
                                Call WarpUserChar(UserIndex, nPos.map, nPos.X, nPos.Y, FxFlag)
                            End If
                        End If
                    Else 'No es armada
                        Call WriteConsoleMsg(1, UserIndex, "Mapa exclusivo para miembros del ejército Real", FontTypeNames.FONTTYPE_INFO)
                        Call ClosestStablePos(UserList(UserIndex).Pos, nPos)
                        
                        If nPos.X <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(UserIndex, nPos.map, nPos.X, nPos.Y, FxFlag)
                        End If
                    End If
                ElseIf UCase$(MapInfo(.TileExit.map).Restringir) = "CAOS" Then '¿Es mapa de Caos?
                    '¿El usuario es Caos?
                    If esCaos(UserIndex) Or EsGM(UserIndex) Then
                        If LegalPos(.TileExit.map, .TileExit.X, .TileExit.Y, PuedeAtravesarAgua(UserIndex)) Then
                            Call WarpUserChar(UserIndex, .TileExit.map, .TileExit.X, .TileExit.Y, FxFlag)
                        Else
                            Call ClosestLegalPos(.TileExit, nPos)
                            If nPos.X <> 0 And nPos.Y <> 0 Then
                                Call WarpUserChar(UserIndex, nPos.map, nPos.X, nPos.Y, FxFlag)
                            End If
                        End If
                    Else 'No es caos
                        Call WriteConsoleMsg(1, UserIndex, "Mapa exclusivo para miembros de la Horda.", FontTypeNames.FONTTYPE_INFO)
                        Call ClosestStablePos(UserList(UserIndex).Pos, nPos)
                        
                        If nPos.X <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(UserIndex, nPos.map, nPos.X, nPos.Y, FxFlag)
                        End If
                    End If
                ElseIf UCase$(MapInfo(.TileExit.map).Restringir) = "FACCION" Then '¿Es mapa de faccionarios?
                    '¿El usuario es Armada o Caos?
                    If esArmada(UserIndex) Or esCaos(UserIndex) Or EsGM(UserIndex) Then
                        If LegalPos(.TileExit.map, .TileExit.X, .TileExit.Y, PuedeAtravesarAgua(UserIndex)) Then
                            Call WarpUserChar(UserIndex, .TileExit.map, .TileExit.X, .TileExit.Y, FxFlag)
                        Else
                            Call ClosestLegalPos(.TileExit, nPos)
                            If nPos.X <> 0 And nPos.Y <> 0 Then
                                Call WarpUserChar(UserIndex, nPos.map, nPos.X, nPos.Y, FxFlag)
                            End If
                        End If
                    Else 'No es Faccionario
                        Call WriteConsoleMsg(1, UserIndex, "Solo se permite entrar al Mapa si eres miembro de alguna Facción", FontTypeNames.FONTTYPE_INFO)
                        Call ClosestStablePos(UserList(UserIndex).Pos, nPos)
                        
                        If nPos.X <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(UserIndex, nPos.map, nPos.X, nPos.Y, FxFlag)
                        End If
                    End If
                Else 'No es un mapa de newbies, ni Armadas, ni Caos, ni faccionario.
                    If LegalPos(.TileExit.map, .TileExit.X, .TileExit.Y, PuedeAtravesarAgua(UserIndex)) Then
                        Call WarpUserChar(UserIndex, .TileExit.map, .TileExit.X, .TileExit.Y, FxFlag)
                    Else
                        Call ClosestLegalPos(.TileExit, nPos)
                        If nPos.X <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(UserIndex, nPos.map, nPos.X, nPos.Y, FxFlag)
                        End If
                    End If
                End If
                
                'Te fusite del mapa. La criatura ya no es más tuya ni te reconoce como que vos la atacaste.
                Dim aN As Integer
                
                aN = UserList(UserIndex).flags.AtacadoPorNpc
                If aN > 0 Then
                   Npclist(aN).Movement = Npclist(aN).flags.OldMovement
                   Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
                   Npclist(aN).flags.AttackedBy = 0
                End If
            
                aN = UserList(UserIndex).flags.NPCAtacado
                If aN > 0 Then
                    If Npclist(aN).flags.AttackedFirstBy = UserList(UserIndex).name Then
                        Npclist(aN).flags.AttackedFirstBy = vbNullString
                    End If
                End If
                UserList(UserIndex).flags.AtacadoPorNpc = 0
                UserList(UserIndex).flags.NPCAtacado = 0
            End If
        End With
    End If
Exit Sub

Errhandler:
    Call LogError("Error en DotileEvents. Error: " & Err.Number & " - Desc: " & Err.description)
End Sub

Function InRangoVision(ByVal UserIndex As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean

If X > UserList(UserIndex).Pos.X - MinXBorder And X < UserList(UserIndex).Pos.X + MinXBorder Then
    If Y > UserList(UserIndex).Pos.Y - MinYBorder And Y < UserList(UserIndex).Pos.Y + MinYBorder Then
        InRangoVision = True
        Exit Function
    End If
End If
InRangoVision = False

End Function

Function InRangoVisionNPC(ByVal NpcIndex As Integer, X As Integer, Y As Integer) As Boolean

If X > Npclist(NpcIndex).Pos.X - MinXBorder And X < Npclist(NpcIndex).Pos.X + MinXBorder Then
    If Y > Npclist(NpcIndex).Pos.Y - MinYBorder And Y < Npclist(NpcIndex).Pos.Y + MinYBorder Then
        InRangoVisionNPC = True
        Exit Function
    End If
End If
InRangoVisionNPC = False

End Function


Function InMapBounds(ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
            
If (map <= 0 Or map > NumMaps) Or X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
    InMapBounds = False
Else
    InMapBounds = True
End If

End Function

Sub ClosestLegalPos(Pos As WorldPos, ByRef nPos As WorldPos, Optional PuedeAgua As Boolean = False, Optional PuedeTierra As Boolean = True)
'*****************************************************************
'Author: Unknown (original version)
'Last Modification: 24/01/2007 (ToxicWaste)
'Encuentra la posicion legal mas cercana y la guarda en nPos
'*****************************************************************

Dim Notfound As Boolean
Dim LoopC As Integer
Dim tX As Long
Dim tY As Long

nPos.map = Pos.map

Do While Not LegalPos(Pos.map, nPos.X, nPos.Y, PuedeAgua, PuedeTierra)
    If LoopC > 12 Then
        Notfound = True
        Exit Do
    End If
    
    For tY = Pos.Y - LoopC To Pos.Y + LoopC
        For tX = Pos.X - LoopC To Pos.X + LoopC
            
            If LegalPos(nPos.map, tX, tY, PuedeAgua, PuedeTierra) Then
                nPos.X = tX
                nPos.Y = tY
                '¿Hay objeto?
                
                tX = Pos.X + LoopC
                tY = Pos.Y + LoopC
  
            End If
        
        Next tX
    Next tY
    
    LoopC = LoopC + 1
    
Loop

If Notfound = True Then
    nPos.X = 0
    nPos.Y = 0
End If

End Sub

Private Sub ClosestStablePos(Pos As WorldPos, ByRef nPos As WorldPos)
'*****************************************************************
'Encuentra la posicion legal mas cercana que no sea un portal y la guarda en nPos
'*****************************************************************

Dim Notfound As Boolean
Dim LoopC As Integer
Dim tX As Long
Dim tY As Long

nPos.map = Pos.map

Do While Not LegalPos(Pos.map, nPos.X, nPos.Y)
    If LoopC > 12 Then
        Notfound = True
        Exit Do
    End If
    
    For tY = Pos.Y - LoopC To Pos.Y + LoopC
        For tX = Pos.X - LoopC To Pos.X + LoopC
            
            If LegalPos(nPos.map, tX, tY) And MapData(nPos.map, tX, tY).TileExit.map = 0 Then
                nPos.X = tX
                nPos.Y = tY
                '¿Hay objeto?
                
                tX = Pos.X + LoopC
                tY = Pos.Y + LoopC
  
            End If
        
        Next tX
    Next tY
    
    LoopC = LoopC + 1
    
Loop

If Notfound = True Then
    nPos.X = 0
    nPos.Y = 0
End If

End Sub

Function NameIndex(ByVal name As String) As Integer
    Dim UserIndex As Long
    
    '¿Nombre valido?
    If LenB(name) = 0 Then
        NameIndex = 0
        Exit Function
    End If
    
    If InStrB(name, "+") <> 0 Then
        name = UCase$(Replace(name, "+", " "))
    End If
    
    UserIndex = 1
    Do Until UCase$(UserList(UserIndex).name) = UCase$(name)
        
        UserIndex = UserIndex + 1
        
        If UserIndex > MaxUsers Then
            NameIndex = 0
            Exit Function
        End If
    Loop
     
    NameIndex = UserIndex
End Function

Function CheckForSameIP(ByVal UserIndex As Integer, ByVal UserIP As String) As Boolean
    Dim LoopC As Long
    
    For LoopC = 1 To MaxUsers
        If UserList(LoopC).flags.UserLogged = True Then
            If UserList(LoopC).ip = UserIP And UserIndex <> LoopC Then
                CheckForSameIP = True
                Exit Function
            End If
        End If
    Next LoopC
    
    CheckForSameIP = False
End Function

Function CheckForSameName(ByVal name As String) As Boolean
'Controlo que no existan usuarios con el mismo nombre
    Dim LoopC As Long
    
    For LoopC = 1 To LastUser
        If UserList(LoopC).flags.UserLogged Then
            
            'If UCase$(UserList(LoopC).Name) = UCase$(Name) And UserList(LoopC).ConnID <> -1 Then
            'OJO PREGUNTAR POR EL CONNID <> -1 PRODUCE QUE UN PJ EN DETERMINADO
            'MOMENTO PUEDA ESTAR LOGUEADO 2 VECES (IE: CIERRA EL SOCKET DESDE ALLA)
            'ESE EVENTO NO DISPARA UN SAVE USER, LO QUE PUEDE SER UTILIZADO PARA DUPLICAR ITEMS
            'ESTE BUG EN ALKON PRODUJO QUE EL SERVIDOR ESTE CAIDO DURANTE 3 DIAS. ATENTOS.
            
            If UCase$(UserList(LoopC).name) = UCase$(name) Then
                CheckForSameName = True
                UserList(LoopC).Counters.Saliendo = True
                UserList(LoopC).Counters.Salir = 1
                Exit Function
            End If
        End If
    Next LoopC
    
    CheckForSameName = False
End Function

Sub HeadtoPos(ByVal Head As eHeading, ByRef Pos As WorldPos)
'*****************************************************************
'Toma una posicion y se mueve hacia donde esta perfilado
'*****************************************************************
    Select Case Head
        Case eHeading.NORTH
            Pos.Y = Pos.Y - 1
        
        Case eHeading.SOUTH
            Pos.Y = Pos.Y + 1
        
        Case eHeading.EAST
            Pos.X = Pos.X + 1
        
        Case eHeading.WEST
            Pos.X = Pos.X - 1
    End Select
End Sub

Function LegalPos(ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal PuedeAgua As Boolean = False, Optional ByVal PuedeTierra As Boolean = True) As Boolean
'***************************************************
'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
'Last Modification: 23/01/2007
'Checks if the position is Legal.
'***************************************************
'¿Es un mapa valido?
If (map <= 0 Or map > NumMaps) Or _
   (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
            LegalPos = False
Else
    If PuedeAgua And PuedeTierra Then
        LegalPos = (MapData(map, X, Y).Blocked <> 1) And _
                   (MapData(map, X, Y).UserIndex = 0) And _
                   (MapData(map, X, Y).NpcIndex = 0)
    ElseIf PuedeTierra And Not PuedeAgua Then
        LegalPos = (MapData(map, X, Y).Blocked <> 1) And _
                   (MapData(map, X, Y).UserIndex = 0) And _
                   (MapData(map, X, Y).NpcIndex = 0) And _
                   (Not HayAgua(map, X, Y))
    ElseIf PuedeAgua And Not PuedeTierra Then
        LegalPos = (MapData(map, X, Y).Blocked <> 1) And _
                   (MapData(map, X, Y).UserIndex = 0) And _
                   (MapData(map, X, Y).NpcIndex = 0) And _
                   (HayAgua(map, X, Y))
    Else
        LegalPos = False
    End If
   
End If

End Function

Function MoveToLegalPos(ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal PuedeAgua As Boolean = False, Optional ByVal PuedeTierra As Boolean = True) As Boolean
'***************************************************
'Autor: ZaMa
'Last Modification: 26/03/2009
'Checks if the position is Legal, but considers that if there's a casper, it's a legal movement.
'***************************************************

Dim UserIndex As Integer
Dim IsDeadChar As Boolean


'¿Es un mapa valido?
If (map <= 0 Or map > NumMaps) Or _
   (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
            MoveToLegalPos = False
    Else
        UserIndex = MapData(map, X, Y).UserIndex
        If UserIndex > 0 Then
            IsDeadChar = UserList(UserIndex).flags.Muerto = 1
        Else
            IsDeadChar = False
        End If
    
    If PuedeAgua And PuedeTierra Then
        MoveToLegalPos = (MapData(map, X, Y).Blocked <> 1) And _
                   (UserIndex = 0 Or IsDeadChar) And _
                   (MapData(map, X, Y).NpcIndex = 0)
    ElseIf PuedeTierra And Not PuedeAgua Then
        MoveToLegalPos = (MapData(map, X, Y).Blocked <> 1) And _
                   (UserIndex = 0 Or IsDeadChar) And _
                   (MapData(map, X, Y).NpcIndex = 0) And _
                   (Not HayAgua(map, X, Y))
    ElseIf PuedeAgua And Not PuedeTierra Then
        MoveToLegalPos = (MapData(map, X, Y).Blocked <> 1) And _
                   (UserIndex = 0 Or IsDeadChar) And _
                   (MapData(map, X, Y).NpcIndex = 0) And _
                   (HayAgua(map, X, Y))
    Else
        MoveToLegalPos = False
    End If
  
End If


End Function

Public Sub FindLegalPos(ByVal UserIndex As Integer, ByVal map As Integer, ByRef X As Integer, ByRef Y As Integer)
'***************************************************
'Autor: ZaMa
'Last Modification: 26/03/2009
'Search for a Legal pos for the user who is being teleported.
'***************************************************

    If MapData(map, X, Y).UserIndex <> 0 Or _
        MapData(map, X, Y).NpcIndex <> 0 Then
                    
        ' Se teletransporta a la misma pos a la que estaba
        If MapData(map, X, Y).UserIndex = UserIndex Then Exit Sub
                            
        Dim FoundPlace As Boolean
        Dim tX As Long
        Dim tY As Long
        Dim Rango As Long
        Dim OtherUserIndex As Integer
    
        For Rango = 1 To 5
            For tY = Y - Rango To Y + Rango
                For tX = X - Rango To X + Rango
                    'Reviso que no haya User ni NPC
                    If MapData(map, tX, tY).UserIndex = 0 And _
                        MapData(map, tX, tY).NpcIndex = 0 Then
                        
                        If InMapBounds(map, tX, tY) Then FoundPlace = True
                        
                        Exit For
                    End If

                Next tX
        
                If FoundPlace Then _
                    Exit For
            Next tY
            
            If FoundPlace Then _
                    Exit For
        Next Rango

    
        If FoundPlace Then 'Si encontramos un lugar, listo, nos quedamos ahi
            X = tX
            Y = tY
        Else
            'Muy poco probable, pero..
            'Si no encontramos un lugar, sacamos al usuario que tenemos abajo, y si es un NPC, lo pisamos.
            OtherUserIndex = MapData(map, X, Y).UserIndex
            If OtherUserIndex <> 0 Then
                'Si no encontramos lugar, y abajo teniamos a un usuario, lo pisamos y cerramos su comercio seguro
                If UserList(OtherUserIndex).ComUsu.DestUsu > 0 Then
                    'Le avisamos al que estaba comerciando que se tuvo que ir.
                    If UserList(UserList(OtherUserIndex).ComUsu.DestUsu).flags.UserLogged Then
                        Call FinComerciarUsu(UserList(OtherUserIndex).ComUsu.DestUsu)
                        Call WriteConsoleMsg(1, UserList(OtherUserIndex).ComUsu.DestUsu, "Comercio cancelado. El otro usuario se ha desconectado.", FontTypeNames.FONTTYPE_TALK)
                        Call FlushBuffer(UserList(OtherUserIndex).ComUsu.DestUsu)
                    End If
                    'Lo sacamos.
                    If UserList(OtherUserIndex).flags.UserLogged Then
                        Call FinComerciarUsu(OtherUserIndex)
                        Call WriteErrorMsg(OtherUserIndex, "Alguien se ha conectado donde te encontrabas, por favor reconéctate...")
                        Call FlushBuffer(OtherUserIndex)
                    End If
                End If
            
                Call CloseSocket(OtherUserIndex)
            End If
        End If
    End If

End Sub

Function LegalPosNPC(ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal AguaValida As Byte) As Boolean
'***************************************************
'Autor: Unkwnown
'Last Modification: 27/04/2009
'Checks if it's a Legal pos for the npc to move to.
'***************************************************
Dim IsDeadChar As Boolean
Dim UserIndex As Integer

    If (map <= 0 Or map > NumMaps) Or _
        (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
        LegalPosNPC = False
        Exit Function
    End If

    UserIndex = MapData(map, X, Y).UserIndex
    If UserIndex > 0 Then
        IsDeadChar = UserList(UserIndex).flags.Muerto = 1
    Else
        IsDeadChar = False
    End If
    
    If AguaValida = 0 Then
        LegalPosNPC = (MapData(map, X, Y).Blocked <> 1) And _
        (MapData(map, X, Y).UserIndex = 0 Or IsDeadChar) And _
        (MapData(map, X, Y).NpcIndex = 0) And _
        (MapData(map, X, Y).Trigger <> eTrigger.POSINVALIDA) _
        And Not HayAgua(map, X, Y)
    Else
        LegalPosNPC = (MapData(map, X, Y).Blocked <> 1) And _
        (MapData(map, X, Y).UserIndex = 0 Or IsDeadChar) And _
        (MapData(map, X, Y).NpcIndex = 0) And _
        (MapData(map, X, Y).Trigger <> eTrigger.POSINVALIDA)
    End If
End Function

Sub SendHelp(ByVal index As Integer)
Dim NumHelpLines As Integer
Dim LoopC As Integer

NumHelpLines = val(GetVar(DatPath & "Help.dat", "INIT", "NumLines"))

For LoopC = 1 To NumHelpLines
    Call WriteConsoleMsg(1, index, GetVar(DatPath & "Help.dat", "Help", "Line" & LoopC), FontTypeNames.FONTTYPE_INFO)
Next LoopC

End Sub

Public Sub Expresar(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
    If Npclist(NpcIndex).NroExpresiones > 0 Then
        Dim randomi
        randomi = RandomNumber(1, Npclist(NpcIndex).NroExpresiones)
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(Npclist(NpcIndex).Expresiones(randomi), Npclist(NpcIndex).Char.CharIndex, vbWhite))
    End If
End Sub

Sub LookatTile(ByVal UserIndex As Integer, ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 26/03/2009
'13/02/2009: ZaMa - EL nombre del gm que aparece por consola al clickearlo, tiene el color correspondiente a su rango
'***************************************************


'Responde al click del usuario sobre el mapa
Dim FoundChar As Byte
Dim FoundSomething As Byte
Dim TempCharIndex As Integer
Dim Stat As String
Dim ft As FontTypeNames

'¿Rango Visión? (ToxicWaste)
If (Abs(UserList(UserIndex).Pos.Y - Y) > RANGO_VISION_Y) Or (Abs(UserList(UserIndex).Pos.X - X) > RANGO_VISION_X) Then
    Exit Sub
End If

'¿Posicion valida?
If InMapBounds(map, X, Y) Then
    UserList(UserIndex).flags.TargetMap = map
    UserList(UserIndex).flags.TargetX = X
    UserList(UserIndex).flags.TargetY = Y
    '¿Es un obj?
    If MapData(map, X, Y).ObjInfo.ObjIndex > 0 Then
        'Informa el nombre
        UserList(UserIndex).flags.TargetObjMap = map
        UserList(UserIndex).flags.TargetObjX = X
        UserList(UserIndex).flags.TargetObjY = Y
        FoundSomething = 1
    ElseIf MapData(map, X + 1, Y).ObjInfo.ObjIndex > 0 Then
        'Informa el nombre
        If ObjData(MapData(map, X + 1, Y).ObjInfo.ObjIndex).OBJType = eOBJType.otPuertas Then
            UserList(UserIndex).flags.TargetObjMap = map
            UserList(UserIndex).flags.TargetObjX = X + 1
            UserList(UserIndex).flags.TargetObjY = Y
            FoundSomething = 1
        End If
    ElseIf MapData(map, X + 1, Y + 1).ObjInfo.ObjIndex > 0 Then
        If ObjData(MapData(map, X + 1, Y + 1).ObjInfo.ObjIndex).OBJType = eOBJType.otPuertas Then
            'Informa el nombre
            UserList(UserIndex).flags.TargetObjMap = map
            UserList(UserIndex).flags.TargetObjX = X + 1
            UserList(UserIndex).flags.TargetObjY = Y + 1
            FoundSomething = 1
        End If
    ElseIf MapData(map, X, Y + 1).ObjInfo.ObjIndex > 0 Then
        If ObjData(MapData(map, X, Y + 1).ObjInfo.ObjIndex).OBJType = eOBJType.otPuertas Then
            'Informa el nombre
            UserList(UserIndex).flags.TargetObjMap = map
            UserList(UserIndex).flags.TargetObjX = X
            UserList(UserIndex).flags.TargetObjY = Y + 1
            FoundSomething = 1
        End If
    End If
    
    If FoundSomething = 1 Then
        UserList(UserIndex).flags.TargetObj = MapData(map, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).ObjInfo.ObjIndex
        If MostrarCantidad(UserList(UserIndex).flags.TargetObj) Then
                        Call WriteConsoleMsg(1, UserIndex, ObjData(UserList(UserIndex).flags.TargetObj).name & " (" & MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).ObjInfo.amount & ")", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(1, UserIndex, ObjData(UserList(UserIndex).flags.TargetObj).name, FontTypeNames.FONTTYPE_INFO)
        End If
    
    End If
    '¿Es un personaje?
    If Y + 1 <= YMaxMapSize Then
        If MapData(map, X, Y + 1).UserIndex > 0 Then
            TempCharIndex = MapData(map, X, Y + 1).UserIndex
            FoundChar = 1
        End If
        If MapData(map, X, Y + 1).NpcIndex > 0 Then
            TempCharIndex = MapData(map, X, Y + 1).NpcIndex
            FoundChar = 2
        End If
    End If
    '¿Es un personaje?
    If FoundChar = 0 Then
        If MapData(map, X, Y).UserIndex > 0 Then
            TempCharIndex = MapData(map, X, Y).UserIndex
            FoundChar = 1
        End If
        If MapData(map, X, Y).NpcIndex > 0 Then
            TempCharIndex = MapData(map, X, Y).NpcIndex
            FoundChar = 2
        End If
    End If
    
    'Reaccion al personaje
    If FoundChar = 1 Then '  ¿Encontro un Usuario?
            
       If UserList(TempCharIndex).flags.AdminInvisible = 0 Or UserList(UserIndex).flags.Privilegios And PlayerType.Dios Then
            
            If LenB(UserList(TempCharIndex).DescRM) = 0 And UserList(TempCharIndex).showName Then 'No tiene descRM y quiere que se vea su nombre.
                Stat = Stat & "(" & ListaClases(UserList(TempCharIndex).Clase) & " " & ListaRazas(UserList(UserIndex).Raza) & " Nivel "
                
                If UserList(UserIndex).Stats.ELV + 10 < UserList(TempCharIndex).Stats.ELV Then
                    Stat = Stat & "?? "
                Else
                    Stat = Stat & UserList(TempCharIndex).Stats.ELV & " "
                End If
                
                If UserList(TempCharIndex).Stats.MinHP < (UserList(TempCharIndex).Stats.MaxHP * 0.05) Then
                    Stat = Stat & "| Muerto"
                ElseIf UserList(TempCharIndex).Stats.MinHP < (UserList(TempCharIndex).Stats.MaxHP * 0.1) Then
                    Stat = Stat & "| Casi muerto"
                ElseIf UserList(TempCharIndex).Stats.MinHP < (UserList(TempCharIndex).Stats.MaxHP * 0.5) Then
                    Stat = Stat & "| Malherido"
                ElseIf UserList(TempCharIndex).Stats.MinHP < (UserList(TempCharIndex).Stats.MaxHP * 0.75) Then
                    Stat = Stat & "| Herido"
                ElseIf UserList(TempCharIndex).Stats.MinHP < (UserList(TempCharIndex).Stats.MaxHP) Then
                    Stat = Stat & "| Levemente Herido"
                Else
                    Stat = Stat & "| Intacto"
                End If
                
                If UserList(TempCharIndex).flags.Comerciando = True Then
                    Stat = Stat & " | Comerciando)"
                Else
                    Stat = Stat & ")"
                End If
                
                If UserList(TempCharIndex).Faccion.ArmadaReal = 1 Then
                    Stat = Stat & " <Sagrada Orden> " & "<" & TituloReal(TempCharIndex) & ">"
                ElseIf UserList(TempCharIndex).Faccion.FuerzasCaos = 1 Then
                    Stat = Stat & " <Fuerzas del caos> " & "<" & TituloCaos(TempCharIndex) & ">"
                ElseIf UserList(TempCharIndex).Faccion.Milicia = 1 Then
                    Stat = Stat & " <Milicia Republicana> " & "<" & TituloMilicia(TempCharIndex) & ">"
                End If
                
                If UserList(TempCharIndex).GuildIndex > 0 Then
                    Stat = Stat & " <" & modGuilds.GuildName(UserList(TempCharIndex).GuildIndex) & ">"
                End If
                
                If Len(UserList(TempCharIndex).desc) > 0 Then
                    Stat = UserList(TempCharIndex).name & " " & Stat & " - " & UserList(TempCharIndex).desc
                Else
                    Stat = UserList(TempCharIndex).name & " " & Stat
                End If
                
                                
                If Not UserList(TempCharIndex).flags.Privilegios And PlayerType.User Then
                    Stat = Stat & " <Dungeon Master>"
                        
                    ' Elijo el color segun el rango del GM
                    If UserList(TempCharIndex).flags.Privilegios = PlayerType.Dios Then
                        Stat = Stat & "~232~225~0~1~0"
                    ElseIf UserList(TempCharIndex).flags.Privilegios = PlayerType.SemiDios Then
                        ft = FontTypeNames.FONTTYPE_GM
                    ElseIf UserList(TempCharIndex).flags.Privilegios = PlayerType.Consejero Then
                        ft = FontTypeNames.FONTTYPE_CONSE
                    End If
                        
                ElseIf UserList(TempCharIndex).Faccion.Ciudadano = 1 And Not UserList(TempCharIndex).Faccion.ArmadaReal = 1 Then
                    Stat = Stat & " <Imperial> ~0~0~255~1~0"
                    ft = FontTypeNames.FONTTYPE_CITIZEN
                ElseIf UserList(TempCharIndex).Faccion.Renegado = 1 And Not UserList(TempCharIndex).Faccion.FuerzasCaos = 1 Then
                    Stat = Stat & " <Renegado> ~128~128~128~1~0"
                ElseIf UserList(TempCharIndex).Faccion.Republicano = 1 And Not UserList(TempCharIndex).Faccion.Milicia = 1 Then
                    Stat = Stat & " <Republicano> ~255~128~0~1~0"
                    ft = FontTypeNames.FONTTYPE_CITIZEN
                ElseIf UserList(TempCharIndex).Faccion.Milicia = 1 Then
                    Stat = Stat & "~255~128~0~1~0"
                ElseIf UserList(TempCharIndex).Faccion.FuerzasCaos = 1 Then
                    Stat = Stat & "~190~0~0~1~0"
                ElseIf UserList(TempCharIndex).Faccion.ArmadaReal = 1 Then
                    Stat = Stat & "~0~190~200~1~0"
                End If
            Else  'Si tiene descRM la muestro siempre.
                Stat = UserList(TempCharIndex).DescRM
                ft = FontTypeNames.FONTTYPE_INFOBOLD
            End If
            
            If LenB(Stat) > 0 Then
                Call WriteConsoleMsg(1, UserIndex, Stat, ft)
            End If
            
            FoundSomething = 1
            UserList(UserIndex).flags.TargetUser = TempCharIndex
            UserList(UserIndex).flags.TargetNPC = 0
            UserList(UserIndex).flags.TargetNpcTipo = eNPCType.Comun
       End If

    End If
    If FoundChar = 2 Then '¿Encontro un NPC?
            Dim estatus As String
            
            If UserList(UserIndex).flags.Privilegios And (PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin) Then
                estatus = "(" & Npclist(TempCharIndex).Stats.MinHP & "/" & Npclist(TempCharIndex).Stats.MaxHP & ") "
            Else
                If UserList(UserIndex).flags.Muerto = 0 Then
                    If UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) >= 0 And UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) <= 10 Then
                        estatus = "(Dudoso) "
                    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) > 10 And UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) <= 10 Then
                        If Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP / 2) Then
                            estatus = "(Herido) "
                        Else
                            estatus = "(Sano) "
                        End If
                    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) > 20 And UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) <= 10 Then
                        If Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.5) Then
                            estatus = "(Herido) "
                        ElseIf Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.75) Then
                            estatus = "(Levemente Herido) "
                        Else
                            estatus = "(Intacto) "
                        End If
                    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) > 30 And UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) <= 20 Then
                        If Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.25) Then
                            estatus = "(Muy malherido) "
                        ElseIf Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.5) Then
                            estatus = "(Herido) "
                        ElseIf Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.75) Then
                            estatus = "(Levemente herido) "
                        Else
                            estatus = "(Sano) "
                        End If
                    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) > 40 And UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) < 40 Then
                        If Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.05) Then
                            estatus = "(Agonizando) "
                        ElseIf Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.1) Then
                            estatus = "(Casi muerto) "
                        ElseIf Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.25) Then
                            estatus = "(Muy Malherido) "
                        ElseIf Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.5) Then
                            estatus = "(Herido) "
                        ElseIf Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP * 0.75) Then
                            estatus = "(Levemente herido) "
                        ElseIf Npclist(TempCharIndex).Stats.MinHP < (Npclist(TempCharIndex).Stats.MaxHP) Then
                            estatus = "(Sano) "
                        Else
                            estatus = "(Intacto) "
                        End If
                    ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) >= 0 Then
                        estatus = "(" & Npclist(TempCharIndex).Stats.MinHP & "/" & Npclist(TempCharIndex).Stats.MaxHP & ") "
                    Else
                        estatus = "!error!"
                    End If
                End If
            End If
            
            If Len(Npclist(TempCharIndex).desc) > 1 Then
                Call WriteChatOverHead(UserIndex, Npclist(TempCharIndex).desc, Npclist(TempCharIndex).Char.CharIndex, vbWhite)
            ElseIf TempCharIndex = CentinelaNPCIndex Then
                'Enviamos nuevamente el texto del centinela según quien pregunta
                Call modCentinela.CentinelaSendClave(UserIndex)
            Else
                If Npclist(TempCharIndex).MaestroUser > 0 Then
                    Call WriteConsoleMsg(1, UserIndex, estatus & Npclist(TempCharIndex).name & " es mascota de " & UserList(Npclist(TempCharIndex).MaestroUser).name, FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(1, UserIndex, estatus & Npclist(TempCharIndex).name & ".", FontTypeNames.FONTTYPE_INFO)
                    If UserList(UserIndex).flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
                        Call WriteConsoleMsg(1, UserIndex, "Le pegó primero: " & Npclist(TempCharIndex).flags.AttackedFirstBy & ".", FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
                
            End If
            FoundSomething = 1
            UserList(UserIndex).flags.TargetNpcTipo = Npclist(TempCharIndex).NPCtype
            UserList(UserIndex).flags.TargetNPC = TempCharIndex
            UserList(UserIndex).flags.TargetUser = 0
            UserList(UserIndex).flags.TargetObj = 0
        'Else
        '    Dim Bot As Integer
        '    Bot = Npclist(TempCharIndex).Bot
        '    With Bots(Bot)
        '        estatus = "Vida: " & .MinHP & "/" & .MaxHP & " "
        '        estatus = estatus & "Mana: " & .MinMAN & "/" & .MaxMAN
        '    End With
        '    Call WriteConsoleMsg(1,UserIndex, estatus, FontTypeNames.FONTTYPE_INFO)
        'End If
    End If
    
    If FoundChar = 0 Then
        UserList(UserIndex).flags.TargetNPC = 0
        UserList(UserIndex).flags.TargetNpcTipo = eNPCType.Comun
        UserList(UserIndex).flags.TargetUser = 0
    End If
    
    '*** NO ENCOTRO NADA ***
    If FoundSomething = 0 Then
        UserList(UserIndex).flags.TargetNPC = 0
        UserList(UserIndex).flags.TargetNpcTipo = eNPCType.Comun
        UserList(UserIndex).flags.TargetUser = 0
        UserList(UserIndex).flags.TargetObj = 0
        UserList(UserIndex).flags.TargetObjMap = 0
        UserList(UserIndex).flags.TargetObjX = 0
        UserList(UserIndex).flags.TargetObjY = 0
    End If

Else
    If FoundSomething = 0 Then
        UserList(UserIndex).flags.TargetNPC = 0
        UserList(UserIndex).flags.TargetNpcTipo = eNPCType.Comun
        UserList(UserIndex).flags.TargetUser = 0
        UserList(UserIndex).flags.TargetObj = 0
        UserList(UserIndex).flags.TargetObjMap = 0
        UserList(UserIndex).flags.TargetObjX = 0
        UserList(UserIndex).flags.TargetObjY = 0
    End If
End If


End Sub

Function FindDirection(ByVal NPCI As Integer, Target As WorldPos) As eHeading
'*****************************************************************
'Devuelve la direccion en la cual el target se encuentra
'desde pos, 0 si la direc es igual
'*****************************************************************
Dim X As Integer
Dim Y As Integer
Dim Pos As WorldPos
Dim puedeX As Boolean
Dim puedeY As Boolean

Pos = Npclist(NPCI).Pos
X = Npclist(NPCI).Pos.X - Target.X
Y = Npclist(NPCI).Pos.Y - Target.Y

'misma
If Sgn(X) = 0 And Sgn(Y) = 0 Then
    FindDirection = 0
    Exit Function
End If

'Lo tenemos al lado
If Distancia(Pos, Target) = 1 Then
    FindDirection = 0
    Exit Function
End If

If Rodeado(Target) Then
    FindDirection = 0
    Exit Function
End If

'Sur
If Sgn(X) = 0 And Sgn(Y) = -1 Then
    If Not PuedeNpc(Pos.map, Pos.X, Pos.Y + 1) Then
        If RandomNumber(1, 10) > 5 Then
            If PuedeNpc(Pos.map, Pos.X - 1, Pos.Y) Then
                FindDirection = eHeading.WEST: Exit Function
            Else
                FindDirection = eHeading.EAST: Exit Function
            End If
        Else
            If PuedeNpc(Pos.map, Pos.X + 1, Pos.Y) Then
                FindDirection = eHeading.EAST: Exit Function
            Else
                FindDirection = eHeading.WEST: Exit Function
            End If
        End If
    Else
        FindDirection = eHeading.SOUTH: Exit Function
    End If
End If

'norte
If Sgn(X) = 0 And Sgn(Y) = 1 Then
    If Not PuedeNpc(Pos.map, Pos.X, Pos.Y - 1) Then
        If RandomNumber(1, 10) > 5 Then
            If PuedeNpc(Pos.map, Pos.X - 1, Pos.Y) Then
                FindDirection = eHeading.WEST: Exit Function
            Else
                FindDirection = eHeading.EAST: Exit Function
            End If
        Else
            If PuedeNpc(Pos.map, Pos.X + 1, Pos.Y) Then
                FindDirection = eHeading.EAST: Exit Function
            Else
                FindDirection = eHeading.WEST: Exit Function
            End If
        End If
    Else
        FindDirection = eHeading.NORTH: Exit Function
    End If
End If

'oeste
If Sgn(X) = 1 And Sgn(Y) = 0 Then
    If Not PuedeNpc(Pos.map, Pos.X - 1, Pos.Y) Then
        If RandomNumber(1, 10) > 5 Then
            If PuedeNpc(Pos.map, Pos.X, Pos.Y - 1) Then
                FindDirection = eHeading.NORTH: Exit Function
            Else
                FindDirection = eHeading.SOUTH: Exit Function
            End If
        Else
            If PuedeNpc(Pos.map, Pos.X, Pos.Y + 1) Then
                FindDirection = eHeading.SOUTH: Exit Function
            Else
                FindDirection = eHeading.NORTH: Exit Function
            End If
        End If
    Else
        FindDirection = eHeading.WEST: Exit Function
    End If
End If

'este
If Sgn(X) = -1 And Sgn(Y) = 0 Then
    If Not PuedeNpc(Pos.map, Pos.X + 1, Pos.Y) Then
        If RandomNumber(1, 10) > 5 Then
            If PuedeNpc(Pos.map, Pos.X, Pos.Y - 1) Then
                FindDirection = eHeading.NORTH: Exit Function
            Else
                FindDirection = eHeading.SOUTH: Exit Function
            End If
        Else
            If PuedeNpc(Pos.map, Pos.X, Pos.Y + 1) Then
                FindDirection = eHeading.SOUTH: Exit Function
            Else
                FindDirection = eHeading.NORTH: Exit Function
            End If
        End If
    Else
        FindDirection = eHeading.EAST: Exit Function
    End If
End If

'NW
If Sgn(X) = 1 And Sgn(Y) = 1 Then
    puedeX = PuedeNpc(Pos.map, Pos.X - 1, Pos.Y)
    puedeY = PuedeNpc(Pos.map, Pos.X, Pos.Y + 1)
    If puedeX And puedeY Then
        puedeX = Not (Npclist(NPCI).oldPos.X = Pos.X - 1)
        puedeY = Not (Npclist(NPCI).oldPos.Y = Pos.Y - 1)
        If puedeX And puedeY Then
            If RandomNumber(1, 20) < 10 Then
                FindDirection = eHeading.WEST: Exit Function
            Else
                FindDirection = eHeading.NORTH: Exit Function
            End If
        Else
            If puedeX Then
                FindDirection = eHeading.WEST: Exit Function
            ElseIf puedeY Then
                FindDirection = eHeading.NORTH: Exit Function
            End If
        End If
    End If
    
    If Not puedeY Or Npclist(NPCI).oldPos.Y = Pos.Y - 1 Then
        FindDirection = eHeading.WEST: Exit Function
    Else
        FindDirection = eHeading.NORTH: Exit Function
    End If
    
    'llego aca porque no pudo en nada
    puedeX = PuedeNpc(Pos.map, Pos.X - 1, Pos.Y)
    puedeY = PuedeNpc(Pos.map, Pos.X, Pos.Y + 1)
    If Not puedeY Or Npclist(NPCI).oldPos.Y = Pos.Y + 1 Then
        FindDirection = eHeading.EAST: Exit Function
    Else
        FindDirection = eHeading.SOUTH: Exit Function
    End If
End If

'NE
If Sgn(X) = -1 And Sgn(Y) = 1 Then
    puedeX = PuedeNpc(Pos.map, Pos.X + 1, Pos.Y)
    puedeY = PuedeNpc(Pos.map, Pos.X, Pos.Y + 1)
    If puedeX And puedeY Then
        puedeX = Not (Npclist(NPCI).oldPos.X = Pos.X + 1)
        puedeY = Not (Npclist(NPCI).oldPos.Y = Pos.Y - 1)
        If puedeX And puedeY Then
            If RandomNumber(1, 20) < 10 Then
                FindDirection = eHeading.EAST: Exit Function
            Else
                FindDirection = eHeading.NORTH: Exit Function
            End If
        Else
            If puedeX Then
                FindDirection = eHeading.EAST: Exit Function
            ElseIf puedeY Then
                FindDirection = eHeading.NORTH: Exit Function
            End If
        End If
    End If
    
    If Not puedeY Or Npclist(NPCI).oldPos.Y = Pos.Y - 1 Then
        FindDirection = eHeading.EAST: Exit Function
    Else
        FindDirection = eHeading.NORTH: Exit Function
    End If
    
    'llego aca porque no pudo en nada
    puedeX = PuedeNpc(Pos.map, Pos.X - 1, Pos.Y)
    puedeY = PuedeNpc(Pos.map, Pos.X, Pos.Y + 1)
    If Not puedeY Or Npclist(NPCI).oldPos.Y = Pos.Y + 1 Then
        FindDirection = eHeading.WEST: Exit Function
    Else
        FindDirection = eHeading.SOUTH: Exit Function
    End If
End If

'SW
If Sgn(X) = 1 And Sgn(Y) = -1 Then
    puedeX = PuedeNpc(Pos.map, Pos.X - 1, Pos.Y)
    puedeY = PuedeNpc(Pos.map, Pos.X, Pos.Y + 1)
    If puedeX And puedeY Then
        puedeX = Not (Npclist(NPCI).oldPos.X = Pos.X - 1)
        puedeY = Not (Npclist(NPCI).oldPos.Y = Pos.Y + 1)
        If puedeX And puedeY Then
            If RandomNumber(1, 20) < 10 Then
                FindDirection = eHeading.WEST: Exit Function
            Else
                FindDirection = eHeading.SOUTH: Exit Function
            End If
        Else
            If puedeX Then
                FindDirection = eHeading.WEST: Exit Function
            ElseIf puedeY Then
                FindDirection = eHeading.SOUTH: Exit Function
            End If
        End If
    End If
    
    If Not puedeY Or Npclist(NPCI).oldPos.Y = Pos.Y + 1 Then
        FindDirection = eHeading.WEST: Exit Function
    Else
        FindDirection = eHeading.SOUTH: Exit Function
    End If
    
    'llego aca porque no pudo en nada
    puedeX = PuedeNpc(Pos.map, Pos.X + 1, Pos.Y)
    puedeY = PuedeNpc(Pos.map, Pos.X, Pos.Y - 1)
    If Not puedeY Or Npclist(NPCI).oldPos.Y = Pos.Y - 1 Then
        FindDirection = eHeading.EAST: Exit Function
    Else
        FindDirection = eHeading.NORTH: Exit Function
    End If
End If

'SE
If Sgn(X) = -1 And Sgn(Y) = -1 Then
    puedeX = PuedeNpc(Pos.map, Pos.X + 1, Pos.Y)
    puedeY = PuedeNpc(Pos.map, Pos.X, Pos.Y + 1)
    If puedeX And puedeY Then
        puedeX = Not (Npclist(NPCI).oldPos.X = Pos.X + 1)
        puedeY = Not (Npclist(NPCI).oldPos.Y = Pos.Y + 1)
        If puedeX And puedeY Then
            If RandomNumber(1, 20) < 10 Then
                FindDirection = eHeading.EAST: Exit Function
            Else
                FindDirection = eHeading.SOUTH: Exit Function
            End If
        Else
            If puedeX Then
                FindDirection = eHeading.EAST: Exit Function
            ElseIf puedeY Then
                FindDirection = eHeading.SOUTH: Exit Function
            End If
        End If
    End If
    
    If Not puedeY And Npclist(NPCI).oldPos.Y = Pos.Y + 1 Then
        FindDirection = eHeading.EAST: Exit Function
    ElseIf puedeY And Not Npclist(NPCI).oldPos.Y = Pos.Y + 1 Then
        FindDirection = eHeading.SOUTH: Exit Function
    End If
    
    'llego aca porque no pudo en nada
    puedeX = PuedeNpc(Pos.map, Pos.X - 1, Pos.Y)
    puedeY = PuedeNpc(Pos.map, Pos.X, Pos.Y - 1)
    If Not puedeY Or Npclist(NPCI).oldPos.Y = Pos.Y - 1 Then
        FindDirection = eHeading.WEST: Exit Function
    Else
        FindDirection = eHeading.NORTH: Exit Function
    End If
End If

End Function
Function Rodeado(ByRef Pos As WorldPos) As Boolean
    If Not PuedeNpc(Pos.map, Pos.X + 1, Pos.Y) Then
        If Not PuedeNpc(Pos.map, Pos.X + 1, Pos.Y) Then
            If Not PuedeNpc(Pos.map, Pos.X + 1, Pos.Y) Then
                If Not PuedeNpc(Pos.map, Pos.X + 1, Pos.Y) Then
                    Rodeado = True
                End If
            End If
        End If
    End If
End Function
Function PuedeNpc(ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer)
    PuedeNpc = (MapData(map, X, Y).NpcIndex = 0 And _
                MapData(map, X, Y).Blocked = 0 And _
                MapData(map, X, Y).UserIndex = 0)
End Function
'[Barrin 30-11-03]
Public Function ItemNoEsDeMapa(ByVal index As Integer) As Boolean

ItemNoEsDeMapa = ObjData(index).OBJType <> eOBJType.otPuertas And _
            ObjData(index).OBJType <> eOBJType.otForos And _
            ObjData(index).OBJType <> eOBJType.otCarteles And _
            ObjData(index).OBJType <> eOBJType.otArboles And _
            ObjData(index).OBJType <> eOBJType.otYacimiento And _
            ObjData(index).OBJType <> eOBJType.otTeleport
End Function
'[/Barrin 30-11-03]

Public Function MostrarCantidad(ByVal index As Integer) As Boolean
MostrarCantidad = ObjData(index).OBJType <> eOBJType.otPuertas And _
            ObjData(index).OBJType <> eOBJType.otForos And _
            ObjData(index).OBJType <> eOBJType.otCarteles And _
            ObjData(index).OBJType <> eOBJType.otArboles And _
            ObjData(index).OBJType <> eOBJType.otYacimiento And _
            ObjData(index).OBJType <> eOBJType.otTeleport
End Function

Public Function EsObjetoFijo(ByVal OBJType As eOBJType) As Boolean

EsObjetoFijo = OBJType = eOBJType.otForos Or _
               OBJType = eOBJType.otCarteles Or _
               OBJType = eOBJType.otArboles Or _
               OBJType = eOBJType.otYacimiento

End Function
Public Function ParticleToLevel(ByVal UserIndex As Integer) As Integer
If UserList(UserIndex).Stats.ELV < 13 Then
    ParticleToLevel = 42
ElseIf UserList(UserIndex).Stats.ELV < 25 Then
    ParticleToLevel = 81
ElseIf UserList(UserIndex).Stats.ELV < 35 Then
    ParticleToLevel = 41
ElseIf UserList(UserIndex).Stats.ELV < 50 Then
    ParticleToLevel = 39
ElseIf UserList(UserIndex).Stats.ELV = 57 Then
    ParticleToLevel = 36
ElseIf UserList(UserIndex).Stats.ELV = 59 Then
    ParticleToLevel = 26
ElseIf UserList(UserIndex).Stats.ELV = 60 Then
    ParticleToLevel = 107
    If UserList(UserIndex).Faccion.Renegado = 50 Then
        ParticleToLevel = 109
    ElseIf UserList(UserIndex).Faccion.Ciudadano = 50 Then
        ParticleToLevel = 113
    ElseIf UserList(UserIndex).Faccion.Republicano = 50 Then
        ParticleToLevel = 110
    ElseIf UserList(UserIndex).Faccion.ArmadaReal = 50 Then
        ParticleToLevel = 112
    ElseIf UserList(UserIndex).Faccion.FuerzasCaos = 50 Then
        ParticleToLevel = 108
    ElseIf UserList(UserIndex).Faccion.Milicia = 50 Then
        ParticleToLevel = 110
    End If
Else
    ParticleToLevel = 107
End If
End Function
Public Function Tilde(data As String) As String
 
Tilde = Replace(Replace(Replace(Replace(Replace(UCase$(data), "Á", "A"), "É", "E"), "Í", "I"), "Ó", "O"), "Ú", "U")
 
End Function
