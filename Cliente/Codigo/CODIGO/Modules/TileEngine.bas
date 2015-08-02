Attribute VB_Name = "Mod_TileEngine"
Option Explicit
'[CODE 001]:MatuX
Public Enum PlayLoop
    plnone = 0
    plLluviain = 1
    plLluviaout = 2
End Enum

Public map_base_light As Long

'Map sizes in tiles
Public Const XMinMapSize As Byte = 1
Public Const XMaxMapSize As Byte = 100
Public Const YMinMapSize As Byte = 1
Public Const YMaxMapSize As Byte = 100

Public CurMap As Byte
Public grhCount As Long
Public AmbientColor As D3DCOLORVALUE

'Sets a Grh animation to loop indefinitely.
Private Const INFINITE_LOOPS As Integer = -1

'Posicion en un mapa
Public Type Position
    X As Long
    Y As Long
End Type

'Posicion en el Mundo
Public Type WorldPos
    map As Integer
    X As Integer
    Y As Integer
End Type


Type GrhData
    sX As Integer
    sY As Integer
    FileNum As Integer
    pixelWidth As Integer
    pixelHeight As Integer
    TileWidth As Single
    TileHeight As Single
    NumFrames As Integer
    Frames() As Integer
    speed As Single 'Integer
    mini_map_color As Long
End Type


'apunta a una estructura grhdata y mantiene la animacion
Public Type Grh
    grhindex As Integer
    FrameCounter As Single
    speed As Single
    Started As Byte
    Loops As Integer
    angle As Single
End Type

'Lista de cuerpos
Public Type BodyData
    Walk(E_Heading.NORTH To E_Heading.WEST) As Grh
    HeadOffset As Position
End Type

'Lista de cabezas
Public Type HeadData
    Head(E_Heading.NORTH To E_Heading.WEST) As Grh
End Type

'Lista de las animaciones de las armas
Type WeaponAnimData
    WeaponWalk(E_Heading.NORTH To E_Heading.WEST) As Grh
End Type

'Lista de las animaciones de los escudos
Type ShieldAnimData
    ShieldWalk(E_Heading.NORTH To E_Heading.WEST) As Grh
End Type

Type tAura
    Grh As Grh
    color As Long
End Type

'Apariencia del personaje
Public Type Char
    active As Byte
    heading As E_Heading
    Pos As Position
    
    label_color(3) As Long
    
    
    iBody As Integer
    body As BodyData
    
    iHead As Integer
    Head As HeadData
    
    Casco As HeadData
    
    Arma As WeaponAnimData
    UsandoArma As Boolean
    
    Escudo As ShieldAnimData
    ShieldOffSetY As Integer
    
    plusGrh(2) As tAura
    
    fX As Grh
    fxIndex As Integer
    
    AlphaX As Integer
    last_tick As Long
    
    Criminal As Byte
    
    nombre As String
    
    
    group_index As Integer
        
    particle_count As Integer
    particle_group() As Long
    
    scrollDirectionX As Integer
    scrollDirectionY As Integer
    
    Moving As Byte
    MoveOffsetX As Single
    MoveOffsetY As Single
    
    pie As Boolean
    Muerto As Boolean
    invisible As Boolean
    Priv As Byte
End Type

'Info de un objeto
Public Type Obj
    TieneLuz As Byte
    OBJIndex As Integer
    Amount As Integer
    name As String
End Type

'Tipo de las celdas del mapa
Public Type MapBlock
    Graphic(1 To 4) As Grh
    CharIndex As Integer
    ObjGrh As Grh
    
    light_value(0 To 3) As Long
    
    luz As Integer
    color(3) As Long
    
    particle_group_index As Integer
    effectIndex As Integer
    
    NPCIndex As Integer
    OBJInfo As Obj
    TileExit As WorldPos
    Blocked As Byte
    
    Trigger As Integer
End Type

'Info de cada mapa
Public Type MapInfo
    Music As String
    name As String
    StartPos As WorldPos
    MapVersion As Integer
End Type

Public IniPath As String
Public MapPath As String

'Status del user
Public UserI As Integer
Public UserMoving As Byte
Public UserBody As Integer
Public UserHead As Integer
Public UserPos As Position 'Posicion
Public AddtoUserPos As Position 'Si se mueve
Public UserCharIndex As Integer

Public FPS As Long

Public TilePixelHeight As Integer
Public TilePixelWidth As Integer

Dim timerElapsedTime As Single
Dim timerTicksPerFrame As Single
Public engineBaseSpeed As Single

Public NumBodies As Integer
Public Numheads As Integer
Public NumFxs As Integer

Public NumChars As Integer
Public LastChar As Integer
Public NumWeaponAnims As Integer
Public NumShieldAnims As Integer

Public GrhData() As GrhData
Public BodyData() As BodyData
Public HeadData() As HeadData
Public FxData() As tIndiceFx
Public WeaponAnimData() As WeaponAnimData
Public ShieldAnimData() As ShieldAnimData
Public CascoAnimData() As HeadData


Public MapConnect() As MapBlock
Public MapAccount() As MapBlock
Public MapData() As MapBlock
Public MapInfo As MapInfo

Public bRain        As Boolean 'está raineando?
Public bTecho       As Boolean 'hay techo?
Public brstTick     As Long

Public charlist(1 To 10000) As Char

Private Type size
    cx As Long
    cy As Long
End Type



Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function StretchBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As size) As Long

Sub ConvertCPtoTP(ByVal viewPortX As Integer, ByVal viewPortY As Integer, ByRef tX As Byte, ByRef tY As Byte)
'******************************************
'Converts where the mouse is in the main window to a tile position. MUST be called eveytime the mouse moves.
'******************************************
    tX = UserPos.X + viewPortX \ 32 - frmMain.MainViewPic.ScaleWidth \ 64
    tY = UserPos.Y + viewPortY \ 32 - frmMain.MainViewPic.ScaleHeight \ 64
    Debug.Print tX; tY
End Sub

Sub ResetCharInfo(ByVal CharIndex As Integer)
    With charlist(CharIndex)
        Call Engine.Char_Particle_Group_Remove_All(CharIndex)
        .active = 0
        .Criminal = 0
        .fxIndex = 0
        .invisible = False
        .Moving = 0
        .Muerto = False
        .nombre = ""
        .pie = False
        .Pos.X = 0
        .Pos.Y = 0
        .UsandoArma = False
    End With
End Sub
Sub MakeChar(ByVal CharIndex As Integer, ByVal body As Integer, ByVal Head As Integer, ByVal heading As Byte, ByVal X As Integer, ByVal Y As Integer, ByVal Arma As Integer, ByVal Escudo As Integer, ByVal Casco As Integer)
On Error Resume Next
    'Apuntamos al ultimo Char
    If CharIndex > LastChar Then LastChar = CharIndex
    
    With charlist(CharIndex)
        'If the char wasn't allready active (we are rewritting it) don't increase char count
        If .active = 0 Then _
            NumChars = NumChars + 1
        
        If Arma = 0 Then Arma = 2
        If Escudo = 0 Then Escudo = 2
        If Casco = 0 Then Casco = 2
        
        .iHead = Head
        .iBody = body
        .Head = HeadData(Head)
        .body = BodyData(body)
        
        If Not Arma = 29 Then .Arma = WeaponAnimData(Arma)
        
        .Escudo = ShieldAnimData(Escudo)
        .Casco = CascoAnimData(Casco)
        
        .heading = heading
        
        'Reset moving stats
        .Moving = 0
        .MoveOffsetX = 0
        .MoveOffsetY = 0
        
        'Update position
        .Pos.X = X
        .Pos.Y = Y
        
        'Make active
        .active = 1
        
        Select Case .Priv
            Case 1 'Gris
                Engine.Long_To_RGB_List .label_color, D3DColorXRGB(128, 128, 128)
            Case 2 'Azul
                Engine.Long_To_RGB_List .label_color, D3DColorXRGB(0, 0, 230)
            Case 3 'Rojo
                Engine.Long_To_RGB_List .label_color, D3DColorXRGB(190, 0, 0)
            Case 4 'Naranja
                Engine.Long_To_RGB_List .label_color, D3DColorXRGB(255, 128, 0)
            Case 5 'Verde
                Engine.Long_To_RGB_List .label_color, D3DColorXRGB(232, 225, 0)
            Case 6 'Azul Armada real
                Engine.Long_To_RGB_List .label_color, D3DColorXRGB(0, 190, 180)
        End Select
    End With
    
    Call PonerAura(CharIndex, Escudo, Arma, body)
    
    'Plot on map
    MapData(X, Y).CharIndex = CharIndex
End Sub
Sub PonerAura(ByVal CharIndex As Integer, ByVal Escudo As Byte, ByVal Arma As Byte, ByVal body As Integer)
With charlist(CharIndex)
    If body = 255 Then
        InitGrh .plusGrh(2).Grh, 20206
        .plusGrh(2).color = &HFFFD7E
    Else
        .plusGrh(2).Grh.grhindex = 0
    End If

    If Escudo = 14 Then
        InitGrh .plusGrh(1).Grh, 20203
        .plusGrh(1).color = &HFFCC33
    Else
        .plusGrh(1).Grh.grhindex = 0
    End If

    If Arma = 23 Then
        InitGrh .plusGrh(0).Grh, 20128
        .plusGrh(0).color = &HFFCC33
    ElseIf Arma = 24 Then
        InitGrh .plusGrh(0).Grh, 20133
        .plusGrh(0).color = &HFF3300
    ElseIf Arma = 25 Then
        InitGrh .plusGrh(0).Grh, 20152
        .plusGrh(0).color = &HFF0000
    ElseIf Arma = 26 Then
        InitGrh .plusGrh(0).Grh, 20185
        .plusGrh(0).color = -65536
    ElseIf Arma = 31 Then
        InitGrh .plusGrh(0).Grh, 20155
        .plusGrh(0).color = &HFF0000
    ElseIf Arma = 27 Then
        InitGrh .plusGrh(0).Grh, 20151
        .plusGrh(0).color = &HFFFF00
    ElseIf Arma = 28 Then
        InitGrh .plusGrh(0).Grh, 20147
        .plusGrh(0).color = &HFF
    ElseIf Arma = 29 Then
        InitGrh .plusGrh(0).Grh, 20146
        .plusGrh(0).color = &H6B1B
    ElseIf Arma = 30 Then
        InitGrh .plusGrh(0).Grh, 20200
        .plusGrh(0).color = &HCCFF33
    ElseIf Arma = 32 Then
        InitGrh .plusGrh(0).Grh, 20147
        .plusGrh(0).color = &HFF
    ElseIf Arma = 33 Then
        InitGrh .plusGrh(0).Grh, 20058
        .plusGrh(0).color = &HFF
     ElseIf Arma = 34 Then
        InitGrh .plusGrh(0).Grh, 20108
        .plusGrh(0).color = &HFF
     ElseIf Arma = 35 Then
        InitGrh .plusGrh(0).Grh, 20109
        .plusGrh(0).color = &HFF
     ElseIf Arma = 36 Then
        InitGrh .plusGrh(0).Grh, 20109
        .plusGrh(0).color = &HFF
     ElseIf Arma = 37 Then
        InitGrh .plusGrh(0).Grh, 20153
        .plusGrh(0).color = &HFF
    ElseIf .MoveOffsetY = 35 Then
        InitGrh .plusGrh(0).Grh, 20147
        .plusGrh(0).color = &HCCFF99
    Else
        .plusGrh(0).Grh.grhindex = 0
    End If
    
    If body = 291 Then
        .ShieldOffSetY = 30
    ElseIf body = 415 Or body = 384 Or body = 382 Then
        .ShieldOffSetY = 16
    ElseIf body = 416 Then
        .ShieldOffSetY = 32
    ElseIf body = 282 Or body = 292 Then
        .ShieldOffSetY = 20
    ElseIf body = 381 Or body = 383 Then
        .ShieldOffSetY = 24
    ElseIf body = 317 Or body = 292 Then
        .ShieldOffSetY = 20
    Else
        .ShieldOffSetY = 0
    End If
    
    If BodyData(body).HeadOffset.Y = -28 Then
        .ShieldOffSetY = .ShieldOffSetY - 5
    End If
    
End With
End Sub

Sub EraseChar(ByVal CharIndex As Integer)
'*****************************************************************
'Erases a character from CharList and map
'*****************************************************************
On Error Resume Next
    charlist(CharIndex).active = 0
    
    'Update lastchar
    If CharIndex = LastChar Then
        Do Until charlist(LastChar).active = 1
            LastChar = LastChar - 1
            If LastChar = 0 Then Exit Do
        Loop
    End If
    
    If Not charlist(CharIndex).Pos.Y = 0 And Not charlist(CharIndex).Pos.X = 0 Then MapData(charlist(CharIndex).Pos.X, charlist(CharIndex).Pos.Y).CharIndex = 0
    
    'Remove char's dialog
    Call Dialogos.RemoveDialog(CharIndex)
    
    Call ResetCharInfo(CharIndex)
    
    'Update NumChars
    NumChars = NumChars - 1
End Sub

Public Sub InitGrh(ByRef Grh As Grh, ByVal grhindex As Integer, Optional ByVal Started As Byte = 2)
'*****************************************************************
'Sets up a grh. MUST be done before rendering
'*****************************************************************
    Grh.grhindex = grhindex
    
    If Started = 2 Then
        If GrhData(Grh.grhindex).NumFrames > 1 Then
            Grh.Started = 1
        Else
            Grh.Started = 0
        End If
    Else
        'Make sure the graphic can be started
        If GrhData(Grh.grhindex).NumFrames = 1 Then Started = 0
        Grh.Started = Started
    End If
    
    
    If Grh.Started Then
        Grh.Loops = INFINITE_LOOPS
    Else
        Grh.Loops = 0
    End If
    
    Grh.FrameCounter = 1
    
    If GrhData(Grh.grhindex).NumFrames > 1 Then
        Grh.speed = 0.4
    End If
    
    'Grh.speed = GrhData(Grh.grhindex).NumFrames / 0.018  'GrhData(Grh.grhindex).speed
End Sub

Public Sub DoFogataFx()
    Dim location As Position
    
    If bFogata Then
        bFogata = HayFogata(location)
        If Not bFogata Then
            Call Audio.StopWave(FogataBufferIndex)
            FogataBufferIndex = 0
        End If
    Else
        bFogata = HayFogata(location)
        If bFogata And FogataBufferIndex = 0 Then FogataBufferIndex = Audio.PlayWave("fuego.wav", location.X, location.Y, LoopStyle.Enabled)
    End If
End Sub

Private Function EstaPCarea(ByVal CharIndex As Integer) As Boolean
    With charlist(CharIndex).Pos
        EstaPCarea = .X > UserPos.X - MinXBorder And .X < UserPos.X + MinXBorder And .Y > UserPos.Y - MinYBorder And .Y < UserPos.Y + MinYBorder
    End With
End Function

Sub DoPasosFx(ByVal CharIndex As Integer)
    If Not UserNavegando Then
        With charlist(CharIndex)
            If .Muerto = False And EstaPCarea(CharIndex) = True And (.Priv = 0 Or .Priv > 5) Then
                .pie = Not .pie
                
                If .pie Then
                    Call Audio.PlayWave(SND_PASOS1, .Pos.X, .Pos.Y)
                Else
                    Call Audio.PlayWave(SND_PASOS2, .Pos.X, .Pos.Y)
                End If
            End If
        End With
    ElseIf UserMontando = True Then
        
    ElseIf UserNavegando = True Then
' TODO : Actually we would have to check if the CharIndex char is in the water or not....
        Call Audio.PlayWave(SND_NAVEGANDO, charlist(CharIndex).Pos.X, charlist(CharIndex).Pos.Y)
    End If
End Sub

Sub MoveCharbyPos(ByVal CharIndex As Integer, ByVal nX As Integer, ByVal nY As Integer)
On Error Resume Next
    Dim X As Integer
    Dim Y As Integer
    Dim addx As Integer
    Dim addy As Integer
    Dim nHeading As E_Heading
    
    With charlist(CharIndex)
        X = .Pos.X
        Y = .Pos.Y
        
        MapData(X, Y).CharIndex = 0
        
        addx = nX - X
        addy = nY - Y
        
        If Sgn(addx) = 1 Then
            nHeading = E_Heading.EAST
        End If
        
        If Sgn(addx) = -1 Then
            nHeading = E_Heading.WEST
        End If
        
        If Sgn(addy) = -1 Then
            nHeading = E_Heading.NORTH
        End If
        
        If Sgn(addy) = 1 Then
            nHeading = E_Heading.SOUTH
        End If
        
        MapData(nX, nY).CharIndex = CharIndex
        
        
        .Pos.X = nX
        .Pos.Y = nY
        
        .MoveOffsetX = -1 * (TilePixelWidth * addx)
        .MoveOffsetY = -1 * (TilePixelHeight * addy)
        
        .Moving = 1
        .heading = nHeading
        
        .scrollDirectionX = Sgn(addx)
        .scrollDirectionY = Sgn(addy)
        
        'parche para que no medite cuando camina
        If .fxIndex = FxMeditar.CHICO Or .fxIndex = FxMeditar.GRANDE Or .fxIndex = FxMeditar.MEDIANO Or .fxIndex = FxMeditar.XGRANDE Or .fxIndex = FxMeditar.XXGRANDE Then
            .fxIndex = 0
        End If
    End With
    
    If Not EstaPCarea(CharIndex) Then Call Dialogos.RemoveDialog(CharIndex)
    
    If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
        Call EraseChar(CharIndex)
    End If
End Sub
Function InMapBounds(ByVal X As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is in the maps bounds
'*****************************************************************
    If X < XMinMapSize Or X > XMaxMapSize Or Y < YMinMapSize Or Y > YMaxMapSize Then
        Exit Function
    End If
    
    InMapBounds = True
End Function
Sub MoveScreen(ByVal nHeading As E_Heading)
'******************************************
'Starts the screen moving in a direction
'******************************************
    Dim X As Integer
    Dim Y As Integer
    Dim tX As Integer
    Dim tY As Integer
    
    'Figure out which way to move
    Select Case nHeading
        Case E_Heading.NORTH
            Y = -1
        
        Case E_Heading.EAST
            X = 1
        
        Case E_Heading.SOUTH
            Y = 1
        
        Case E_Heading.WEST
            X = -1
    End Select
    
    'Fill temp pos
    tX = UserPos.X + X
    tY = UserPos.Y + Y
    
    'Check to see if its out of bounds
    If tX < XMinMapSize Or tX > XMaxMapSize Or tY < YMinMapSize Or tY > YMaxMapSize Then
        Exit Sub
    Else
        'Start moving... MainLoop does the rest
        AddtoUserPos.X = X
        UserPos.X = tX
        AddtoUserPos.Y = Y
        UserPos.Y = tY
        UserMoving = 1
        
        bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
                MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
                MapData(UserPos.X, UserPos.Y).Trigger = 4 Or _
                MapData(UserPos.X, UserPos.Y).Trigger >= 20, True, False)
    End If
    Call DibujarMiniMapPos
End Sub

Private Function HayFogata(ByRef location As Position) As Boolean
    Dim j As Long
    Dim k As Long
    
    For j = UserPos.X - 8 To UserPos.X + 8
        For k = UserPos.Y - 6 To UserPos.Y + 6
            If InMapBounds(j, k) Then
                If MapData(j, k).ObjGrh.grhindex = 1521 Then
                    location.X = j
                    location.Y = k
                    
                    HayFogata = True
                    Exit Function
                End If
            End If
        Next k
    Next j
End Function

Function NextOpenChar() As Integer
'*****************************************************************
'Finds next open char slot in CharList
'*****************************************************************
    Dim loopc As Long
    Dim Dale As Boolean
    
    loopc = 1
    Do While charlist(loopc).active And Dale
        loopc = loopc + 1
        Dale = (loopc <= UBound(charlist))
    Loop
    
    NextOpenChar = loopc
End Function

Function LegalPos(ByVal X As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is legal
'*****************************************************************
    'Limites del mapa
    If X < XMinMapSize Or X > XMaxMapSize Or Y < YMinMapSize Or Y > YMaxMapSize Then
        Exit Function
    End If
    
    'Tile Bloqueado?
    If MapData(X, Y).Blocked = 1 Then
        Exit Function
    End If
    
    '¿Hay un personaje?
    If MapData(X, Y).CharIndex > 0 Then
        Exit Function
    End If
   
    If UserNavegando <> HayAgua(X, Y) Then
        Exit Function
    End If
    
    If UserMontando = True Then
        If MapData(X, Y).Trigger = 1 Or MapData(X, Y).Trigger = 2 Or MapData(X, Y).Trigger = 4 Or MapData(X, Y).Trigger >= 20 Then
            Exit Function
        End If
    End If
    
    LegalPos = True
End Function
Public Function RenderSounds()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/30/2008
'Actualiza todos los sonidos del mapa.
'**************************************************************
    If bRain Then
        If bTecho Then
            If frmMain.IsPlaying <> PlayLoop.plLluviain Then
                If RainBufferIndex Then _
                    Call Audio.StopWave(RainBufferIndex)
                RainBufferIndex = Audio.PlayWave("lluviain.wav", 0, 0, LoopStyle.Enabled)
                frmMain.IsPlaying = PlayLoop.plLluviain
            End If
        Else
            If frmMain.IsPlaying <> PlayLoop.plLluviaout Then
                If RainBufferIndex Then _
                    Call Audio.StopWave(RainBufferIndex)
                RainBufferIndex = Audio.PlayWave("lluviaout.wav", 0, 0, LoopStyle.Enabled)
                frmMain.IsPlaying = PlayLoop.plLluviaout
            End If
        End If
    End If
    
    DoFogataFx
End Function

Function HayUserAbajo(ByVal X As Integer, ByVal Y As Integer, ByVal grhindex As Integer) As Boolean
    If grhindex > 0 Then
        HayUserAbajo = _
            charlist(UserCharIndex).Pos.X >= X - (GrhData(grhindex).TileWidth \ 2) _
                And charlist(UserCharIndex).Pos.X <= X + (GrhData(grhindex).TileWidth \ 2) _
                And charlist(UserCharIndex).Pos.Y >= Y - (GrhData(grhindex).TileHeight - 1) _
                And charlist(UserCharIndex).Pos.Y <= Y
    End If
End Function


Function HayAgua(ByVal X As Integer, ByVal Y As Integer) As Boolean
    HayAgua = ((MapData(X, Y).Graphic(1).grhindex >= 1505 And MapData(X, Y).Graphic(1).grhindex <= 1520) Or _
            (MapData(X, Y).Graphic(1).grhindex >= 5665 And MapData(X, Y).Graphic(1).grhindex <= 5680) Or _
            (MapData(X, Y).Graphic(1).grhindex >= 13547 And MapData(X, Y).Graphic(1).grhindex <= 13562)) And _
                MapData(X, Y).Graphic(2).grhindex = 0
                
End Function
Public Sub RefreshAllChars()
'*****************************************************************
'Goes through the charlist and replots all the characters on the map
'Used to make sure everyone is visible
'*****************************************************************
    Dim loopc As Long
    
    For loopc = 1 To LastChar
        If charlist(loopc).active = 1 Then
            MapData(charlist(loopc).Pos.X, charlist(loopc).Pos.Y).CharIndex = loopc
        End If
    Next loopc
End Sub
Sub MoveTo(ByVal Direccion As E_Heading)
'***************************************************
'Author: Alejandro Santos (AlejoLp)
'Last Modify Date: 06/28/2008
'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
' 06/03/2006: AlejoLp - Elimine las funciones Move[NSWE] y las converti a esta
' 12/08/2007: Tavo    - Si el usuario esta paralizado no se puede mover.
' 06/28/2008: NicoNZ - Saqué lo que impedía que si el usuario estaba paralizado se ejecute el sub.
'***************************************************
    Dim LegalOk As Boolean
    
    Select Case Direccion
        Case E_Heading.NORTH
            LegalOk = LegalPos(UserPos.X, UserPos.Y - 1)
        Case E_Heading.EAST
            LegalOk = LegalPos(UserPos.X + 1, UserPos.Y)
        Case E_Heading.SOUTH
            LegalOk = LegalPos(UserPos.X, UserPos.Y + 1)
        Case E_Heading.WEST
            LegalOk = LegalPos(UserPos.X - 1, UserPos.Y)
    End Select
    
    If LegalOk And Not UserParalizado Then
        Call WriteWalk(Direccion)
        If Not UserDescansar And Not UserMeditar Then
            Engine.Char_Move_By_Head UserCharIndex, Direccion
            MoveScreen Direccion
        End If
    Else
        If charlist(UserCharIndex).heading <> Direccion Then
            Call WriteChangeHeading(Direccion)
        End If
    End If
    
    If frmMain.macrotrabajo.Enabled Then frmMain.DesactivarMacroTrabajo
    
    ' Update 3D sounds!
    Call Audio.MoveListener(UserPos.X, UserPos.Y)
    
End Sub

Sub RandomMove()
'***************************************************
'Author: Alejandro Santos (AlejoLp)
'Last Modify Date: 06/03/2006
' 06/03/2006: AlejoLp - Ahora utiliza la funcion MoveTo
'***************************************************
    Call MoveTo(RandomNumber(NORTH, WEST))
End Sub

Public Function BodyIsBoat(ByVal body As Integer) As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 10/07/2008
'Checks if a given body index is a boat
'**************************************************************
'TODO : This should be checked somehow else. This is nasty....
    If body = iFragataReal Or body = iFragataCaos Or body = iBarcaPk Or _
            body = iGaleraPk Or body = iGaleonPk Or body = iBarcaCiuda Or _
            body = iGaleraCiuda Or body = iGaleonCiuda Or body = iFragataFantasmal Then
        BodyIsBoat = True
    End If
End Function
