Attribute VB_Name = "Mod_General"
Option Explicit
'Set mouse speed
Private Declare Function SystemParametersInfo Lib "user32" Alias _
    "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, _
    ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
 
Private Const SPI_SETMOUSESPEED = 113
Private Const SPI_GETMOUSESPEED = 112
Private Const GWL_EXSTYLE As Long = (-20)
Private Const WS_EX_TRANSPARENT As Long = &H20&
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_ALPHA = &H2&
Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    lParam As Any) As Long
    
'***************************
'Sinuhe - Map format .CSM
'***************************
Private Type tMapHeader
    NumeroBloqueados As Long
    NumeroLayers(2 To 4) As Long
    NumeroTriggers As Long
    NumeroLuces As Long
    NumeroParticulas As Long
    NumeroNPCs As Long
    NumeroOBJs As Long
    NumeroTE As Long
End Type

Private Type tDatosBloqueados
    X As Integer
    Y As Integer
End Type

Private Type tDatosGrh
    X As Integer
    Y As Integer
    grhindex As Long
End Type

Private Type tDatosTrigger
    X As Integer
    Y As Integer
    Trigger As Integer
End Type

Private Type tDatosLuces
    X As Integer
    Y As Integer
    color As Long
    Rango As Byte
End Type

Private Type tDatosParticulas
    X As Integer
    Y As Integer
    Particula As Long
End Type

Private Type tDatosNPC
    X As Integer
    Y As Integer
    NPCIndex As Integer
End Type

Private Type tDatosObjs
    X As Integer
    Y As Integer
    OBJIndex As Integer
    ObjAmmount As Integer
End Type

Private Type tDatosTE
    X As Integer
    Y As Integer
    DestM As Integer
    DestX As Integer
    DestY As Integer
End Type

Private Type tMapSize
    XMax As Integer
    XMin As Integer
    YMax As Integer
    YMin As Integer
End Type

Private Type tMapDat
    map_name As String * 64
    battle_mode As Byte
    backup_mode As Byte
    restrict_mode As String * 4
    music_number As String * 16
    zone As String * 16
    terrain As String * 16
    ambient As String * 16
    base_light As Long
    letter_grh As Long
    extra1 As Long
    extra2 As Long
    extra3 As String * 32
End Type

Private MapSize As tMapSize
Private MapDat As tMapDat

Public bFogata As Boolean
Public Sub General_Var_Write(ByVal file As String, ByVal Main As String, ByVal Var As String, ByVal value As String)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Writes a var to a text file
'*****************************************************************
    writeprivateprofilestring Main, Var, value, file
End Sub
Sub AddtoRichTextBox(ByRef RichTextBox As RichTextBox, ByVal Text As String, Optional ByVal red As Integer = -1, Optional ByVal green As Integer, Optional ByVal blue As Integer, Optional ByVal bold As Boolean = False, Optional ByVal italic As Boolean = False, Optional ByVal bCrLf As Boolean = False)

    With RichTextBox
        If Len(.Text) > 1000 Then
            'Get rid of first line
            .SelStart = InStr(1, .Text, vbCrLf) + 1
            .SelLength = Len(.Text) - .SelStart + 2
            .TextRTF = .SelRTF
        End If
        
        .SelStart = Len(RichTextBox.Text)
        .SelLength = 0
        .SelBold = bold
        .SelItalic = italic
        
        If Not red = -1 Then .SelColor = RGB(red, green, blue)
        
        .SelText = IIf(bCrLf, Text, Text & vbCrLf)

    End With
End Sub
Sub UnloadAllForms()
On Error Resume Next


    Dim mifrm As Form
    
    For Each mifrm In Forms
        Unload mifrm
    Next
End Sub

Function LegalCharacter(ByVal KeyAscii As Integer) As Boolean
'*****************************************************************
'Only allow characters that are Win 95 filename compatible
'*****************************************************************
    'if backspace allow
    If KeyAscii = 8 Then
        LegalCharacter = True
        Exit Function
    End If
    
    'Only allow space, numbers, letters and special characters
    If KeyAscii < 32 Or KeyAscii = 44 Then
        Exit Function
    End If
    
    If KeyAscii > 126 Then
        Exit Function
    End If
    
    'Check for bad special characters in between
    If KeyAscii = 34 Or KeyAscii = 42 Or KeyAscii = 47 Or KeyAscii = 58 Or KeyAscii = 60 Or KeyAscii = 62 Or KeyAscii = 63 Or KeyAscii = 92 Or KeyAscii = 124 Then
        Exit Function
    End If
    
    'else everything is cool
    LegalCharacter = True
End Function

Sub SetConnected()
'*****************************************************************
'Sets the client to "Connect" mode
'*****************************************************************
    'Set Connected
    Connected = True
    
    'Unload the connect form
    Unload frmCrearPersonaje
    Unload frmConnect
    Unload frmCrearCuenta
    Unload frmPanelAccount
    
    frmMain.lblNick = UserName
    
    Dim i As Integer
    For i = 1 To 11
    Next i
    
    'Load main form
    frmMain.Visible = True
    
    Call AddtoRichTextBox(frmMain.RecChat, "Bienvenido a CoverAO, estamos en una Etapa de cambios Implementando cosas Nuevas pedimos Tengan paciencia y Comprensión Los Invitamos a entrar a nuestra pagina Web www.CoverAO.jimdo.com Para mas Información.", 255, 255, 0, 0, 0)
   ' Call AddtoRichTextBox(frmMain.RecChat, "Al jugar nuestro servidor estás aceptando el reglamento de nuestra Pagina Web.", 255, 128, 1)
    'Call AddtoRichTextBox(frmMain.RecChat, "Para Más Informacion Pagina De Tu Ao.", 1, 1, 255)

End Sub




Private Sub CheckKeys()
'*****************************************************************
'Checks keys and respond
'*****************************************************************
    Static lastMovement As Long
    
    'No input allowed while Kega is not the active window
    If Not IsAppActive() Then Exit Sub
    
    'No walking when in commerce or banking.
    If Comerciando Then Exit Sub
    
    'No walking while writting in the forum.
    If frmForo.Visible Then Exit Sub

    'If game is paused, abort movement.
    If Pausa Then Exit Sub
    
    'Control movement interval (this enforces the 1 step loss when meditating / resting client-side)
    If GetTickCount - lastMovement > 56 Then
        lastMovement = GetTickCount
    Else
        Exit Sub
    End If
    
    'Don't allow any these keys during movement..
    If UserMoving = 0 Then
        If Not UserEstupido Then
            'Move Up
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyUp)) < 0 Then
                Call MoveTo(NORTH)
                Exit Sub
            End If
            
            'Move Right
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyRight)) < 0 Then
                Call MoveTo(EAST)
                Exit Sub
            End If
        
            'Move down
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyDown)) < 0 Then
                Call MoveTo(SOUTH)
                Exit Sub
            End If
        
            'Move left
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyLeft)) < 0 Then
                Call MoveTo(WEST)
                Exit Sub
            End If
            
            ' We haven't moved - Update 3D sounds!
            Call Audio.MoveListener(UserPos.X, UserPos.Y)
        Else
            Dim kp As Boolean
            kp = (GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyUp)) < 0) Or _
                GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyRight)) < 0 Or _
                GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyDown)) < 0 Or _
                GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyLeft)) < 0
            
            If kp Then
                Call RandomMove
            Else
                ' We haven't moved - Update 3D sounds!
                Call Audio.MoveListener(UserPos.X, UserPos.Y)
            End If
            
            Call DibujarMiniMapPos
        End If
    End If
End Sub

Sub SwitchMap(ByVal MapRoute As String, Optional Client_Mode As Boolean = False)

Engine.Char_Clean
Engine.Particle_Group_Remove_All
Engine.Light_Remove_All

On Error GoTo ErrorHandler

Dim fh As Integer
Dim MH As tMapHeader
Dim Blqs() As tDatosBloqueados
Dim L1() As Long
Dim L2() As tDatosGrh
Dim L3() As tDatosGrh
Dim L4() As tDatosGrh
Dim Triggers() As tDatosTrigger
Dim Luces() As tDatosLuces
Dim Particulas() As tDatosParticulas
Dim Objetos() As tDatosObjs
Dim NPCs() As tDatosNPC
Dim TEs() As tDatosTE

Dim i As Long
Dim j As Long

Extract_File Maps, App.Path & "\Recursos\", "mapa" & MapRoute & ".csm", App.Path & "\Recursos\"

fh = FreeFile
Open App.Path & "\Recursos\mapa" & MapRoute & ".csm" For Binary Access Read As fh
    Get #fh, , MH
    Get #fh, , MapSize
    Get #fh, , MapDat
    
    ReDim MapData(MapSize.XMin To MapSize.XMax, MapSize.YMin To MapSize.YMax) As MapBlock
    ReDim L1(MapSize.XMin To MapSize.XMax, MapSize.YMin To MapSize.YMax) As Long
    
    Get #fh, , L1
    
    With MH
        If .NumeroBloqueados > 0 Then
            ReDim Blqs(1 To .NumeroBloqueados)
            Get #fh, , Blqs
            For i = 1 To .NumeroBloqueados
                MapData(Blqs(i).X, Blqs(i).Y).Blocked = 1
            Next i
        End If
        
        If .NumeroLayers(2) > 0 Then
            ReDim L2(1 To .NumeroLayers(2))
            Get #fh, , L2
            For i = 1 To .NumeroLayers(2)
                InitGrh MapData(L2(i).X, L2(i).Y).Graphic(2), L2(i).grhindex
            Next i
        End If
        
        If .NumeroLayers(3) > 0 Then
            ReDim L3(1 To .NumeroLayers(3))
            Get #fh, , L3
            For i = 1 To .NumeroLayers(3)
                InitGrh MapData(L3(i).X, L3(i).Y).Graphic(3), L3(i).grhindex
            Next i
        End If
        
        If .NumeroLayers(4) > 0 Then
            ReDim L4(1 To .NumeroLayers(4))
            Get #fh, , L4
            For i = 1 To .NumeroLayers(4)
                InitGrh MapData(L4(i).X, L4(i).Y).Graphic(4), L4(i).grhindex
            Next i
        End If
        
        If .NumeroTriggers > 0 Then
            ReDim Triggers(1 To .NumeroTriggers)
            Get #fh, , Triggers
            For i = 1 To .NumeroTriggers
                MapData(Triggers(i).X, Triggers(i).Y).Trigger = Triggers(i).Trigger
            Next i
        End If
        
        If .NumeroParticulas > 0 Then
            ReDim Particulas(1 To .NumeroParticulas)
            Get #fh, , Particulas
            For i = 1 To .NumeroParticulas
                MapData(Particulas(i).X, Particulas(i).Y).particle_group_index = General_Particle_Create(Particulas(i).Particula, Particulas(i).X, Particulas(i).Y)
            Next i
        End If
        
        If .NumeroLuces > 0 Then
            ReDim Luces(1 To .NumeroLuces)
            Get #fh, , Luces
            For i = 1 To .NumeroLuces
                Call Engine.Light_Create(Luces(i).X, Luces(i).Y, Luces(i).color, Luces(i).Rango)
            Next i
        End If
        
    End With

Close fh


For j = MapSize.YMin To MapSize.YMax
    For i = MapSize.XMin To MapSize.XMax
        If L1(i, j) > 0 Then
            InitGrh MapData(i, j).Graphic(1), L1(i, j)
        End If
    Next i
Next j

Dim r As Integer, g As Integer, b As Integer
'Common light value verify
If MapDat.base_light = 0 Then
    map_base_light = -1
Else
    General_Long_Color_to_RGB MapDat.base_light, r, g, b
    map_base_light = D3DColorXRGB(r, g, b)
End If

'*******************************
'Render lights
Engine.Light_Render_All
'*******************************
Debug.Print "mapa" & MapRoute & ":" & MapDat.extra1
Debug.Print "mapa" & MapRoute & ":" & MapDat.extra2
Debug.Print "mapa" & MapRoute & ":" & MapDat.zone
Debug.Print "mapa" & MapRoute & ":" & MapDat.battle_mode
Debug.Print "mapa" & MapRoute & ":" & MapDat.terrain
frmMain.Minimap.Cls

    Dim map_x As Long
    Dim map_y As Long
    Dim screen_x As Long
    Dim screen_y As Long
    Dim grh_index As Long
    
    '*********************
    'Layer 1 (Floor level) and walls
    '*********************
    For map_y = MapSize.XMin To MapSize.XMax
        For map_x = MapSize.YMin To MapSize.YMax
            screen_x = (map_x - 1) * 2
            screen_y = (map_y - 1) * 2
            '*** Start Layer 1 ***
            If MapData(map_x, map_y).Graphic(1).grhindex Then
                grh_index = MapData(map_x, map_y).Graphic(1).grhindex
                SetPixel frmMain.Minimap.hdc, map_x, map_y, GrhData(grh_index).mini_map_color
            End If
            '*** End Layer 1 ***
        Next map_x
    Next map_y
    
    For map_y = MapSize.XMin To MapSize.XMax
        For map_x = MapSize.YMin To MapSize.YMax
            screen_x = (map_x - 1) * 2
            screen_y = (map_y - 1) * 2
            '*** Start Layer 2 ***
            If MapData(map_x, map_y).Graphic(2).grhindex Then
                grh_index = MapData(map_x, map_y).Graphic(2).grhindex
                SetPixel frmMain.Minimap.hdc, map_x, map_y, GrhData(grh_index).mini_map_color
            End If
            '*** End Layer 2 ***
        Next map_x
    Next map_y

    For map_y = MapSize.XMin To MapSize.XMax
        For map_x = MapSize.YMin To MapSize.YMax
            screen_x = (map_x - 1) * 2
            screen_y = (map_y - 1) * 2
            '*** Start Layer 2 ***
            If MapData(map_x, map_y).Graphic(4).grhindex Then
                grh_index = MapData(map_x, map_y).Graphic(4).grhindex
                SetPixel frmMain.Minimap.hdc, map_x, map_y, GrhData(grh_index).mini_map_color
            End If
            '*** End Layer 2 ***
        Next map_x
    Next map_y
MapDat.map_name = Trim$(MapDat.map_name)

Exit Sub

ErrorHandler:
    If fh <> 0 Then Close fh
End Sub

Function ReadField(ByVal Pos As Integer, ByRef Text As String, ByVal SepASCII As Byte) As String
'*****************************************************************
'Gets a field from a delimited string
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/15/2004
'*****************************************************************
    Dim i As Long
    Dim LastPos As Long
    Dim CurrentPos As Long
    Dim delimiter As String * 1
    
    delimiter = Chr$(SepASCII)
    
    For i = 1 To Pos
        LastPos = CurrentPos
        CurrentPos = InStr(LastPos + 1, Text, delimiter, vbBinaryCompare)
    Next i
    
    If CurrentPos = 0 Then
        ReadField = mid$(Text, LastPos + 1, Len(Text) - LastPos)
    Else
        ReadField = mid$(Text, LastPos + 1, CurrentPos - LastPos - 1)
    End If
End Function

Function FieldCount(ByRef Text As String, ByVal SepASCII As Byte) As Long
'*****************************************************************
'Gets the number of fields in a delimited string
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 07/29/2007
'*****************************************************************
    Dim count As Long
    Dim curPos As Long
    Dim delimiter As String * 1
    
    If LenB(Text) = 0 Then Exit Function
    
    delimiter = Chr$(SepASCII)
    
    curPos = 0
    
    Do
        curPos = InStr(curPos + 1, Text, delimiter)
        count = count + 1
    Loop While curPos <> 0
    
    FieldCount = count
End Function

Function FileExist(ByVal file As String, ByVal FileType As VbFileAttribute) As Boolean
    FileExist = (Dir$(file, FileType) <> "")
End Function

Sub Main()
    
    Set SurfaceDB = New clsTexManager
    Set Audio = New clsAudio
    
    Call Protocol.InitFonts
    Call InicializarNombres
    Call Engine.setup_ambient
    
    Dim eligen As Byte
    Call Engine.Init(eligen)
    
    frmCargando.Show
    frmCargando.Refresh
    Call frmCargando.EstablecerProgreso(0)

    Call frmCargando.progresoConDelay(50)
    
    Call LoadGrhData
    Call frmCargando.progresoConDelay(55)
    Call CargarCuerpos
    Call CargarCabezas
    Call frmCargando.progresoConDelay(60)
    Call CargarCascos
    Call CargarFxs
    Call frmCargando.progresoConDelay(65)
    Call CargarParticulas
    Call CargarAnimArmas
    Call frmCargando.progresoConDelay(70)
    Call CargarAnimEscudos
    Call LoadMacros
    Call frmCargando.progresoConDelay(90)
    
    'Inicializamos el sonido
    Call Audio.Initialize(frmMain.hWnd, App.Path & "\RECURSOS\WAV\", App.Path & "\RECURSOS\MIDI\")
    Call frmCargando.progresoConDelay(95)
    
    'Inicializamos el inventario gráfico
    Call Inventario.Initialize(frmMain.picInv)
    frmMain.Socket1.Startup
    Call frmCargando.progresoConDelay(100)

    Unload frmCargando
    Call frmCargando.progresoConDelay(100)
    frmCargando.Visible = False
    Unload frmCargando
    
    'Inicialización de variables globales
    PrimeraVez = True
    prgRun = True
    Pausa = False
    
    
    frmConnect.Visible = True
    
    'Set the intervals of timers
    Call MainTimer.SetInterval(TimersIndex.Attack, INT_ATTACK)
    Call MainTimer.SetInterval(TimersIndex.Work, INT_WORK)
    Call MainTimer.SetInterval(TimersIndex.UseItemWithU, INT_USEITEMU)
    Call MainTimer.SetInterval(TimersIndex.UseItemWithDblClick, INT_USEITEMDCK)
    Call MainTimer.SetInterval(TimersIndex.SendRPU, INT_SENTRPU)
    Call MainTimer.SetInterval(TimersIndex.CastSpell, INT_CAST_SPELL)
    Call MainTimer.SetInterval(TimersIndex.Arrows, INT_ARROWS)
    Call MainTimer.SetInterval(TimersIndex.CastAttack, INT_CAST_ATTACK)
    
    frmMain.macrotrabajo.Interval = INT_MACRO_TRABAJO
    frmMain.macrotrabajo.Enabled = False
    
   'Init timers
    Call MainTimer.Start(TimersIndex.Attack)
    Call MainTimer.Start(TimersIndex.Work)
    Call MainTimer.Start(TimersIndex.UseItemWithU)
    Call MainTimer.Start(TimersIndex.UseItemWithDblClick)
    Call MainTimer.Start(TimersIndex.SendRPU)
    Call MainTimer.Start(TimersIndex.CastSpell)
    Call MainTimer.Start(TimersIndex.Arrows)
    Call MainTimer.Start(TimersIndex.CastAttack)
    
    Do While prgRun
        'Sólo dibujamos si la ventana no está minimizada
        If IsAppActive Then
        
            If frmMain.Visible Then
                Call Engine.Render
                Call RenderSounds
                Call CheckKeys
            End If
            
            If RenderInv And frmMain.Visible Then Engine.DrawInv
            
        Else
            If frmMain.Visible Then RenderInv = True
            Sleep 10
        End If

        ' If there is anything to be sent, we send it
        Call FlushBuffer
        
        DoEvents
    Loop
    
    Call CloseClient
End Sub




Private Sub InicializarNombres()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/27/2005
'Inicializa los nombres de razas, ciudades, clases, skills, atributos, etc.
'**************************************************************
    Ciudades(eCiudad.cNix) = "Nix"
    
    ListaRazas(eRaza.HUMANO) = "Humano"
    ListaRazas(eRaza.ELFO) = "Elfo"
    ListaRazas(eRaza.ElfoOscuro) = "Drow"
    ListaRazas(eRaza.Gnomo) = "Gnomo"
    ListaRazas(eRaza.Enano) = "Enano"
    ListaRazas(eRaza.Orco) = "Orco"

    ListaClases(eClass.Mago) = "Mago"
    ListaClases(eClass.Clerigo) = "Clerigo"
    ListaClases(eClass.Guerrero) = "Guerrero"
    ListaClases(eClass.Asesino) = "Asesino"
    ListaClases(eClass.Ladron) = "Ladron"
    ListaClases(eClass.Bardo) = "Bardo"
    ListaClases(eClass.Druida) = "Druida"
    ListaClases(eClass.Paladin) = "Paladin"
    ListaClases(eClass.Cazador) = "Cazador"
    ListaClases(eClass.Pescador) = "Pescador"
    ListaClases(eClass.Herrero) = "Herrero"
    ListaClases(eClass.Leñador) = "Leñador"
    ListaClases(eClass.Minero) = "Minero"
    ListaClases(eClass.Carpintero) = "Carpintero"
    ListaClases(eClass.Mercenario) = "Mercenario"
    ListaClases(eClass.Nigromonte) = "Nigromonte"
    ListaClases(eClass.Sastre) = "Sastre"
    ListaClases(eClass.Gladiador) = "Gladiador"
    
    SkillsNames(eSkill.Magia) = "Magia"
    SkillsNames(eSkill.Robar) = "Robar"
    SkillsNames(eSkill.Tacticas) = "Tacticas de combate"
    SkillsNames(eSkill.Armas) = "Combate con armas"
    SkillsNames(eSkill.Meditar) = "Meditar"
    SkillsNames(eSkill.Apuñalar) = "Apuñalar"
    SkillsNames(eSkill.Ocultarse) = "Ocultarse"
    SkillsNames(eSkill.Supervivencia) = "Supervivencia"
    SkillsNames(eSkill.Talar) = "Talar arboles"
    SkillsNames(eSkill.Comercio) = "Comercio"
    SkillsNames(eSkill.DefensaEscudos) = "Defensa con escudos"
    SkillsNames(eSkill.Pesca) = "Pesca"
    SkillsNames(eSkill.Mineria) = "Mineria"
    SkillsNames(eSkill.Carpinteria) = "Carpinteria"
    SkillsNames(eSkill.Herreria) = "Herreria"
    SkillsNames(eSkill.Liderazgo) = "Liderazgo"
    SkillsNames(eSkill.Domar) = "Domar animales"
    SkillsNames(eSkill.Proyectiles) = "Armas de proyectiles"
    SkillsNames(eSkill.Artes) = "Artes Marciales"
    SkillsNames(eSkill.Navegacion) = "Navegacion"
    SkillsNames(eSkill.Alquimia) = "Alquimia"
    SkillsNames(eSkill.Arrojadizas) = "Armas Arrojadizas"
    SkillsNames(eSkill.Botanica) = "Botanica"
    SkillsNames(eSkill.Equitacion) = "Equitacion"
    SkillsNames(eSkill.Musica) = "Musica"
    SkillsNames(eSkill.Resistencia) = "Resistencia Magica"
    SkillsNames(eSkill.Sastreria) = "Sastreria"

    AtributosNames(eAtributos.Fuerza) = "Fuerza"
    AtributosNames(eAtributos.Agilidad) = "Agilidad"
    AtributosNames(eAtributos.Inteligencia) = "Inteligencia"
    AtributosNames(eAtributos.Carisma) = "Carisma"
    AtributosNames(eAtributos.Constitucion) = "Constitucion"
    
    ReDim Head_Range(1 To NUMRAZAS) As tHeadRange

'Male heads
Head_Range(HUMANO).mStart = 1
Head_Range(HUMANO).mEnd = 30
Head_Range(Enano).mStart = 301
Head_Range(Enano).mEnd = 315
Head_Range(ELFO).mStart = 101
Head_Range(ELFO).mEnd = 121
Head_Range(ElfoOscuro).mStart = 202
Head_Range(ElfoOscuro).mEnd = 212
Head_Range(Gnomo).mStart = 401
Head_Range(Gnomo).mEnd = 409
Head_Range(Orco).mStart = 501
Head_Range(Orco).mEnd = 514

'Female heads
Head_Range(HUMANO).fStart = 70
Head_Range(HUMANO).fEnd = 80
Head_Range(Enano).fStart = 370
Head_Range(Enano).fEnd = 373
Head_Range(ELFO).fStart = 170
Head_Range(ELFO).fEnd = 189
Head_Range(ElfoOscuro).fStart = 270
Head_Range(ElfoOscuro).fEnd = 278
Head_Range(Gnomo).fStart = 470
Head_Range(Gnomo).fEnd = 481
Head_Range(Orco).fStart = 570
Head_Range(Orco).fEnd = 573
End Sub

''
' Removes all text from the console and dialogs


Public Sub Auto_Drag(ByVal hWnd As Long)
Call ReleaseCapture
Call SendMessage(hWnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&)
End Sub
Public Sub CloseClient()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 8/14/2007
'Frees all used resources, cleans up and leaves
'**************************************************************
    ' Allow new instances of the client to be opened
    Call PrevInstance.ReleaseInstance
    
    'Stop tile engine
    Call Engine.EndInit
    
    'Destruimos los objetos públicos creados
    Set CustomKeys = Nothing
    Set SurfaceDB = Nothing
    Set Dialogos = Nothing
    Set DialogosClanes = Nothing
    Set Audio = Nothing
    Set Inventario = Nothing
    Set MainTimer = Nothing
    Set incomingData = Nothing
    Set outgoingData = Nothing
    
    Call UnloadAllForms
    
    End
End Sub
Public Function General_Char_Particle_Create(ByVal ParticulaInd As Long, ByVal char_index As Integer, ByVal PartPos As Byte, Optional ByVal particle_life As Long = 0) As Long

On Error Resume Next

If ParticulaInd <= 0 Then Exit Function

Dim rgb_list(0 To 3) As Long
rgb_list(0) = RGB(StreamData(ParticulaInd).colortint(0).r, StreamData(ParticulaInd).colortint(0).g, StreamData(ParticulaInd).colortint(0).b)
rgb_list(1) = RGB(StreamData(ParticulaInd).colortint(1).r, StreamData(ParticulaInd).colortint(1).g, StreamData(ParticulaInd).colortint(1).b)
rgb_list(2) = RGB(StreamData(ParticulaInd).colortint(2).r, StreamData(ParticulaInd).colortint(2).g, StreamData(ParticulaInd).colortint(2).b)
rgb_list(3) = RGB(StreamData(ParticulaInd).colortint(3).r, StreamData(ParticulaInd).colortint(3).g, StreamData(ParticulaInd).colortint(3).b)

'General_Char_Particle_Create = engine.Char_Particle_Group_Create(char_index, StreamData(ParticulaInd).grh_list, rgb_list(), PartPos, StreamData(ParticulaInd).NumOfParticles, ParticulaInd, _
    StreamData(ParticulaInd).AlphaBlend, IIf(particle_life = 0, StreamData(ParticulaInd).life_counter, particle_life), StreamData(ParticulaInd).speed, , StreamData(ParticulaInd).x1, StreamData(ParticulaInd).y1, StreamData(ParticulaInd).angle, _
    StreamData(ParticulaInd).vecx1, StreamData(ParticulaInd).vecx2, StreamData(ParticulaInd).vecy1, StreamData(ParticulaInd).vecy2, _
    StreamData(ParticulaInd).life1, StreamData(ParticulaInd).life2, StreamData(ParticulaInd).friction, StreamData(ParticulaInd).spin_speedL, _
    StreamData(ParticulaInd).gr, StreamData(ParticulaInd).gravity, StreamData(ParticulaInd).grav_strength, StreamData(ParticulaInd).bounce_strength, StreamData(ParticulaInd).x2, _
    StreamData(ParticulaInd).y2, StreamData(ParticulaInd).XMove, StreamData(ParticulaInd).move_x1, StreamData(ParticulaInd).move_x2, StreamData(ParticulaInd).move_y1, _
    StreamData(ParticulaInd).move_y2, StreamData(ParticulaInd).YMove, StreamData(ParticulaInd).spin_speedH, StreamData(ParticulaInd).spin)

End Function

Public Function General_Particle_Create(ByVal ParticulaInd As Long, ByVal X As Integer, ByVal Y As Integer, Optional ByVal particle_life As Long = 0, Optional ByVal OffsetX As Integer, Optional ByVal OffsetY As Integer) As Long

Dim rgb_list(0 To 3) As Long
rgb_list(0) = RGB(StreamData(ParticulaInd).colortint(0).r, StreamData(ParticulaInd).colortint(0).g, StreamData(ParticulaInd).colortint(0).b)
rgb_list(1) = RGB(StreamData(ParticulaInd).colortint(1).r, StreamData(ParticulaInd).colortint(1).g, StreamData(ParticulaInd).colortint(1).b)
rgb_list(2) = RGB(StreamData(ParticulaInd).colortint(2).r, StreamData(ParticulaInd).colortint(2).g, StreamData(ParticulaInd).colortint(2).b)
rgb_list(3) = RGB(StreamData(ParticulaInd).colortint(3).r, StreamData(ParticulaInd).colortint(3).g, StreamData(ParticulaInd).colortint(3).b)

General_Particle_Create = Engine.Particle_Group_Create(X, Y, StreamData(ParticulaInd).grh_list, rgb_list(), StreamData(ParticulaInd).NumOfParticles, ParticulaInd, _
    StreamData(ParticulaInd).AlphaBlend, IIf(particle_life = 0, StreamData(ParticulaInd).life_counter, particle_life), StreamData(ParticulaInd).speed, , StreamData(ParticulaInd).x1 + OffsetX, StreamData(ParticulaInd).y1 + OffsetY, StreamData(ParticulaInd).angle, _
    StreamData(ParticulaInd).vecx1, StreamData(ParticulaInd).vecx2, StreamData(ParticulaInd).vecy1, StreamData(ParticulaInd).vecy2, _
    StreamData(ParticulaInd).life1, StreamData(ParticulaInd).life2, StreamData(ParticulaInd).friction, StreamData(ParticulaInd).spin_speedL, _
    StreamData(ParticulaInd).gravity, StreamData(ParticulaInd).grav_strength, StreamData(ParticulaInd).bounce_strength, StreamData(ParticulaInd).x2, _
    StreamData(ParticulaInd).y2, StreamData(ParticulaInd).XMove, StreamData(ParticulaInd).move_x1, StreamData(ParticulaInd).move_x2, StreamData(ParticulaInd).move_y1, _
    StreamData(ParticulaInd).move_y2, StreamData(ParticulaInd).YMove, StreamData(ParticulaInd).spin_speedH, StreamData(ParticulaInd).spin)

End Function
Public Function Map_NameLoad(ByVal map_num As Integer) As String

On Error GoTo ErrorHandler

If FileExist(App.Path & "\Recursos\Mapas\mapa" & map_num & ".csm", vbNormal) Then
    SwitchMap map_num
    Map_NameLoad = MapDat.map_name
    If LenB(Map_NameLoad) = 0 Then
        Map_NameLoad = "Mapa Desconocido"
    End If
Else
    Map_NameLoad = "Mapa Desconocido"
End If

Exit Function

ErrorHandler:
    Map_NameLoad = "Mapa Desconocido"

End Function
Public Sub General_Long_Color_to_RGB(ByVal long_color As Long, ByRef red As Integer, ByRef green As Integer, ByRef blue As Integer)
'***********************************
'Coded by Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'Last Modified: 2/19/03
'Takes a long value and separates RGB values to the given variables
'***********************************
    Dim temp_color As String
    
    temp_color = Hex(long_color)
    If Len(temp_color) < 6 Then
        'Give is 6 digits for easy RGB conversion.
        temp_color = String(6 - Len(temp_color), "0") + temp_color
    End If
    
    red = CLng("&H" + mid$(temp_color, 1, 2))
    green = CLng("&H" + mid$(temp_color, 3, 2))
    blue = CLng("&H" + mid$(temp_color, 5, 2))
End Sub
Public Function General_Get_Mouse_Speed() As Long
'**************************************************************
'Author: Augusto José Rando
'Last Modify Date: 6/11/2005
'
'**************************************************************
 
SystemParametersInfo SPI_GETMOUSESPEED, 0, General_Get_Mouse_Speed, 0
 
End Function
 
Public Sub General_Set_Mouse_Speed(ByVal lngSpeed As Long)
'**************************************************************
'Author: Augusto José Rando
'Last Modify Date: 6/11/2005
'
'**************************************************************
 
SystemParametersInfo SPI_SETMOUSESPEED, 0, ByVal lngSpeed, 0
 
End Sub
