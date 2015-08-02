Attribute VB_Name = "Mod_Declaraciones"
'Argentum Online 0.11.6
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
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

Public LuzMouse As Boolean

Public MouseS As Long

Public Type Macroo
    Comando As String
    Equipar As Byte
    Usar As Byte
    Hechizo As Byte
End Type
 
Public Macros(9) As Macroo

'Objetos públicos
Public DialogosClanes As New clsGuildDlg
Public Dialogos As New clsDialogs
Public Audio As New clsAudio
Public Inventario As New clsGrapchicalInventory
Public CustomKeys As New clsCustomKeys
Public MainTimer As New clsTimer
Public incomingData As New clsByteQueue
Public outgoingData As New clsByteQueue
Public UserIndex As Integer
Public Windows_Temp_Dir As String

'Sonidos
Public Const SND_CLICK As String = "click.Wav"
Public Const SND_PASOS1 As String = "23.Wav"
Public Const SND_PASOS2 As String = "24.Wav"
Public Const SND_NAVEGANDO As String = "50.wav"
Public Const SND_OVER As String = "click2.Wav"
Public Const SND_DICE As String = "cupdice.Wav"
Public Const SND_LLUVIAINEND As String = "lluviainend.wav"
Public Const SND_LLUVIAOUTEND As String = "lluviaoutend.wav"

' Head index of the casper. Used to know if a char is killed

' Constantes de intervalo
Public Const INT_MACRO_HECHIS As Integer = 1000
Public Const INT_MACRO_TRABAJO As Integer = 900

Public Const INT_ATTACK As Integer = 600
Public Const INT_ARROWS As Integer = 600
Public Const INT_CAST_SPELL As Integer = 800
Public Const INT_CAST_ATTACK As Integer = 800
Public Const INT_WORK As Integer = 700
Public Const INT_USEITEMU As Integer = 450
Public Const INT_USEITEMDCK As Integer = 125
Public Const INT_SENTRPU As Integer = 2000

Public MacroBltIndex As Integer

Public Const CASPER_HEAD As Integer = 500
Public Const FRAGATA_FANTASMAL As Integer = 87

Public Const NUMATRIBUTES As Byte = 5

Public RenderInv As Boolean
Public Default_RGB(0 To 3) As Long

Public CreandoClan As Boolean
Public ClanName As String
Public Site As String

Public UserCiego As Boolean
Public UserEstupido As Boolean

Public NoRes As Boolean 'no cambiar la resolucion

Public RainBufferIndex As Long
Public FogataBufferIndex As Long

Public Const bCabeza = 1
Public Const bPiernaIzquierda = 2
Public Const bPiernaDerecha = 3
Public Const bBrazoDerecho = 4
Public Const bBrazoIzquierdo = 5
Public Const bTorso = 6

'Timers de GetTickCount
Public Const tAt = 2000
Public Const tUs = 600

Public Const PrimerBodyBarco = 84
Public Const UltimoBodyBarco = 87

Public NumEscudosAnims As Integer

Public ArmasHerrero(0 To 100) As Integer
Public ArmadurasHerrero(0 To 100) As Integer
Public ObjCarpintero(0 To 100) As Integer

Public Versiones(1 To 7) As Integer

Public UsaMacro As Boolean
Public CnTd As Byte

Public Const MAX_BANCOINVENTORY_SLOTS As Byte = 40
Public UserBancoInventory(1 To MAX_BANCOINVENTORY_SLOTS) As Inventory

'Direcciones
Public Enum E_Heading
    NORTH = 1
    EAST = 2
    SOUTH = 3
    WEST = 4
End Enum

'Lista de cabezas
Public Type tIndiceCabeza
    Head(1 To 4) As Integer
End Type

Public Type tIndiceCuerpo
    body(1 To 4) As Integer
    HeadOffsetX As Integer
    HeadOffsetY As Integer
End Type

Public Type tIndiceFx
    Animacion As Integer
    OffsetX As Integer
    OffsetY As Integer
End Type

'Objetos
Public Const MAX_INVENTORY_OBJS As Integer = 10000
Public Const MAX_INVENTORY_SLOTS As Byte = 25
Public Const MAX_NPC_INVENTORY_SLOTS As Byte = 50
Public Const MAXHECHI As Byte = 35

Public Const MAXSKILLPOINTS As Byte = 100

Public Const FLAGORO As Integer = MAX_INVENTORY_SLOTS + 1

Public Const Fogata As Integer = 1521

Public Enum eClass
    Clerigo = 1
    Mago = 2
    Guerrero = 3
    Asesino = 4
    Ladron = 5
    Bardo = 6
    Druida = 7
    Gladiador = 8
    Paladin = 9
    Cazador = 10
    Pescador = 11
    Herrero = 12
    Leñador = 13
    Minero = 14
    Carpintero = 15
    Sastre = 16
    Mercenario = 17
    Nigromonte = 18
End Enum

Public Enum eCiudad
    cNix = 1
End Enum

Enum eRaza
    HUMANO = 1
    ELFO
    ElfoOscuro
    Gnomo
    Enano
    Orco
End Enum

Public Enum eSkill
    Tacticas = 1
    Armas = 2
    Artes = 3
    Apuñalar = 4
    Arrojadizas = 5
    Proyectiles = 6
    DefensaEscudos = 7
    Magia = 8
    Resistencia = 9
    Meditar = 10
    Ocultarse = 11
    Domar = 12
    Musica = 13
    Robar = 14
    Comercio = 15
    Supervivencia = 16
    Liderazgo = 17
    Pesca = 18
    Mineria = 19
    Talar = 20
    Botanica = 21
    Herreria = 22
    Carpinteria = 23
    Alquimia = 24
    Sastreria = 25
    Navegacion = 26
    Equitacion = 27
End Enum

Public Enum eAtributos
    Fuerza = 1
    Agilidad = 2
    Inteligencia = 3
    Carisma = 4
    Constitucion = 5
End Enum

Enum eGenero
    Hombre = 1
    Mujer
End Enum

Public Enum PlayerType
    User = &H1
    Consejero = &H2
    SemiDios = &H4
    Dios = &H8
    Admin = &H10
    RoleMaster = &H20
    ChaosCouncil = &H40
    RoyalCouncil = &H80
End Enum

Public Enum eObjType
    otUseOnce = 1
    otWeapon = 2
    otArmadura = 3
    otArboles = 4
    otGuita = 5
    otPuertas = 6
    otContenedores = 7
    otCarteles = 8
    otLlaves = 9
    otForos = 10
    otPociones = 11
    otBebidas = 13
    otLeña = 14
    otFogata = 15
    otESCUDO = 16
    otCASCO = 17
    otAnillo = 18
    otTeleport = 19
    otYacimiento = 22
    otMinerales = 23
    otPergaminos = 24
    otInstrumentos = 26
    otYunque = 27
    otFragua = 28
    otBarcos = 31
    otFlechas = 32
    otBotellaVacia = 33
    otBotellaLlena = 34
    otManchas = 35          'No se usa
    otMapas = 36
    otCualquiera = 1000
End Enum

Type tHeadRange
    mStart As Integer
    mEnd As Integer
    fStart As Integer
    fEnd As Integer
End Type
Public Head_Range() As tHeadRange

Public Const FundirMetal As Integer = 88

'
' Mensajes
'
' MENSAJE_*  --> Mensajes de texto que se muestran en el cuadro de texto
'

Public Const MENSAJE_CRIATURA_FALLA_GOLPE As String = "¡¡¡La criatura fallo el golpe!!!"
Public Const MENSAJE_CRIATURA_MATADO As String = "¡¡¡La criatura te ha matado!!!"
Public Const MENSAJE_RECHAZO_ATAQUE_ESCUDO As String = "¡¡¡Has rechazado el ataque con el escudo!!!"
Public Const MENSAJE_USUARIO_RECHAZO_ATAQUE_ESCUDO  As String = "¡¡¡El usuario rechazo el ataque con su escudo!!!"
Public Const MENSAJE_FALLADO_GOLPE As String = "¡¡¡Has fallado el golpe!!!"
Public Const MENSAJE_SEGURO_ACTIVADO As String = "Seguro activado"
Public Const MENSAJE_SEGURO_DESACTIVADO As String = "Seguro desactivado"
Public Const MENSAJE_PIERDE_NOBLEZA As String = "¡¡Has perdido puntaje de nobleza y ganado puntaje de criminalidad!! Si sigues ayudando a criminales te convertirás en uno de ellos y serás perseguido por las tropas de las ciudades."
Public Const MENSAJE_USAR_MEDITANDO As String = "¡Estás meditando! Debes dejar de meditar para usar objetos."

Public Const MENSAJE_SEGURO_RESU_ON As String = "SEGURO DE RESURRECCION ACTIVADO"
Public Const MENSAJE_SEGURO_RESU_OFF As String = "SEGURO DE RESURRECCION DESACTIVADO"

Public Const MENSAJE_GOLPE_CABEZA As String = "¡¡La criatura te ha pegado en la cabeza por "
Public Const MENSAJE_GOLPE_BRAZO_IZQ As String = "¡¡La criatura te ha pegado el brazo izquierdo por "
Public Const MENSAJE_GOLPE_BRAZO_DER As String = "¡¡La criatura te ha pegado el brazo derecho por "
Public Const MENSAJE_GOLPE_PIERNA_IZQ As String = "¡¡La criatura te ha pegado la pierna izquierda por "
Public Const MENSAJE_GOLPE_PIERNA_DER As String = "¡¡La criatura te ha pegado la pierna derecha por "
Public Const MENSAJE_GOLPE_TORSO  As String = "¡¡La criatura te ha pegado en el torso por "

' MENSAJE_[12]: Aparecen antes y despues del valor de los mensajes anteriores (MENSAJE_GOLPE_*)
Public Const MENSAJE_1 As String = "¡¡"
Public Const MENSAJE_2 As String = "!!"

Public Const MENSAJE_GOLPE_CRIATURA_1 As String = "¡¡Le has pegado a la criatura por "

Public Const MENSAJE_ATAQUE_FALLO As String = " te ataco y fallo!!"

Public Const MENSAJE_RECIVE_IMPACTO_CABEZA As String = " te ha pegado en la cabeza por "
Public Const MENSAJE_RECIVE_IMPACTO_BRAZO_IZQ As String = " te ha pegado el brazo izquierdo por "
Public Const MENSAJE_RECIVE_IMPACTO_BRAZO_DER As String = " te ha pegado el brazo derecho por "
Public Const MENSAJE_RECIVE_IMPACTO_PIERNA_IZQ As String = " te ha pegado la pierna izquierda por "
Public Const MENSAJE_RECIVE_IMPACTO_PIERNA_DER As String = " te ha pegado la pierna derecha por "
Public Const MENSAJE_RECIVE_IMPACTO_TORSO As String = " te ha pegado en el torso por "

Public Const MENSAJE_PRODUCE_IMPACTO_1 As String = "¡¡Le has pegado a "
Public Const MENSAJE_PRODUCE_IMPACTO_CABEZA As String = " en la cabeza por "
Public Const MENSAJE_PRODUCE_IMPACTO_BRAZO_IZQ As String = " en el brazo izquierdo por "
Public Const MENSAJE_PRODUCE_IMPACTO_BRAZO_DER As String = " en el brazo derecho por "
Public Const MENSAJE_PRODUCE_IMPACTO_PIERNA_IZQ As String = " en la pierna izquierda por "
Public Const MENSAJE_PRODUCE_IMPACTO_PIERNA_DER As String = " en la pierna derecha por "
Public Const MENSAJE_PRODUCE_IMPACTO_TORSO As String = " en el torso por "

Public Const MENSAJE_TRABAJO_MAGIA As String = "Haz click sobre el objetivo..."
Public Const MENSAJE_TRABAJO_PESCA As String = "Haz click sobre el sitio donde quieres pescar..."
Public Const MENSAJE_TRABAJO_ROBAR As String = "Haz click sobre la víctima..."
Public Const MENSAJE_TRABAJO_TALAR As String = "Haz click sobre el árbol..."
Public Const MENSAJE_TRABAJO_MINERIA As String = "Haz click sobre el yacimiento..."
Public Const MENSAJE_TRABAJO_FUNDIRMETAL As String = "Haz click sobre la fragua..."
Public Const MENSAJE_TRABAJO_PROYECTILES As String = "Haz click sobre la victima..."

Public Const MENSAJE_ENTRAR_PARTY_1 As String = "Si deseas entrar en una party con "
Public Const MENSAJE_ENTRAR_PARTY_2 As String = ", escribe /entrarparty"

Public Const MENSAJE_NENE As String = "Cantidad de NPCs: "

'Inventario
Type Inventory
    OBJIndex As Integer
    name As String
    grhindex As Integer
    Amount As Long
    Equipped As Byte
    Valor As Single
    OBJType As Integer
    Def As Integer
    MaxHit As Integer
    MinHit As Integer
    PuedeUsar As Byte
End Type

Type NpCinV
    OBJIndex As Integer
    name As String
    grhindex As Integer
    Amount As Integer
    Valor As Single
    OBJType As Integer
    Def As Integer
    MaxHit As Integer
    MinHit As Integer
    C1 As String
    C2 As String
    C3 As String
    C4 As String
    C5 As String
    C6 As String
    C7 As String
End Type

Type tReputacion 'Fama del usuario
    NobleRep As Long
    BurguesRep As Long
    PlebeRep As Long
    LadronesRep As Long
    BandidoRep As Long
    AsesinoRep As Long
    
    Promedio As Long
End Type

Type tEstadisticasUsu
    CiudadanosMatados As Long
    RenegadosMatados As Long
    RepublicanosMatados As Long
    ArmadaMatados As Long
    MiliciaMatados As Long
    CaosMatados As Long
    UsuariosMatados As Long
    NpcMatados As Long
    Clase As Byte
    Raza As Byte
    Genero As Byte
End Type

Public Nombres As Boolean

'User status vars
Global OtroInventario(1 To MAX_INVENTORY_SLOTS) As Inventory

Public UserHechizos(1 To MAXHECHI) As Integer
Public UserInventory(1 To MAX_INVENTORY_SLOTS) As Inventory

Public NPCInventory(1 To MAX_NPC_INVENTORY_SLOTS) As NpCinV
Public UserMeditar As Boolean
Public UserName As String
Public UserAccount As String
Public UserAnswer As String
Public UserQuestion As Byte
Public UserPassword As String
Public UserMaxHP As Integer
Public UserPet As tFamiliar
Public UserMinHP As Integer
Public UserMaxMAN As Integer
Public UserMinMAN As Integer
Public UserMaxSTA As Integer
Public UserMinSTA As Integer
Public UserMaxAGU As Byte
Public UserMinAGU As Byte
Public UserMaxHAM As Byte
Public UserMinHAM As Byte
Public UserGLD As Long
Public UserLVL As Integer
Public UserPort As Integer
Public UserServerIP As String
Public UserEstado As Byte '0 = Vivo & 1 = Muerto
Public UserPasarNivel As Long
Public UserExp As Long
Public UserReputacion As tReputacion
Public UserEstadisticas As tEstadisticasUsu
Public UserDescansar As Boolean
Public tipf As String
Public PrimeraVez As Boolean
Public FPSFLAG As Boolean
Public Pausa As Boolean
Public IScombate As Boolean
Public UserParalizado As Boolean
Public UserNavegando As Boolean
Public UserHogar As eCiudad
Public UserMontando As Boolean
Public Comerciando As Boolean
Public UserClase As eClass
Public UserSexo As eGenero
Public UserRaza As eRaza
Public UserEmail As String

Public Const NUMCIUDADES As Byte = 1
Public Const NUMSKILLS As Integer = 27
Public Const NUMATRIBUTOS As Integer = 5
Public Const NUMCLASES As Integer = 18
Public Const NUMRAZAS As Integer = 6

Public UserSkills(1 To NUMSKILLS) As Byte
Public SkillsOrig(1 To NUMSKILLS) As Byte
Public SkillsNames(1 To NUMSKILLS) As String

Public UserAtributos(1 To NUMATRIBUTOS) As Byte
Public AtributosNames(1 To NUMATRIBUTOS) As String

Public Ciudades(1 To NUMCIUDADES) As String

Public ListaRazas(1 To NUMRAZAS) As String
Public ListaClases(1 To NUMCLASES) As String

Public SkillPoints As Integer
Public Alocados As Integer
Public flags() As Integer
Public Oscuridad As Integer
Public Logged As Boolean

Public UsingSkill As Integer

Public MD5HushYo As String * 16

Public pingTime As Long

Public Enum E_MODO
    Normal = 1
    CrearNuevoPj = 2
    Dados = 3
    CrearNuevaCuenta = 4
    ConectarCuenta = 5
End Enum

Public EstadoLogin As E_MODO
   
Public Enum FxMeditar
    CHICO = 4
    MEDIANO = 5
    GRANDE = 6
    XGRANDE = 16
    XXGRANDE = 34
End Enum

Public Enum eClanType
    ct_RoyalArmy
    ct_Evil
    ct_Neutral
    ct_GM
    ct_Legal
    ct_Criminal
End Enum

Public Enum eEditOptions
    eo_Gold = 1
    eo_Experience
    eo_Body
    eo_Head
    eo_CiticensKilled
    eo_CriminalsKilled
    eo_Level
    eo_Class
    eo_Skills
    eo_SkillPointsLeft
    eo_Nobleza
    eo_Asesino
    eo_Sex
    eo_Raza
    eo_Part
End Enum

''
' TRIGGERS
'
' @param NADA nada
' @param BAJOTECHO bajo techo
' @param trigger_2 ???
' @param POSINVALIDA los npcs no pueden pisar tiles con este trigger
' @param ZONASEGURA no se puede robar o pelear desde este trigger
' @param ANTIPIQUETE
' @param ZONAPELEA al pelear en este trigger no se caen las cosas y no cambia el estado de ciuda o crimi
'
Public Enum eTrigger
    NADA = 0
    BAJOTECHO = 1
    trigger_2 = 2
    POSINVALIDA = 3
    ZONASEGURA = 4
    ANTIPIQUETE = 5
    ZONAPELEA = 6
End Enum

'Server stuff
Public stxtbuffer As String 'Holds temp raw data from server
Public stxtbuffercmsg As String 'Holds temp raw data from server
Public Connected As Boolean 'True when connected to server
Public UserMap As Integer

'Control
Public prgRun As Boolean 'When true the program ends


'********** FUNCIONES API ***********
Public Declare Function GetTickCount Lib "kernel32" () As Long

'Teclado
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetActiveWindow Lib "user32" () As Long

'Para ejecutar el Internet Explorer para el manual
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public trueno As Byte
Public TalkMode As Byte

'Bordes del mapa
Public MinXBorder As Byte
Public MaxXBorder As Byte
Public MinYBorder As Byte
Public MaxYBorder As Byte

'********************************************
'*************Configuracion******************
'********************************************
Public Sound As Byte
Public Music As Byte
Public EffectSound As Byte
Public VolumeSound As Integer
Public VolumeMusic As Integer
Public TileBufferSize As Byte
Public cSombras As Byte
Public cTechos As Byte
Public cLimitarFps As Byte
Public cObjName As Byte
'********************************************
'*************/Configuracion*****************
'********************************************

'Particle Groups
Public TotalStreams As Integer
Public StreamData() As Stream

'RGB Type
Public Type RGB
    r As Long
    g As Long
    b As Long
End Type

Public Type Stream
    name As String
    NumOfParticles As Long
    NumGrhs As Long
    id As Long
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
    angle As Long
    vecx1 As Long
    vecx2 As Long
    vecy1 As Long
    vecy2 As Long
    life1 As Long
    life2 As Long
    friction As Long
    spin As Byte
    spin_speedL As Single
    spin_speedH As Single
    AlphaBlend As Byte
    gravity As Byte
    grav_strength As Long
    bounce_strength As Long
    XMove As Byte
    YMove As Byte
    move_x1 As Long
    move_x2 As Long
    move_y1 As Long
    move_y2 As Long
    grh_list() As Long
    colortint(0 To 3) As RGB
    
    speed As Single
    life_counter As Long
    
    grh_resize As Boolean
    grh_resizex As Integer
    grh_resizey As Integer
End Type

Public meteo_particle As Integer

'****************************************************************
'****************************************************************
'**********************SISTEMA DE CUENTAS************************
'****************************************************************
'****************************************************************
Public Type PjCuenta
    nombre      As String
    Head        As Integer
    body        As Integer
    Shield      As Byte
    Casco       As Byte
    Weapon      As Byte
    Nivel       As Byte
    Mapa        As Integer
    Clase       As Byte
    color       As Byte
End Type

Public cPJ(0 To 9) As PjCuenta
'****************************************************************
'****************************************************************
'****************************************************************
'****************************************************************
'****************************************************************


Type tServerInfo
    port As Integer
    Ip As String
    name As String
End Type
Public lServer(1 To 2) As tServerInfo

Public Type tCabecera 'Cabecera de los con
    Desc As String * 255
    CRC As Long
    MagicWord As Long
End Type

Public MiCabecera As tCabecera

'IIIIIIIICCCCCCCOOOOOOOOOONNNNNNNNNNOOOOOOOOOOSSSSSSSSSS
'To Put Grafical Cursors
Public Const GLC_HCURSOR = (-12)
Public hSwapCursor As Long
Public Declare Function LoadCursorFromFile Lib "user32" Alias "LoadCursorFromFileA" (ByVal lpfilename As String) As Long
Public Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Type tBoton
    TipoAccion As Integer
    SendString As String
    hlist As Integer
    invslot As Integer
End Type

Public MacroKeys() As tBoton
Public BotonElegido As Integer

Public base_light As Long
Public day_r_old As Byte
Public day_g_old As Byte
Public day_b_old As Byte

Type luzxhora
    r As Long
    g As Long
    b As Long
End Type

'Declaramos las luces
Public luz_dia(0 To 24) As luzxhora '¬¬ la hora 24 dura 1 minuto entre las 24 y las 0

'Vamos con las declaraciones del procesador :)
Public m_objCPUSet As SWbemObjectSet
Public m_objWMINameSpace As SWbemServices
Public oCpu As SWbemObject
'Fin............................................

'Ahora la resolucion
Public Const SW_Normal = 1

Type tListaFamiliares
    name As String
    Desc As String
    Imagen As String
End Type

Public ListaFamiliares() As tListaFamiliares

Type tFamiliar
    TieneFamiliar As Integer
    nombre As String
    ELV As Integer
    MinHP As Integer
    MaxHP As Integer
    ELU As Long
    EXP As Long
    MinHit As Integer
    MaxHit As Integer
    Abilidad As String
    TIPO As String
End Type

Public Enum HabilidadesFamiliar
    HABILIDAD_INMO = 1
    HABILIDAD_PARA = 2
    HABILIDAD_DESCARGA = 3
    HABILIDAD_TORMENTA = 4
    HABILIDAD_DESENCANTAR = 5
    HABILIDAD_CURAR = 6
    HABILIDAD_MISIL = 7
    HABILIDAD_DETECTAR = 8
    HABILIDAD_GOLPE_PARALIZA = 9
    HABILIDAD_GOLPE_ENTORPECE = 10
    HABILIDAD_GOLPE_DESARMA = 11
    HABILIDAD_GOLPE_ENCEGA = 12
    HABILIDAD_GOLPE_ENVENENA = 13
End Enum


Public Const iFragataFantasmal = 87
Public Const iFragataReal = 190
Public Const iFragataCaos = 189
Public Const iBarca = 84
Public Const iGalera = 85
Public Const iGaleon = 86
Public Const iBarcaCiuda = 84
Public Const iBarcaPk = 396
Public Const iGaleraCiuda = 397
Public Const iGaleraPk = 398
Public Const iGaleonCiuda = 399
Public Const iGaleonPk = 400
