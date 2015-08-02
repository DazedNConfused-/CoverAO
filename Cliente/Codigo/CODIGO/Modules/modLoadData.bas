Attribute VB_Name = "modLoadData"
Option Explicit
Public Const CANT_GRH_INDEX As Long = 40000
Public Sub LoadGrhData()
'*****************************************************************
'Loads Grh.dat
'*****************************************************************

On Error GoTo ErrorHandler

Dim Grh As Integer
Dim Frame As Integer
Dim tempInt As Integer
Dim f As Integer

ReDim GrhData(0 To CANT_GRH_INDEX) As GrhData

Extract_File Scripts, App.Path & "\Init", "graficos.ind", App.Path & "\Init\"

f = FreeFile()
Open App.Path & "\Init\Graficos.ind" For Binary Access Read As #f
    
    Seek #f, 1
    
    Get #f, , tempInt
    Get #f, , tempInt
    Get #f, , tempInt
    Get #f, , tempInt
    Get #f, , tempInt

    'Get first Grh Number
    Get #f, , Grh
    
    Do Until Grh <= 0
        'Get number of frames
        Get #f, , GrhData(Grh).NumFrames
        
        If GrhData(Grh).NumFrames <= 0 Then
            GoTo ErrorHandler
        End If
        
        ReDim GrhData(Grh).Frames(1 To GrhData(Grh).NumFrames)
        
        If GrhData(Grh).NumFrames > 1 Then
        
            'Read a animation GRH set
            For Frame = 1 To GrhData(Grh).NumFrames
                Get #f, , GrhData(Grh).Frames(Frame)
                If GrhData(Grh).Frames(Frame) <= 0 Or GrhData(Grh).Frames(Frame) > CANT_GRH_INDEX Then GoTo ErrorHandler
            Next Frame
        
            Get #f, , tempInt
            
            If tempInt <= 0 Then GoTo ErrorHandler
            GrhData(Grh).speed = GrhData(Grh).NumFrames * 0.018 'CLng(TempInt)
            
            'Compute width and height
            GrhData(Grh).pixelHeight = GrhData(GrhData(Grh).Frames(1)).pixelHeight
            If GrhData(Grh).pixelHeight <= 0 Then GoTo ErrorHandler
            
            GrhData(Grh).pixelWidth = GrhData(GrhData(Grh).Frames(1)).pixelWidth
            If GrhData(Grh).pixelWidth <= 0 Then GoTo ErrorHandler

            GrhData(Grh).TileWidth = GrhData(GrhData(Grh).Frames(1)).TileWidth
            If GrhData(Grh).TileWidth <= 0 Then GoTo ErrorHandler

            GrhData(Grh).TileHeight = GrhData(GrhData(Grh).Frames(1)).TileHeight
            If GrhData(Grh).TileHeight <= 0 Then GoTo ErrorHandler
        Else
            'Read in normal GRH data
            Get #f, , GrhData(Grh).FileNum
            If GrhData(Grh).FileNum <= 0 Then GoTo ErrorHandler

            Get #f, , GrhData(Grh).sX
            If GrhData(Grh).sX < 0 Then GoTo ErrorHandler
            
            Get #f, , GrhData(Grh).sY
            If GrhData(Grh).sY < 0 Then GoTo ErrorHandler

            Get #f, , GrhData(Grh).pixelWidth
            If GrhData(Grh).pixelWidth <= 0 Then GoTo ErrorHandler

            Get #f, , GrhData(Grh).pixelHeight
            If GrhData(Grh).pixelHeight <= 0 Then GoTo ErrorHandler

            'Compute width and height
            GrhData(Grh).TileWidth = GrhData(Grh).pixelWidth / 32
            GrhData(Grh).TileHeight = GrhData(Grh).pixelHeight / 32
            
            GrhData(Grh).Frames(1) = Grh
        End If
        'Get Next Grh Number
        Get #f, , Grh
    Loop
    
Close #f

Extract_File Scripts, App.Path & "\Init", "minimap.dat", App.Path & "\Init\"

Dim count As Long
f = FreeFile
Open App.Path & "\Init\minimap.dat" For Binary As #f
    Seek #1, 1
    For count = 1 To CANT_GRH_INDEX
        If Grh_Check(count) Then
            Get #f, , GrhData(count).mini_map_color
        End If
    Next count
Close #f

Exit Sub

ErrorHandler:
    Close #1
    MsgBox "Error al cargar el recurso de índice de gráficos: " & Err.Description & " (" & Grh & ")", vbCritical, "Error al cargar"

End Sub
Sub CargarCuerpos()
    Dim N As Integer
    Dim i As Long
    Dim NumCuerpos As Integer
    Dim MisCuerpos() As tIndiceCuerpo
    
    Extract_File Scripts, App.Path & "\Init", "personajes.ind", App.Path & "\Init\"
    
    N = FreeFile()
    Open App.Path & "\Init\personajes.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumCuerpos
    
    'Resize array
    ReDim BodyData(0 To NumCuerpos) As BodyData
    ReDim MisCuerpos(0 To NumCuerpos) As tIndiceCuerpo
    
    For i = 1 To NumCuerpos
        Get #N, , MisCuerpos(i)
        
        If MisCuerpos(i).body(1) Then
            InitGrh BodyData(i).Walk(1), MisCuerpos(i).body(1), 0
            InitGrh BodyData(i).Walk(2), MisCuerpos(i).body(2), 0
            InitGrh BodyData(i).Walk(3), MisCuerpos(i).body(3), 0
            InitGrh BodyData(i).Walk(4), MisCuerpos(i).body(4), 0
            
            BodyData(i).HeadOffset.x = MisCuerpos(i).HeadOffsetX
            BodyData(i).HeadOffset.y = MisCuerpos(i).HeadOffsetY
        End If
    Next i
    
    Close #N
    
End Sub
Sub CargarCabezas()
    Dim N As Integer
    Dim i As Long
    Dim Numheads As Integer
    Dim Miscabezas() As tIndiceCabeza
    
    
    Extract_File Scripts, App.Path & "\Init", "cabezas.ind", App.Path & "\Init\"
    N = FreeFile()
    Open App.Path & "\Init\cabezas.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , Numheads
    
    'Resize array
    ReDim HeadData(0 To Numheads) As HeadData
    ReDim Miscabezas(0 To Numheads) As tIndiceCabeza
    
    For i = 1 To Numheads
        Get #N, , Miscabezas(i)
        
        If Miscabezas(i).Head(1) Then
            Call InitGrh(HeadData(i).Head(1), Miscabezas(i).Head(1), 0)
            Call InitGrh(HeadData(i).Head(2), Miscabezas(i).Head(2), 0)
            Call InitGrh(HeadData(i).Head(3), Miscabezas(i).Head(3), 0)
            Call InitGrh(HeadData(i).Head(4), Miscabezas(i).Head(4), 0)
        End If
    Next i
    
    Close #N
    
End Sub
Sub CargarCascos()
    Dim N As Integer
    Dim i As Long
    Dim NumCascos As Integer

    Dim Miscabezas() As tIndiceCabeza
    
    Extract_File Scripts, App.Path & "\Init", "cascos.ind", App.Path & "\Init\"
    
    N = FreeFile()
    Open App.Path & "\Init\Cascos.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumCascos
    
    'Resize array
    ReDim CascoAnimData(0 To NumCascos) As HeadData
    ReDim Miscabezas(0 To NumCascos) As tIndiceCabeza
    
    For i = 1 To NumCascos
        Get #N, , Miscabezas(i)
        
        If Miscabezas(i).Head(1) Then
            Call InitGrh(CascoAnimData(i).Head(1), Miscabezas(i).Head(1), 0)
            Call InitGrh(CascoAnimData(i).Head(2), Miscabezas(i).Head(2), 0)
            Call InitGrh(CascoAnimData(i).Head(3), Miscabezas(i).Head(3), 0)
            Call InitGrh(CascoAnimData(i).Head(4), Miscabezas(i).Head(4), 0)
        End If
    Next i
    
    Close #N
    
    
End Sub
Sub CargarFxs()
    Dim N As Integer
    Dim i As Long
    Dim NumFxs As Integer
    
    Extract_File Scripts, App.Path & "\Init", "fxs.ind", App.Path & "\Init\"
    
    N = FreeFile()
    Open App.Path & "\Init\fxs.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumFxs
    
    'Resize array
    ReDim FxData(1 To NumFxs) As tIndiceFx
    
    For i = 1 To NumFxs
        Get #N, , FxData(i)
    Next i
    
    Close #N
     
End Sub
Sub CargarAnimArmas()
On Error Resume Next

    Dim loopc As Long
    Dim Leer As New clsIniReader
    
    Extract_File Scripts, App.Path & "\Init", "armas.dat", App.Path & "\Init\"
    
    Leer.Initialize App.Path & "\Init\armas.dat"
    
    NumWeaponAnims = Val(Leer.GetValue("INIT", "NumArmas"))
    
    ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData
    
    For loopc = 1 To NumWeaponAnims
        InitGrh WeaponAnimData(loopc).WeaponWalk(1), Val(Leer.GetValue("ARMA" & loopc, "Dir1")), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(2), Val(Leer.GetValue("ARMA" & loopc, "Dir2")), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(3), Val(Leer.GetValue("ARMA" & loopc, "Dir3")), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(4), Val(Leer.GetValue("ARMA" & loopc, "Dir4")), 0
    Next loopc
    
    Set Leer = Nothing
    
    
End Sub
Sub CargarAnimEscudos()
On Error Resume Next

    Dim loopc As Long
    Dim Leer As New clsIniReader

    Extract_File Scripts, App.Path & "\Init", "escudos.dat", App.Path & "\Init\"
    
    Leer.Initialize App.Path & "\Init\escudos.dat"
    
    NumEscudosAnims = Val(Leer.GetValue("INIT", "NumEscudos"))
    
    ReDim ShieldAnimData(1 To NumEscudosAnims) As ShieldAnimData
    
    For loopc = 1 To NumEscudosAnims
        InitGrh ShieldAnimData(loopc).ShieldWalk(1), Val(Leer.GetValue("ESC" & loopc, "Dir1")), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(2), Val(Leer.GetValue("ESC" & loopc, "Dir2")), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(3), Val(Leer.GetValue("ESC" & loopc, "Dir3")), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(4), Val(Leer.GetValue("ESC" & loopc, "Dir4")), 0
    Next loopc
    
    Set Leer = Nothing
    
    
End Sub

Public Sub CargarParticulas()
    Dim loopc As Long
    Dim i As Long
    Dim GrhListing As String
    Dim TempSet As String
    Dim ColorSet As Long
    Dim Leer As New clsIniReader

    Dim StreamFile As String

    Extract_File Scripts, App.Path & "\Init", "particulas.ini", App.Path & "\Init\"

    StreamFile = App.Path & "\Init\particulas.ini"
    
    Leer.Initialize StreamFile

    TotalStreams = Val(Leer.GetValue("INIT", "Total"))
    
    'resize StreamData array
    ReDim StreamData(1 To TotalStreams) As Stream
    
    'fill StreamData array with info from Particles.ini
    For loopc = 1 To TotalStreams
        StreamData(loopc).name = Leer.GetValue(Val(loopc), "Name")
        StreamData(loopc).NumOfParticles = Leer.GetValue(Val(loopc), "NumOfParticles")
        StreamData(loopc).x1 = Leer.GetValue(Val(loopc), "X1")
        StreamData(loopc).y1 = Leer.GetValue(Val(loopc), "Y1")
        StreamData(loopc).x2 = Leer.GetValue(Val(loopc), "X2")
        StreamData(loopc).y2 = Leer.GetValue(Val(loopc), "Y2")
        StreamData(loopc).angle = Leer.GetValue(Val(loopc), "Angle")
        StreamData(loopc).vecx1 = Leer.GetValue(Val(loopc), "VecX1")
        StreamData(loopc).vecx2 = Leer.GetValue(Val(loopc), "VecX2")
        StreamData(loopc).vecy1 = Leer.GetValue(Val(loopc), "VecY1")
        StreamData(loopc).vecy2 = Leer.GetValue(Val(loopc), "VecY2")
        StreamData(loopc).life1 = Leer.GetValue(Val(loopc), "Life1")
        StreamData(loopc).life2 = Leer.GetValue(Val(loopc), "Life2")
        StreamData(loopc).friction = Leer.GetValue(Val(loopc), "Friction")
        StreamData(loopc).spin = Leer.GetValue(Val(loopc), "Spin")
        StreamData(loopc).spin_speedL = Leer.GetValue(Val(loopc), "Spin_SpeedL")
        StreamData(loopc).spin_speedH = Leer.GetValue(Val(loopc), "Spin_SpeedH")
        StreamData(loopc).AlphaBlend = Leer.GetValue(Val(loopc), "AlphaBlend")
        StreamData(loopc).gravity = Leer.GetValue(Val(loopc), "Gravity")
        StreamData(loopc).grav_strength = Leer.GetValue(Val(loopc), "Grav_Strength")
        StreamData(loopc).bounce_strength = Leer.GetValue(Val(loopc), "Bounce_Strength")
        StreamData(loopc).XMove = Leer.GetValue(Val(loopc), "XMove")
        StreamData(loopc).YMove = Leer.GetValue(Val(loopc), "YMove")
        StreamData(loopc).move_x1 = Leer.GetValue(Val(loopc), "move_x1")
        StreamData(loopc).move_x2 = Leer.GetValue(Val(loopc), "move_x2")
        StreamData(loopc).move_y1 = Leer.GetValue(Val(loopc), "move_y1")
        StreamData(loopc).move_y2 = Leer.GetValue(Val(loopc), "move_y2")
        StreamData(loopc).life_counter = Leer.GetValue(Val(loopc), "life_counter")
        StreamData(loopc).speed = Val(Leer.GetValue(Val(loopc), "Speed"))
        
        Dim temp As Integer
        temp = Leer.GetValue(Val(loopc), "resize")
        
        StreamData(loopc).grh_resize = IIf((temp = -1), True, False)
        StreamData(loopc).grh_resizex = Leer.GetValue(Val(loopc), "rx")
        StreamData(loopc).grh_resizey = Leer.GetValue(Val(loopc), "ry")
        
        'Ai ya tenemos un problema de auras
        'tas? si qe paso, nesesito las auras de mi cumpu se pueden pasar por aca?
        ' cuanto pesan? nada osea es particles.ind dije auras jaja mira dale aca
    
        StreamData(loopc).NumGrhs = Leer.GetValue(Val(loopc), "NumGrhs")
        
        ReDim StreamData(loopc).grh_list(1 To StreamData(loopc).NumGrhs)
        GrhListing = Leer.GetValue(Val(loopc), "Grh_List")
        
        For i = 1 To StreamData(loopc).NumGrhs
            StreamData(loopc).grh_list(i) = Field_Read(str(i), GrhListing, ",")
        Next i
        
        StreamData(loopc).grh_list(i - 1) = StreamData(loopc).grh_list(i - 1)
        
        For ColorSet = 1 To 4
            TempSet = Leer.GetValue(Val(loopc), "ColorSet" & ColorSet)
            StreamData(loopc).colortint(ColorSet - 1).r = Field_Read(1, TempSet, ",")
            StreamData(loopc).colortint(ColorSet - 1).g = Field_Read(2, TempSet, ",")
            StreamData(loopc).colortint(ColorSet - 1).b = Field_Read(3, TempSet, ",")
        Next ColorSet

                
    Next loopc
    
    Set Leer = Nothing
    

End Sub
Private Function Grh_Check(ByVal grh_index As Long) As Boolean
    If grh_index > 0 And grh_index <= CANT_GRH_INDEX Then
        Grh_Check = GrhData(grh_index).NumFrames
    End If
End Function
Public Sub LoadMacros()
    Dim lc As Byte
    Dim Leer As New clsIniReader: Set Leer = New clsIniReader

    ReDim Preserve MacroKeys(1 To 11) As tBoton
    
    Leer.Initialize App.Path & "\Macros.dat"
    For lc = 1 To 11
        MacroKeys(lc).TipoAccion = Val(Leer.GetValue("Bind" & lc, "Accion"))
        MacroKeys(lc).hlist = Val(Leer.GetValue("Bind" & lc, "hlist"))
        MacroKeys(lc).invslot = Val(Leer.GetValue("Bind" & lc, "invslot"))
        MacroKeys(lc).SendString = Leer.GetValue("Bind" & lc, "SndString")
    Next lc
    Set Leer = Nothing


End Sub
