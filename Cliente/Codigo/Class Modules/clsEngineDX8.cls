VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTileEngineX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'******ENGINE****** Shermie80
Public bRunning As Boolean
Public FPS As Integer
Public Dia As String

Private Const FVF = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

Private FramesPerSecCounter As Integer
Private timerElapsedTime As Single
Private timerTicksPerFrame As Double
Private lFrameTimer As Long
Private lFrameLimiter As Long

Private ScrollPixelsPerFrameX As Byte
Private ScrollPixelsPerFrameY As Byte

Private HalfWindowTileWidth As Integer
Private HalfWindowTileHeight As Integer
'******ENGINE******

'Textos
Private Type tFont
    font_size As Integer
    ascii_code(0 To 255) As Integer 'indice de cada letra
End Type

'shermie80 (: - Mamá - tema a tema B tema T - jiji :X

Private font_types(1 To 2) As tFont
'Textos

'Luces

Private Type Light
    active As Boolean 'Do we ignore this light?
    id As Long
    map_x As Integer 'Coordinates
    map_y As Integer
    color As Long 'Start colour
    range As Byte
End Type

Dim light_list() As Light
Dim light_count As Long
Dim light_last As Long
'Luces

'******Particulas******
Private Type Particle
    friction As Single
    x As Single
    y As Single
    vector_x As Single
    vector_y As Single
    angle As Single
    Grh As Grh
    alive_counter As Long
    x1 As Integer
    x2 As Integer
    y1 As Integer
    y2 As Integer
    vecx1 As Integer
    vecx2 As Integer
    vecy1 As Integer
    vecy2 As Integer
    life1 As Long
    life2 As Long
    fric As Integer
    spin_speedL As Single
    spin_speedH As Single
    gravity As Boolean
    grav_strength As Long
    bounce_strength As Long
    spin As Boolean
    XMove As Boolean
    YMove As Boolean
    move_x1 As Integer
    move_x2 As Integer
    move_y1 As Integer
    move_y2 As Integer
    rgb_list(0 To 3) As Long
    grh_resize As Boolean
    grh_resizex As Integer
    grh_resizey As Integer
End Type

Private Type particle_group
    active As Boolean
    id As Long
    map_x As Integer
    map_y As Integer
    char_index As Long

    frame_counter As Single
    frame_speed As Single
    
    stream_type As Byte

    particle_stream() As Particle
    particle_count As Long
    
    grh_index_list() As Long
    grh_index_count As Long
    
    alpha_blend As Boolean
    
    alive_counter As Long
    never_die As Boolean
    
    live As Long
    liv1 As Integer
    liveend As Long
    
    x1 As Integer
    x2 As Integer
    y1 As Integer
    y2 As Integer
    angle As Integer
    vecx1 As Integer
    vecx2 As Integer
    vecy1 As Integer
    vecy2 As Integer
    life1 As Long
    life2 As Long
    fric As Long
    spin_speedL As Single
    spin_speedH As Single
    gravity As Boolean
    grav_strength As Long
    bounce_strength As Long
    spin As Boolean
    XMove As Boolean
    YMove As Boolean
    move_x1 As Integer
    move_x2 As Integer
    move_y1 As Integer
    move_y2 As Integer
    rgb_list(0 To 3) As Long
    
    'Added by Juan Martín Sotuyo Dodero
    speed As Single
    life_counter As Long
    
    'Added by David Justus
    grh_resize As Boolean
    grh_resizex As Integer
    grh_resizey As Integer
End Type
'Particle system
Dim particle_group_list() As particle_group
Dim particle_group_count As Long
Dim particle_group_last As Long
'******Particulas******



Private Function GetElapsedTime() As Single
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Gets the time that past since the last call
'**************************************************************
    Dim start_time As Currency
    Static end_time As Currency
    Static timer_freq As Currency

    'Get the timer frequency
    If timer_freq = 0 Then
        QueryPerformanceFrequency timer_freq
    End If
    
    'Get current time
    Call QueryPerformanceCounter(start_time)
    
    'Calculate elapsed time
    GetElapsedTime = (start_time - end_time) / timer_freq * 1000
    
    'Get next end time
    Call QueryPerformanceCounter(end_time)
End Function
Public Sub Init(ByVal eleccion As Byte)
    Dim DispMode As D3DDISPLAYMODE
    Dim DispModeBK As D3DDISPLAYMODE
    Dim D3DWindow As D3DPRESENT_PARAMETERS
    Dim SWidth As Long
    Dim SHeight As Long
    Dim TEXSIZE As Long
    
    Set SurfaceDB = New clsTexManager
    
    Set dX = New DirectX8
    Set D3D = dX.Direct3DCreate()
    Set D3DX = New D3DX8

    AmbientColor.r = 200
    AmbientColor.g = 200
    AmbientColor.b = 200
    AmbientColor.A = 255
    
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispModeBK
    
    With D3DWindow 'shermie80 - Si no sabes para que es esto Sacalo :D
        .Windowed = True
        .SwapEffect = D3DSWAPEFFECT_COPY
        .BackBufferFormat = DispMode.Format
        .BackBufferWidth = frmMain.MainViewPic.ScaleWidth
        .BackBufferHeight = frmMain.MainViewPic.ScaleHeight
        .EnableAutoDepthStencil = 1
        .AutoDepthStencilFormat = D3DFMT_D16
        .hDeviceWindow = frmMain.MainViewPic.hwnd
    End With
    
   Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.hwnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, _
                                                            D3DWindow)
    HalfWindowTileHeight = 6
    HalfWindowTileWidth = 8
    UserMap = 1
    
    D3DDevice.SetVertexShader FVF
    
    D3DDevice.SetRenderState D3DRS_LIGHTING, False
    
    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, True
    
    Call SurfaceDB.Init(D3DX, D3DDevice, General_Get_Free_Ram_Bytes)

    engineBaseSpeed = 0.018
    
    ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock

    UserPos.x = 50
    UserPos.y = 50
    
    Call Engine_Font_Initialize
    
    MinXBorder = XMinMapSize + (frmMain.MainViewPic.ScaleWidth / 64)
    MaxXBorder = XMaxMapSize - (frmMain.MainViewPic.ScaleWidth / 64)
    MinYBorder = YMinMapSize + (frmMain.MainViewPic.ScaleHeight / 64)
    MaxYBorder = YMaxMapSize - (frmMain.MainViewPic.ScaleHeight / 64)
    
    'Partículas
    D3DDevice.SetRenderState D3DRS_POINTSIZE, FToDW(2)
    D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
    D3DDevice.SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
    D3DDevice.SetRenderState D3DRS_POINTSCALE_ENABLE, 0
        
    bRunning = True

End Sub

Public Sub EndInit()
    Erase MapData
    Erase charlist
    Erase particle_group_list
    Erase font_types
    Set D3DDevice = Nothing
    Set D3D = Nothing
    Set dX = Nothing
End Sub

Public Sub Draw_GrhIndex(ByVal grh_index As Integer, ByVal x As Integer, ByVal y As Integer, Optional ByVal colour As Long = -1)
    If grh_index <= 0 Then Exit Sub
    Dim rgb_list(3) As Long
    
    rgb_list(0) = colour
    rgb_list(1) = colour
    rgb_list(2) = colour
    rgb_list(3) = colour
    
    Device_Box_Textured_Render grh_index, _
        x, y, _
        GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight, _
        rgb_list, _
        GrhData(grh_index).sX, GrhData(grh_index).sY
End Sub


Private Sub Draw_Grh(ByRef Grh As Grh, ByVal x As Integer, ByVal y As Integer, ByVal Center As Byte, ByVal Animate As Byte, ByRef color() As Long, Optional ByVal Alpha As Boolean, Optional ByVal map_x As Byte = 1, Optional ByVal map_y As Byte = 1, Optional ByVal angle As Single, Optional ByVal Shadow As Byte = 0)

    Dim CurrentGrhIndex As Integer
    
    If Grh.grhindex = 0 Then Exit Sub
    If GrhData(Grh.grhindex).NumFrames = 0 Then Exit Sub
    
    'If Animate Then
    '    If Grh.Started = 1 Then
    '        Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.grhindex).NumFrames / Grh.speed)
    '        If Grh.FrameCounter > GrhData(Grh.grhindex).NumFrames Then
    '            Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.grhindex).NumFrames) + 1
    '
    '            If Grh.Loops <> -1 Then
    '                If Grh.Loops > 0 Then
    '                    Grh.Loops = Grh.Loops - 1
    '                Else
    '                    Grh.Started = 0
    '                End If
    '            End If
    '        End If
    '    End If
    'End If
    
    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerTicksPerFrame * 0.5)
            If Grh.FrameCounter > GrhData(Grh.grhindex).NumFrames Then
                Grh.FrameCounter = 1 + Grh.FrameCounter Mod GrhData(Grh.grhindex).NumFrames
                If Grh.Loops <> -1 Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0
                    End If
                End If
            End If
        End If
    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.grhindex).Frames(Grh.FrameCounter)

    'Center Grh over X,Y pos
    If Center Then
        If GrhData(CurrentGrhIndex).TileWidth <> 1 Then
            x = x - Int(GrhData(CurrentGrhIndex).TileWidth * (32 \ 2)) + 32 \ 2
        End If

        If GrhData(Grh.grhindex).TileHeight <> 1 Then
            y = y - Int(GrhData(CurrentGrhIndex).TileHeight * 32) + 32
        End If
    End If

    Device_Box_Textured_Render CurrentGrhIndex, _
        x, y, _
        GrhData(CurrentGrhIndex).pixelWidth, GrhData(CurrentGrhIndex).pixelHeight, _
        color, _
        GrhData(CurrentGrhIndex).sX, GrhData(CurrentGrhIndex).sY, _
        Alpha, angle, Shadow

End Sub
Public Sub Render()

If UserCharIndex = 0 Then Exit Sub

D3DDevice.BeginScene

    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, 1#, 1#, 0
    
    RenderScreen
    
    If GetTickCount - lFrameTimer > 1000 Then
        FPS = FramesPerSecCounter
        FramesPerSecCounter = 0
        lFrameTimer = GetTickCount
    End If
    
D3DDevice.EndScene
D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0

lFrameLimiter = GetTickCount
FramesPerSecCounter = FramesPerSecCounter + 1
timerElapsedTime = GetElapsedTime()
If UserMontando = True Then
    timerTicksPerFrame = timerElapsedTime * (engineBaseSpeed + Val(Chr$(48) & Chr$(46) & Chr$(48) & Chr$(48) & Chr$(52)))
Else
    timerTicksPerFrame = timerElapsedTime * engineBaseSpeed
End If
End Sub

Sub RenderScreen()
    Dim y                   As Integer
    Dim x                   As Integer
    Dim screenminY          As Integer
    Dim screenmaxY          As Integer
    Dim screenminX          As Integer
    Dim screenmaxX          As Integer
    Dim minY                As Integer
    Dim maxY                As Integer
    Dim minX                As Integer
    Dim maxX                As Integer
    Dim ScreenX             As Integer
    Dim ScreenY             As Integer
    Dim minXOffset          As Integer
    Dim minYOffset          As Integer
    Dim PixelOffsetXTemp    As Integer
    Dim PixelOffsetYTemp    As Integer
    Dim CurrentGrhIndex     As Integer
    Dim offx                As Integer
    Dim offy                As Integer
    Dim tilex               As Integer
    Dim tiley               As Integer
    Dim PixelOffSetX        As Integer
    Dim PIxelOffSetY        As Integer
    Dim Techos(0 To 3)      As Long
    Static OffsetCounterX   As Single
    Static OffsetCounterY   As Single
    
    TileBufferSize = 9
    If UserMoving Then
        '****** Move screen Left and Right if needed ******
        If AddtoUserPos.x <> 0 Then
            OffsetCounterX = OffsetCounterX - 8 * AddtoUserPos.x * timerTicksPerFrame
            If Abs(OffsetCounterX) >= Abs(32 * AddtoUserPos.x) Then
                OffsetCounterX = 0
                AddtoUserPos.x = 0
                UserMoving = False
            End If
        End If
        '****** Move screen Up and Down if needed ******
        If AddtoUserPos.y <> 0 Then
            OffsetCounterY = OffsetCounterY - 8 * AddtoUserPos.y * timerTicksPerFrame
            If Abs(OffsetCounterY) >= Abs(32 * AddtoUserPos.y) Then
                OffsetCounterY = 0
                AddtoUserPos.y = 0
                UserMoving = False
            End If
        End If
    End If

    tilex = UserPos.x - AddtoUserPos.x
    tiley = UserPos.y - AddtoUserPos.y
    PixelOffSetX = OffsetCounterX
    PIxelOffSetY = OffsetCounterY
    
    'Figure out Ends and Starts of screen
    screenminY = tiley - HalfWindowTileHeight
    screenmaxY = tiley + HalfWindowTileHeight
    screenminX = tilex - HalfWindowTileWidth
    screenmaxX = tilex + HalfWindowTileWidth
    
    minY = screenminY - TileBufferSize
    maxY = screenmaxY + TileBufferSize
    minX = screenminX - TileBufferSize
    maxX = screenmaxX + TileBufferSize
    
    'Make sure mins and maxs are allways in map bounds
    If minY < XMinMapSize Then
        minYOffset = YMinMapSize - minY
        minY = YMinMapSize
    End If
    
    If maxY > YMaxMapSize Then maxY = YMaxMapSize
    
    If minX < XMinMapSize Then
        minXOffset = XMinMapSize - minX
        minX = XMinMapSize
    End If
    
    If maxX > XMaxMapSize Then maxX = XMaxMapSize
    
    'If we can, we render around the view area to make it smoother
    If screenminY > YMinMapSize Then
        screenminY = screenminY - 1
    Else
        screenminY = 1
        ScreenY = 1
    End If
    
    If screenmaxY < YMaxMapSize Then screenmaxY = screenmaxY + 1
    
    If screenminX > XMinMapSize Then
        screenminX = screenminX - 1
    Else
        screenminX = 1
        ScreenX = 1
    End If
    
    If screenmaxX < XMaxMapSize Then screenmaxX = screenmaxX + 1
    
    RoofAlphaCalculate Techos
    
    'Draw floor layer
    For y = screenminY To screenmaxY
        For x = screenminX To screenmaxX
            If InMapBounds(x, y) Then
                'Layer 1 **********************************
                If MapData(x, y).Graphic(1).grhindex <> 0 Then
                    Call Draw_Grh(MapData(x, y).Graphic(1), (ScreenX - 1) * 32 + PixelOffSetX, (ScreenY - 1) * 32 + PIxelOffSetY, 0, 1, MapData(x, y).light_value, , x, y, , 5)
                End If
                '******************************************
                'Layer 2 **********************************
                If MapData(x, y).Graphic(2).grhindex <> 0 Then
                    Call Draw_Grh(MapData(x, y).Graphic(2), (ScreenX - 1) * 32 + PixelOffSetX, (ScreenY - 1) * 32 + PIxelOffSetY, 1, 1, MapData(x, y).light_value, , x, y)
                End If
                '******************************************
            End If
            ScreenX = ScreenX + 1
        Next x

        'Reset ScreenX to original value and increment ScreenY
        ScreenX = ScreenX - x + screenminX
        ScreenY = ScreenY + 1
    Next y

    ScreenY = minYOffset - TileBufferSize
    For y = minY To maxY
        ScreenX = minXOffset - TileBufferSize
        For x = minX To maxX
            PixelOffsetXTemp = ScreenX * 32 + PixelOffSetX
            PixelOffsetYTemp = ScreenY * 32 + PIxelOffSetY
            With MapData(x, y)
                If InMapBounds(x, y) Then
                    '******************************************
                    'Object Layer **********************************
                    If .ObjGrh.grhindex <> 0 Then
                        Call Draw_Grh(.ObjGrh, PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, MapData(x, y).light_value, , x, y)
                    End If
                        
                    '***********************************************
                    'Layer 3 *****************************************
                    If .Graphic(3).grhindex <> 0 Then
                        Call Draw_Grh(.Graphic(3), PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, MapData(x, y).light_value, , x, y)
                    End If
                    '************************************************
                    'Char layer ************************************
                    If .CharIndex <> 0 Then
                        If charlist(.CharIndex).Pos.x = x And charlist(.CharIndex).Pos.y = y Then
                            Call Char_Render(.CharIndex, PixelOffsetXTemp, PixelOffsetYTemp, x, y)
                        Else
                            .CharIndex = 0
                        End If
                    End If
                    '*************************************************
                End If
            End With
            ScreenX = ScreenX + 1
        Next x
        ScreenY = ScreenY + 1
    Next y
    ScreenY = minYOffset - 5
    
    Call Dialogos.Render
    
    ScreenY = minYOffset - TileBufferSize
    For y = minY To maxY
        ScreenX = minXOffset - TileBufferSize
        For x = minX To maxX
            PixelOffsetXTemp = ScreenX * 32 + PixelOffSetX
            PixelOffsetYTemp = ScreenY * 32 + PIxelOffSetY
            'Particles*****************************************
            If InMapBounds(x, y) Then
                If MapData(x, y).particle_group_index > 0 Then
                    Particle_Group_Render MapData(x, y).particle_group_index, ScreenX * 32 + PixelOffSetX + 16, ScreenY * 32 + PIxelOffSetY + 16
                End If
            End If
            '**************************************************
            ScreenX = ScreenX + 1
        Next x
        ScreenY = ScreenY + 1
    Next y
    ScreenY = minYOffset - 5
    
    
    'Draw blocked tiles and grid
    ScreenY = minYOffset - TileBufferSize
    For y = minY To maxY
        ScreenX = minXOffset - TileBufferSize
        For x = minX To maxX
            If InMapBounds(x, y) Then
                'Layer 4 **********************************
                If MapData(x, y).Graphic(4).grhindex Then
                    Call Draw_Grh(MapData(x, y).Graphic(4), _
                        ScreenX * 32 + PixelOffSetX, _
                        ScreenY * 32 + PIxelOffSetY, _
                        1, 1, Techos, , x, y)
                End If
                '**********************************
            End If
            ScreenX = ScreenX + 1
        Next x
        ScreenY = ScreenY + 1
    Next y
    
    If meteo_particle Then Particle_Group_Render meteo_particle, 250, 0
    
    If trueno And Not trueno = 1 Then
        Trueno_Render
        trueno = trueno - 1
    ElseIf trueno = 1 Then
        Call Audio.PlayWave("105.wav", RandomNumber(1, 100), RandomNumber(1, 100))
        Trueno_Render
        trueno = 0
    End If
    
    Text_Render "FPS:" & FPS, 10, 10, Default_RGB
    
    Call DialogosClanes.Draw
    Call DibujarCartel
    
    day_r_old = luz_dia(Hour(time)).r
    day_g_old = luz_dia(Hour(time)).g
    day_b_old = luz_dia(Hour(time)).b
    base_light = ARGB(day_r_old, day_g_old, day_b_old, 255)

   If LuzMouse Then
            Engine.Light_Move (Engine.Light_Find(20)), UserPos.x + frmMain.MouseX \ 32 - frmMain.MainViewPic.ScaleWidth \ 64, UserPos.y + frmMain.MouseY / 32 - frmMain.MainViewPic.ScaleHeight \ 64
        End If
        Engine.Light_Render_All

End Sub

Private Function Geometry_Create_TLVertex(ByVal x As Single, ByVal y As Single, ByVal Z As Single, _
                                            ByVal rhw As Single, ByVal color As Long, ByVal Specular As Long, tu As Single, _
                                            ByVal tv As Single) As TLVERTEX
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'**************************************************************
    Geometry_Create_TLVertex.x = x
    Geometry_Create_TLVertex.y = y
    Geometry_Create_TLVertex.Z = Z
    Geometry_Create_TLVertex.rhw = rhw
    Geometry_Create_TLVertex.color = color
    Geometry_Create_TLVertex.Specular = Specular
    Geometry_Create_TLVertex.tu = tu
    Geometry_Create_TLVertex.tv = tv
End Function

Private Sub Geometry_Create_Box(ByRef verts() As TLVERTEX, ByRef dest As RECT, ByRef src As RECT, ByRef rgb_list() As Long, _
                                Optional ByRef Textures_Width As Long, Optional ByRef Textures_Height As Long, Optional ByVal angle As Single)
'**************************************************************
'Author: Aaron Perkins
'Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 11/17/2002
'
' * v1      * v3
' |\        |
' |  \      |
' |    \    |
' |      \  |
' |        \|
' * v0      * v2
'**************************************************************
    Dim x_center As Single
    Dim y_center As Single
    Dim radius As Single
    Dim x_Cor As Single
    Dim y_Cor As Single
    Dim left_point As Single
    Dim right_point As Single
    Dim temp As Single
    
    If angle > 0 Then
        'Center coordinates on screen of the square
        x_center = dest.Left + (dest.Right - dest.Left) / 2
        y_center = dest.Top + (dest.bottom - dest.Top) / 2
        
        'Calculate radius
        radius = Sqr((dest.Right - x_center) ^ 2 + (dest.bottom - y_center) ^ 2)
        
        'Calculate left and right points
        temp = (dest.Right - x_center) / radius
        right_point = Atn(temp / Sqr(-temp * temp + 1))
        left_point = PI - right_point
    End If
    
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Left
        y_Cor = dest.bottom
    Else
        x_Cor = x_center + Cos(-left_point - angle) * radius
        y_Cor = y_center - Sin(-left_point - angle) * radius
    End If
    
    
    '0 - Bottom left vertex
    If Textures_Width And Textures_Height Then
        verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, src.Left / Textures_Width, (src.bottom + 1) / Textures_Height)
    Else
        verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, 0, 0)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Left
        y_Cor = dest.Top
    Else
        x_Cor = x_center + Cos(left_point - angle) * radius
        y_Cor = y_center - Sin(left_point - angle) * radius
    End If
    
    
    '1 - Top left vertex
    If Textures_Width And Textures_Height Then
        verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, src.Left / Textures_Width, src.Top / Textures_Height)
    Else
        verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, 0, 1)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Right
        y_Cor = dest.bottom
    Else
        x_Cor = x_center + Cos(-right_point - angle) * radius
        y_Cor = y_center - Sin(-right_point - angle) * radius
    End If
    
    
    '2 - Bottom right vertex
    If Textures_Width And Textures_Height Then
        verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, (src.Right + 1) / Textures_Width, (src.bottom + 1) / Textures_Height)
    Else
        verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, 1, 0)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Right
        y_Cor = dest.Top
    Else
        x_Cor = x_center + Cos(right_point - angle) * radius
        y_Cor = y_center - Sin(right_point - angle) * radius
    End If
    
    
    '3 - Top right vertex
    If Textures_Width And Textures_Height Then
        verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, (src.Right + 1) / Textures_Width, src.Top / Textures_Height)
    Else
        verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, 1, 1)
    End If

End Sub
Public Sub Device_Box_Textured_Render_Advance(ByVal grhindex As Long, ByVal dest_x As Integer, ByVal dest_y As Integer, ByVal src_width As Integer, _
                                            ByVal src_height As Integer, ByRef rgb_list() As Long, ByVal src_x As Integer, _
                                            ByVal src_y As Integer, ByVal dest_width As Integer, Optional ByVal dest_height As Integer, _
                                            Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/15/2003
'Copies the Textures allowing resizing
'Modified by Juan Martín Sotuyo Dodero
'**************************************************************
    Static src_rect As RECT
    Static dest_rect As RECT
    Static temp_verts(3) As TLVERTEX
    Static d3dTextures As D3D8Textures
    Static light_value(0 To 3) As Long

    If grhindex = 0 Then Exit Sub
    
    If GrhData(grhindex).FileNum = 0 Then
        Debug.Print grhindex
    End If
    
    Set d3dTextures.Texture = SurfaceDB.GetTexture(GrhData(grhindex).FileNum, d3dTextures.texwidth, d3dTextures.texheight)
    
    light_value(0) = rgb_list(0)
    light_value(1) = rgb_list(1)
    light_value(2) = rgb_list(2)
    light_value(3) = rgb_list(3)

    light_value(0) = &HFFFFFFFF 'blind_color
    light_value(1) = &HFFFFFFFF 'blind_color
    light_value(2) = &HFFFFFFFF 'blind_color
    light_value(3) = &HFFFFFFFF 'blind_color
    
    'Set up the source rectangle
    With src_rect
        .bottom = src_y + src_height
        .Left = src_x
        .Right = src_x + src_width
        .Top = src_y
    End With
        
    'Set up the destination rectangle
    With dest_rect
        .bottom = dest_y + dest_height
        .Left = dest_x
        .Right = dest_x + dest_width
        .Top = dest_y
    End With
    
    'Set up the TempVerts(3) vertices
    Geometry_Create_Box temp_verts(), dest_rect, src_rect, light_value(), d3dTextures.texwidth, d3dTextures.texheight, angle
        
    'Set Textures
    D3DDevice.SetTexture 0, d3dTextures.Texture
    
    If alpha_blend Then
       'Set Rendering for alphablending
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    End If
    
    'Draw the triangles that make up our square Textures
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0))
    
    If alpha_blend Then
        'Set Rendering for colokeying
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    End If

End Sub

Public Sub Device_Box_Textured_Render(ByVal grhindex As Long, ByVal dest_x As Integer, ByVal dest_y As Integer, ByVal src_width As Integer, _
                                            ByVal src_height As Integer, ByRef rgb_list() As Long, ByVal src_x As Integer, _
                                            ByVal src_y As Integer, Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single, _
                                            Optional ByVal Shadow As Byte = 0)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 2/12/2004
'Just copies the Textures
'**************************************************************
    Static src_rect As RECT
    Static dest_rect As RECT
    Static temp_verts(3) As TLVERTEX
    Static d3dTextures As D3D8Textures
    Static light_value(0 To 3) As Long
    
    If grhindex = 0 Then Exit Sub
    Set d3dTextures.Texture = SurfaceDB.GetTexture(GrhData(grhindex).FileNum, d3dTextures.texwidth, d3dTextures.texheight)
    
    light_value(0) = rgb_list(0)
    light_value(1) = rgb_list(1)
    light_value(2) = rgb_list(2)
    light_value(3) = rgb_list(3)

    If (light_value(0) = 0) Then light_value(0) = map_base_light
    If (light_value(1) = 0) Then light_value(1) = map_base_light
    If (light_value(2) = 0) Then light_value(2) = map_base_light
    If (light_value(3) = 0) Then light_value(3) = map_base_light
    
    'Set up the source rectangle
    With src_rect
        .bottom = src_y + src_height
        .Left = src_x
        .Right = src_x + src_width
        .Top = src_y
    End With
                
    'Set up the destination rectangle
    With dest_rect
        .bottom = dest_y + src_height
        .Left = dest_x
        .Right = dest_x + src_width
        .Top = dest_y
    End With
    
    'Set up the TempVerts(3) vertices
    Geometry_Create_Box temp_verts(), dest_rect, src_rect, light_value(), d3dTextures.texwidth, d3dTextures.texheight, angle

    'Set Textures
    D3DDevice.SetTexture 0, d3dTextures.Texture

    If alpha_blend Then
       'Set Rendering for alphablending
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    End If

    'Draw the triangles that make up our square Textures
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0))
    
    If alpha_blend Then
        'Set Rendering for colokeying
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    End If

End Sub


Public Sub MoveScreen(ByVal nHeading As E_Heading)
'******************************************
'Starts the screen moving in a direction
'******************************************
    Dim x As Integer
    Dim y As Integer
    Dim tX As Integer
    Dim tY As Integer
    
    'Figure out which way to move
    Select Case nHeading
        Case E_Heading.NORTH
            y = -1
        
        Case E_Heading.EAST
            x = 1
        
        Case E_Heading.SOUTH
            y = 1
        
        Case E_Heading.WEST
            x = -1
    End Select
    
    'Fill temp pos
    tX = UserPos.x + x
    tY = UserPos.y + y
    
    'Check to see if its out of bounds
    If tX < XMinMapSize Or tX > XMaxMapSize Or tY < YMinMapSize Or tY > YMaxMapSize Then
        Exit Sub
    Else
        'Start moving... MainLoop does the rest
        AddtoUserPos.x = x
        UserPos.x = tX
        AddtoUserPos.y = y
        UserPos.y = tY
        UserMoving = 1
        
        bTecho = IIf(MapData(UserPos.x, UserPos.y).Trigger = 1 Or _
                MapData(UserPos.x, UserPos.y).Trigger = 2 Or _
                MapData(UserPos.x, UserPos.y).Trigger = 4 Or _
                MapData(UserPos.x, UserPos.y).Trigger >= 20, True, False)
    End If
End Sub

Private Sub Char_Render(ByVal CharIndex As Long, ByVal PixelOffSetX As Integer, ByVal PIxelOffSetY As Integer, ByVal x As Byte, ByVal y As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/03/04
'Draw char's to screen without offcentering them
'***************************************************
    Dim moved As Boolean
    Dim Pos As Integer
    Dim line As String
    Dim color As Long
    Dim rgb_list(0 To 3) As Long
    
    With charlist(CharIndex)
        If .Moving Then
            'If needed, move left and right
            If .scrollDirectionX <> 0 Then
                .MoveOffsetX = .MoveOffsetX + 8 * Sgn(.scrollDirectionX) * timerTicksPerFrame
                
                'Start animations
                If .body.Walk(.heading).speed > 0 Then _
                    .body.Walk(.heading).Started = 1
                .Arma.WeaponWalk(.heading).Started = 1
                .Escudo.ShieldWalk(.heading).Started = 1
                
                'Char moved
                moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionX) = 1 And .MoveOffsetX >= 0) Or _
                        (Sgn(.scrollDirectionX) = -1 And .MoveOffsetX <= 0) Then
                    .MoveOffsetX = 0
                    .scrollDirectionX = 0
                End If
            End If
            
            'If needed, move up and down
            If .scrollDirectionY <> 0 Then
                .MoveOffsetY = .MoveOffsetY + 8 * Sgn(.scrollDirectionY) * timerTicksPerFrame
                
                If .body.Walk(.heading).speed > 0 Then _
                    .body.Walk(.heading).Started = 1
                .Arma.WeaponWalk(.heading).Started = 1
                .Escudo.ShieldWalk(.heading).Started = 1
                
                'Char moved
                moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionY) = 1 And .MoveOffsetY >= 0) Or _
                        (Sgn(.scrollDirectionY) = -1 And .MoveOffsetY <= 0) Then
                    .MoveOffsetY = 0
                    .scrollDirectionY = 0
                End If
            End If
        End If
        
        If .heading = 0 Then .heading = 1
        
        'If done moving stop animation
        If Not moved Then
            'Stop animations
            .body.Walk(.heading).Started = 0
            .body.Walk(.heading).FrameCounter = 1
            
            .Arma.WeaponWalk(.heading).Started = 0
            .Arma.WeaponWalk(.heading).FrameCounter = 1
            
            .Escudo.ShieldWalk(.heading).Started = 0
            .Escudo.ShieldWalk(.heading).FrameCounter = 1
            
            .Moving = False
        End If
        
        PixelOffSetX = PixelOffSetX + .MoveOffsetX
        PIxelOffSetY = PIxelOffSetY + .MoveOffsetY

       If .iHead = 0 And .iBody > 0 Then
        If Len(.nombre) > 0 Then
        line = "< " & .nombre & " >"
        Call Grh_Render(.plusGrh(2).Grh, PixelOffSetX, PIxelOffSetY + 30, rgb_list, , , True)
        End If
        End If

        If .plusGrh(2).Grh.grhindex <> 0 Then
            Call Long_To_RGB_List(rgb_list, .plusGrh(2).color)
            Call Grh_Render(.plusGrh(2).Grh, PixelOffSetX, PIxelOffSetY + 32, rgb_list, , , True)
        End If
        
        If .plusGrh(0).Grh.grhindex <> 0 Then
            Call Long_To_RGB_List(rgb_list, .plusGrh(0).color)
            Call Grh_Render(.plusGrh(0).Grh, PixelOffSetX, PIxelOffSetY + 32, rgb_list, , , True)
        End If
        
        If .plusGrh(1).Grh.grhindex <> 0 Then
            Call Long_To_RGB_List(rgb_list, .plusGrh(1).color)
            Call Grh_Render(.plusGrh(1).Grh, PixelOffSetX, PIxelOffSetY + 32, rgb_list, , , True)
        End If
            
        If Not .invisible And Not .Muerto Then
            .AlphaX = 0
            rgb_list(0) = MapData(x, y).light_value(0)
            rgb_list(1) = MapData(x, y).light_value(1)
            rgb_list(2) = MapData(x, y).light_value(2)
            rgb_list(3) = MapData(x, y).light_value(3)
        Else
            Call RoofAlphaCalculateToAlpha(rgb_list(), CharIndex)
        End If
        
        If .heading = EAST Then
            If .Escudo.ShieldWalk(.heading).grhindex And Not .iBody = 84 Then
                Call Draw_Grh(.Escudo.ShieldWalk(.heading), PixelOffSetX, PIxelOffSetY - .ShieldOffSetY, 1, 1, rgb_list(), , x, y)
            End If
        End If
        
        If .body.Walk(.heading).grhindex Then
            Call Draw_Grh(.body.Walk(.heading), PixelOffSetX, PIxelOffSetY, 1, 1, rgb_list(), , x, y)
        End If

        If .Head.Head(.heading).grhindex Then
            Call Draw_Grh(.Head.Head(.heading), PixelOffSetX + .body.HeadOffset.x, PIxelOffSetY + .body.HeadOffset.y, 1, 0, rgb_list(), , x, y)
        End If
                
        If Not .iBody = 84 Then
            If .Arma.WeaponWalk(.heading).grhindex And isMontura(.iBody) Then
                Call Draw_Grh(.Arma.WeaponWalk(.heading), PixelOffSetX + .body.HeadOffset.x, PIxelOffSetY + .body.HeadOffset.y + 36, 1, 1, rgb_list(), , x, y)
            End If
                
            If .heading <> EAST Then
                If .Escudo.ShieldWalk(.heading).grhindex Then
                    Call Draw_Grh(.Escudo.ShieldWalk(.heading), PixelOffSetX, PIxelOffSetY - .ShieldOffSetY + 3, 1, 1, rgb_list(), , x, y)
                End If
            End If
    
            If .Casco.Head(.heading).grhindex Then
                Call Draw_Grh(.Casco.Head(.heading), PixelOffSetX + .body.HeadOffset.x, PIxelOffSetY + .body.HeadOffset.y, 1, 0, rgb_list(), , x, y)
            End If
        End If
                            
        'Saque el sistema de Donador porque Soy malito jiji xD - Shermie80 (:
                            
        'Draw name over head
        If Nombres And .invisible = False Then
            If Len(.nombre) > 0 Then
                Pos = InStr(.nombre, "<")
                If Pos = 0 Then Pos = Len(.nombre) + 2

                'Nick
                line = Left$(.nombre, Pos - 2)
                Text_Render line, PixelOffSetX - (textwidth(line, 1) / 2) + 16, PIxelOffSetY + 30, .label_color
                    
                'Clan
                If .Priv <> 5 Then
                    line = mid$(.nombre, Pos)
                    Text_Render line, PixelOffSetX - (textwidth(line, 1) / 2) + 16, PIxelOffSetY + 45, .label_color
                Else 'Staff
                    line = "<CoverAO Staff>" 'Shermie80
                    Text_Render line, PixelOffSetX - (textwidth(line, 1) / 2) + 16, PIxelOffSetY + 45, .label_color
                End If
                
                
            End If
        End If
        
        ''Update dialogs
        Call Dialogos.UpdateDialogPos(PixelOffSetX + .body.HeadOffset.x, PIxelOffSetY + .body.HeadOffset.y, CharIndex)
        
        Dim i As Integer
        If charlist(CharIndex).particle_count > 0 Then
            For i = 1 To UBound(charlist(CharIndex).particle_group)
                If charlist(CharIndex).particle_group(i) > 0 Then
                    Particle_Group_Render charlist(CharIndex).particle_group(i), PixelOffSetX + 16, PIxelOffSetY
                End If
            Next i
        End If
        
        
        'Draw FX
        If .fxIndex <> 0 Then
            Call Draw_Grh(.fX, PixelOffSetX + FxData(.fxIndex).OffsetX, PIxelOffSetY + FxData(.fxIndex).OffsetY, 1, 1, MapData(x, y).light_value, True)

            If .fX.Started = 0 Then
                .fxIndex = 0
            End If
            
        End If
    End With
End Sub
Function isMontura(ByVal body As Integer) As Boolean
    isMontura = (body <> 416 And body <> 415 And _
                 body <> 412 And body <> 381 And _
                 body <> 383 And body <> 384 And _
                 body <> 382 And body <> 413 And _
                 body <> 292 And body <> 291 And _
                 body <> 317)
End Function
Public Sub SetCharacterFx(ByVal CharIndex As Integer, ByVal fX As Integer, ByVal Loops As Integer)

    With charlist(CharIndex)
        .fxIndex = fX
        
        If .fxIndex > 0 Then
            Call InitGrh(.fX, FxData(fX).Animacion)
        
            .fX.Loops = Loops
        End If
    End With
    
End Sub

Public Sub Char_Move_By_Head(ByVal CharIndex As Integer, ByVal nHeading As E_Heading)
'*****************************************************************
'Starts the movement of a character in nHeading direction
'*****************************************************************
    Dim addx As Integer
    Dim addy As Integer
    Dim x As Integer
    Dim y As Integer
    Dim nX As Integer
    Dim nY As Integer
    
    With charlist(CharIndex)
        x = .Pos.x
        y = .Pos.y
        
        'Figure out which way to move
        Select Case nHeading
            Case E_Heading.NORTH
                addy = -1
        
            Case E_Heading.EAST
                addx = 1
        
            Case E_Heading.SOUTH
                addy = 1
            
            Case E_Heading.WEST
                addx = -1
        End Select
        
        nX = x + addx
        nY = y + addy
        
        MapData(nX, nY).CharIndex = CharIndex
        .Pos.x = nX
        .Pos.y = nY
        MapData(x, y).CharIndex = 0
        
        .MoveOffsetX = -1 * (32 * addx)
        .MoveOffsetY = -1 * (32 * addy)
        
        .Moving = 1
        .heading = nHeading
        
        .scrollDirectionX = addx
        .scrollDirectionY = addy
    End With
    
    If UserEstado <> 1 Then Call DoPasosFx(CharIndex)
    
    'areas viejos
    If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
        Call EraseChar(CharIndex)
    End If
End Sub
Public Sub Char_Clean()
Dim x As Byte
Dim y As Byte
For x = 1 To 100
    For y = 1 To 100
        If MapData(x, y).CharIndex Then
            EraseChar MapData(x, y).CharIndex
        End If
        If MapData(x, y).ObjGrh.grhindex Then
            MapData(x, y).ObjGrh.grhindex = 0
        End If
    Next y
Next x
End Sub
Public Sub Char_Move_By_Pos(ByVal CharIndex As Integer, ByVal nX As Integer, ByVal nY As Integer)
On Error Resume Next
    Dim x As Integer
    Dim y As Integer
    Dim addx As Integer
    Dim addy As Integer
    Dim nHeading As E_Heading
    
    With charlist(CharIndex)
        x = .Pos.x
        y = .Pos.y
        
        MapData(x, y).CharIndex = 0
        
        addx = nX - x
        addy = nY - y
        
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
        
        
        .Pos.x = nX
        .Pos.y = nY
        
        .MoveOffsetX = -1 * (32 * addx)
        .MoveOffsetY = -1 * (32 * addy)
        
        .Moving = 1
        .heading = nHeading
        
        .scrollDirectionX = Sgn(addx)
        .scrollDirectionY = Sgn(addy)
        
        'parche para que no medite cuando camina
        If .fxIndex = FxMeditar.CHICO Or .fxIndex = FxMeditar.GRANDE Or .fxIndex = FxMeditar.MEDIANO Or .fxIndex = FxMeditar.XGRANDE Or .fxIndex = FxMeditar.XXGRANDE Then
            .fxIndex = 0
        End If
    End With
    
End Sub

Private Function FToDW(f As Single) As Long
' single > long
Dim buf As D3DXBuffer
    Set buf = D3DX.CreateBuffer(4)
    D3DX.BufferSetData buf, 0, 4, 1, f
    D3DX.BufferGetData buf, 0, 4, 1, FToDW
End Function

Public Sub DrawInv()
    Static re As RECT
    re.Left = 0
    re.Top = 0
    re.bottom = 160
    re.Right = 160
    
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 0, 0
    D3DDevice.BeginScene
        Inventario.DrawInventory
    D3DDevice.EndScene
    D3DDevice.Present re, ByVal 0, frmMain.picInv.hwnd, ByVal 0
End Sub

Public Sub DrawGrhToHdcToFile(ByVal desthDC As Long, ByVal screen_x As Integer, ByVal screen_y As Integer, _
ByVal src_x As Integer, _
ByVal src_y As Integer, _
ByVal src_width As Integer, _
ByVal src_height As Integer, _
ByVal file_path As String)
 
On Error GoTo Err
 
Dim hdcsrc As Long
Dim MaskDC As Long
Dim PrevObj As Long
Dim PrevObj2 As Long
 
hdcsrc = CreateCompatibleDC(desthDC)
 
 PrevObj = SelectObject(hdcsrc, LoadPicture(file_path))
 
BitBlt desthDC, screen_x, screen_y, src_width, src_height, hdcsrc, src_x, src_y, vbSrcCopy
TransparentBlt desthDC, screen_x, screen_y, src_width, src_height, hdcsrc, src_x, src_y, src_width, src_height, RGB(0, 0, 0)
 
Call DeleteObject(SelectObject(hdcsrc, PrevObj))
DeleteDC hdcsrc
 
Err:
If Err.number = 481 Then MsgBox "Imposible cargar recurso error: " & "481"
 
End Sub
Public Sub DrawGrhToHdc(ByVal desthDC As Long, Grh As Integer, ByVal screen_x As Integer, ByVal screen_y As Integer)
 
On Error GoTo Err
 
    Dim file_path As String
    Dim src_x As Integer
    Dim src_y As Integer
    Dim src_width As Integer
    Dim src_height As Integer
    Dim hdcsrc As Long
    Dim MaskDC As Long
    Dim PrevObj As Long
    Dim PrevObj2 As Long
    Dim grh_index As Integer
   
    grh_index = Grh
 
    If grh_index <= 0 Then Exit Sub
    If GrhData(grh_index).NumFrames = 0 Then Exit Sub
 
    If GrhData(grh_index).NumFrames <> 1 Then
        grh_index = GrhData(grh_index).Frames(1)
    End If
 
    src_x = GrhData(grh_index).sX
    src_y = GrhData(grh_index).sY
    src_width = GrhData(grh_index).pixelWidth
    src_height = GrhData(grh_index).pixelHeight
           
    hdcsrc = CreateCompatibleDC(desthDC)
 
    file_path = App.Path & "\GRAFICOS\" & GrhData(grh_index).FileNum & ".bmp"
    PrevObj = SelectObject(hdcsrc, LoadPicture(file_path))
   
    'BitBlt desthDC, screen_x, screen_y, src_width, src_height, hdcsrc, src_x, src_y, vbSrcCopy
   TransparentBlt desthDC, screen_x, screen_y, src_width, src_height, hdcsrc, src_x, src_y, src_width, src_height, RGB(0, 0, 0)
 
   Call DeleteObject(SelectObject(hdcsrc, PrevObj))
   DeleteDC hdcsrc
 
Err:
If Err.number = 481 Then MsgBox "Imposible cargar recurso error: " & "481"
 
End Sub

Private Sub Engine_Font_Initialize()

Dim A As Integer

font_types(1).font_size = 9
font_types(1).ascii_code(48) = 21452
font_types(1).ascii_code(49) = 21453
font_types(1).ascii_code(50) = 21454
font_types(1).ascii_code(51) = 21455
font_types(1).ascii_code(52) = 21456
font_types(1).ascii_code(53) = 21457
font_types(1).ascii_code(54) = 21458
font_types(1).ascii_code(55) = 21459
font_types(1).ascii_code(56) = 21460
font_types(1).ascii_code(57) = 21461
For A = 0 To 25
font_types(1).ascii_code(A + 97) = 21400 + A
Next A

For A = 0 To 25
font_types(1).ascii_code(A + 65) = 21426 + A
Next A
font_types(1).ascii_code(33) = 21462
font_types(1).ascii_code(161) = 21463
font_types(1).ascii_code(34) = 21464
font_types(1).ascii_code(36) = 21465
font_types(1).ascii_code(191) = 21466
font_types(1).ascii_code(35) = 21467
font_types(1).ascii_code(36) = 21468
font_types(1).ascii_code(37) = 21469
font_types(1).ascii_code(38) = 21470
font_types(1).ascii_code(47) = 21471
font_types(1).ascii_code(92) = 21472
font_types(1).ascii_code(40) = 21473
font_types(1).ascii_code(41) = 21474
font_types(1).ascii_code(61) = 21475
font_types(1).ascii_code(39) = 21476
font_types(1).ascii_code(123) = 21477
font_types(1).ascii_code(125) = 21478
font_types(1).ascii_code(95) = 21479
font_types(1).ascii_code(45) = 21480
font_types(1).ascii_code(63) = 21465
font_types(1).ascii_code(64) = 21481
font_types(1).ascii_code(94) = 21482
font_types(1).ascii_code(91) = 21483
font_types(1).ascii_code(93) = 21484
font_types(1).ascii_code(60) = 21485
font_types(1).ascii_code(62) = 21486
font_types(1).ascii_code(42) = 21487
font_types(1).ascii_code(43) = 21488
font_types(1).ascii_code(46) = 21489
font_types(1).ascii_code(44) = 21490
font_types(1).ascii_code(58) = 21491
font_types(1).ascii_code(59) = 21492
font_types(1).ascii_code(124) = 21493
font_types(1).ascii_code(252) = 21800
font_types(1).ascii_code(220) = 21801
font_types(1).ascii_code(225) = 21802
font_types(1).ascii_code(233) = 21803
font_types(1).ascii_code(237) = 21804
font_types(1).ascii_code(243) = 21805
font_types(1).ascii_code(250) = 21806
font_types(1).ascii_code(253) = 21807
font_types(1).ascii_code(193) = 21808
font_types(1).ascii_code(201) = 21809
font_types(1).ascii_code(205) = 21810
font_types(1).ascii_code(211) = 21811
font_types(1).ascii_code(218) = 21812
font_types(1).ascii_code(221) = 21813
font_types(1).ascii_code(224) = 21814
font_types(1).ascii_code(232) = 21815
font_types(1).ascii_code(236) = 21816
font_types(1).ascii_code(242) = 21817
font_types(1).ascii_code(249) = 21818
font_types(1).ascii_code(192) = 21819
font_types(1).ascii_code(200) = 21820
font_types(1).ascii_code(204) = 21821
font_types(1).ascii_code(210) = 21822
font_types(1).ascii_code(217) = 21823
font_types(1).ascii_code(241) = 21824
font_types(1).ascii_code(209) = 21825
font_types(1).ascii_code(196) = 25238
font_types(1).ascii_code(194) = 25239
font_types(1).ascii_code(203) = 25240
font_types(1).ascii_code(207) = 25241
font_types(1).ascii_code(214) = 25242
font_types(1).ascii_code(212) = 25243

font_types(2).font_size = 9
font_types(2).ascii_code(97) = 21936
font_types(2).ascii_code(108) = 21937
font_types(2).ascii_code(115) = 21938
font_types(2).ascii_code(70) = 21939
font_types(2).ascii_code(48) = 21940
font_types(2).ascii_code(49) = 21941
font_types(2).ascii_code(50) = 21942
font_types(2).ascii_code(51) = 21943
font_types(2).ascii_code(52) = 21944
font_types(2).ascii_code(53) = 21945
font_types(2).ascii_code(54) = 21946
font_types(2).ascii_code(55) = 21947
font_types(2).ascii_code(56) = 21948
font_types(2).ascii_code(57) = 21949
font_types(2).ascii_code(33) = 21950
font_types(2).ascii_code(161) = 21951
font_types(2).ascii_code(42) = 21952

End Sub
Sub Text_Render(Texto As String, x As Integer, y As Integer, ByRef text_color() As Long, Optional ByVal font_index As Integer = 1, Optional multi_line As Boolean = False)

Dim A, b, c, d, e, f As Integer
Dim graf As Grh

Dim temp_array(3) As Long 'Si le queres dar color a la letra pasa este parametro dsp xD
temp_array(0) = text_color(0)
temp_array(1) = text_color(1)
temp_array(2) = text_color(2)
temp_array(3) = text_color(3)

If (Len(Texto) = 0) Then Exit Sub

d = 0
If multi_line = False Then
    For A = 1 To Len(Texto)
        b = Asc(mid(Texto, A, 1))
        graf.grhindex = font_types(font_index).ascii_code(b)
        If b <> 32 Then
            If graf.grhindex <> 0 Then
                'mega sombra O-matica
                graf.grhindex = font_types(font_index).ascii_code(b) + 100
                Grh_Render graf, (x + d) + 1, y + 1, temp_array, False, False, False
                graf.grhindex = font_types(font_index).ascii_code(b)
                Grh_Render graf, (x + d), y, temp_array, False, False, False
                d = d + GrhData(GrhData(graf.grhindex).Frames(1)).pixelWidth '+ 1
            End If
        Else
            d = d + 4
        End If
    Next A
Else
    e = 0
    f = 0
    For A = 1 To Len(Texto)
        b = Asc(mid(Texto, A, 1))
        graf.grhindex = font_types(font_index).ascii_code(b)
        If b = 32 Or b = 13 Then
            If e >= 20 Then 'reemplazar por lo que os plazca
                f = f + 1
                e = 0
                d = 0
            Else
                If b = 32 Then d = d + 4
            End If
        Else
            If graf.grhindex > 12 Then
                'mega sombra O-matica
                graf.grhindex = font_types(font_index).ascii_code(b) + 100
                Grh_Render graf, (x + d) + 1, y + 1 + f * 14, temp_array, False, False, False
                graf.grhindex = font_types(font_index).ascii_code(b)
                Grh_Render graf, (x + d), y + f * 14, temp_array, False, False, False '14 es el height de esta fuente dsp lo hacemos dinamico
                d = d + GrhData(GrhData(graf.grhindex).Frames(1)).pixelWidth '+ 1
            End If
       End If
       e = e + 1
    Next A
End If

End Sub


Private Function Particle_Group_Next_Open() As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    If particle_group_last = 0 Then
        Particle_Group_Next_Open = 1
        Exit Function
    End If
    
    loopc = 1
    Do Until particle_group_list(loopc).active = False
        If loopc = particle_group_last Then
            Particle_Group_Next_Open = particle_group_last + 1
            Exit Function
        End If
        loopc = loopc + 1
    Loop
    
    Particle_Group_Next_Open = loopc
Exit Function
ErrorHandler:
    Particle_Group_Next_Open = 1
End Function

Private Function Particle_Group_Check(ByVal particle_group_index As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'**************************************************************
    'check index
    If particle_group_index > 0 And particle_group_index <= particle_group_last Then
        If particle_group_list(particle_group_index).active Then
            Particle_Group_Check = True
        End If
    End If
End Function

Public Function Particle_Group_Create(ByVal map_x As Integer, ByVal map_y As Integer, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                        Optional ByVal particle_count As Long = 20, Optional ByVal stream_type As Long = 1, _
                                        Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                        Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
                                        Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal angle As Integer, _
                                        Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                        Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                        Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                        Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                        Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                        Optional bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                                        Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                        Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                        Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
                                        Optional grh_resizex As Integer, Optional grh_resizey As Integer)
'**************************************************************
'Author: Aaron Perkins
'Modified by: Ryan Cain (Onezero)
'Last Modify Date: 5/14/2003
'Returns the particle_group_index if successful, else 0
'Modified by Juan Martín Sotuyo Dodero
'Modified by Augusto José Rando
'**************************************************************
    
    If (map_x <> -1) And (map_y <> -1) Then
        If Map_Particle_Group_Get(map_x, map_y) = 0 Then
            Particle_Group_Create = Particle_Group_Next_Open
            Particle_Group_Make Particle_Group_Create, map_x, map_y, particle_count, stream_type, grh_index_list(), rgb_list(), alpha_blend, alive_counter, frame_speed, id, x1, y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, x2, y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin, grh_resize, grh_resizex, grh_resizey
        End If
    Else
        Particle_Group_Create = Particle_Group_Next_Open
        Particle_Group_Make Particle_Group_Create, map_x, map_y, particle_count, stream_type, grh_index_list(), rgb_list(), alpha_blend, alive_counter, frame_speed, id, x1, y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, x2, y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin, grh_resize, grh_resizex, grh_resizey
    End If

End Function

Public Function Particle_Group_Remove(ByVal particle_group_index As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    'Make sure it's a legal index
    If Particle_Group_Check(particle_group_index) Then
        Particle_Group_Destroy particle_group_index
        Particle_Group_Remove = True
    End If
End Function

Public Function Particle_Group_Remove_All() As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    Dim Index As Long
    
    For Index = 1 To particle_group_last
        'Make sure it's a legal index
        If Particle_Group_Check(Index) Then
            Particle_Group_Destroy Index
        End If
    Next Index
    
    Particle_Group_Remove_All = True
End Function

Public Function Particle_Group_Find(ByVal id As Long) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'Find the index related to the handle
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    loopc = 1
    Do Until particle_group_list(loopc).id = id
        If loopc = particle_group_last Then
            Particle_Group_Find = 0
            Exit Function
        End If
        loopc = loopc + 1
    Loop
    
    Particle_Group_Find = loopc
Exit Function
ErrorHandler:
    Particle_Group_Find = 0
End Function
Public Function Particle_Get_Type(ByVal particle_group_index As Long) As Byte
On Error GoTo ErrorHandler:
    Particle_Get_Type = particle_group_list(particle_group_index).stream_type
Exit Function
ErrorHandler:
    Particle_Get_Type = 0
End Function
Private Sub Particle_Group_Destroy(ByVal particle_group_index As Long)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
On Error Resume Next
    Dim temp As particle_group
    Dim i As Integer
    
    If particle_group_list(particle_group_index).map_x > 0 And particle_group_list(particle_group_index).map_y > 0 Then
        MapData(particle_group_list(particle_group_index).map_x, particle_group_list(particle_group_index).map_y).particle_group_index = 0
    ElseIf particle_group_list(particle_group_index).char_index Then
        If Char_Check(particle_group_list(particle_group_index).char_index) Then
            For i = 1 To charlist(particle_group_list(particle_group_index).char_index).particle_count
                If charlist(particle_group_list(particle_group_index).char_index).particle_group(i) = particle_group_index Then
                    charlist(particle_group_list(particle_group_index).char_index).particle_group(i) = 0
        
                    'We don't resize arrays by now, it's really a waste...
                    'If i = UBound(CharList(particle_group_list(particle_group_index).char_index).particle_group) Then
                    '    CharList(particle_group_list(particle_group_index).char_index).particle_count = i - 1
                    '    ReDim Preserve CharList(particle_group_list(particle_group_index).char_index).particle_group(1 To (i - 1)) As Long
                    'End If
        
                    Exit For
                End If
            Next i
        End If
    ElseIf particle_group_index = meteo_particle Then
        meteo_particle = 0
    End If
    
    particle_group_list(particle_group_index) = temp
    
    'Update array size
    If particle_group_index = particle_group_last Then
        Do Until particle_group_list(particle_group_last).active
            particle_group_last = particle_group_last - 1
            If particle_group_last = 0 Then
                particle_group_count = 0
                Exit Sub
            End If
        Loop
        Debug.Print particle_group_last & "," & UBound(particle_group_list)
        ReDim Preserve particle_group_list(1 To particle_group_last) As particle_group
    End If
    particle_group_count = particle_group_count - 1
End Sub
Private Sub Particle_Group_Render(ByVal particle_group_index As Long, ByVal screen_x As Integer, ByVal screen_y As Integer)
'*****************************************************************
'Author: Aaron Perkins
'Modified by: Ryan Cain (Onezero)
'Modified by: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/15/2003
'Renders a particle stream at a paticular screen point
'*****************************************************************
    Dim loopc As Long
    Dim temp_rgb(0 To 3) As Long
    Dim no_move As Boolean
    
    If GetTickCount - particle_group_list(particle_group_index).live > particle_group_list(particle_group_index).liv1 And Not particle_group_list(particle_group_index).liv1 = -1 Then
        Particle_Group_Destroy particle_group_index 'mm eso es del resu.-
        Exit Sub
    End If
        
    With particle_group_list(particle_group_index)
        'Set colors
        temp_rgb(0) = .rgb_list(0)
        temp_rgb(1) = .rgb_list(1)
        temp_rgb(2) = .rgb_list(2)
        temp_rgb(3) = .rgb_list(3)
        
        If GetTickCount - .live > .liv1 And Not .liv1 = -1 Then
            Particle_Group_Destroy particle_group_index
            Exit Sub
        End If
        
        'See if it is time to move a particle
        .frame_counter = .frame_counter + timerTicksPerFrame
        If .frame_counter > .frame_speed Then
            .frame_counter = 0
            no_move = False
        Else
            no_move = True
        End If
            
        'If it's still alive render all the particles inside
        For loopc = 1 To .particle_count
                
        'Render particle
            Particle_Render .particle_stream(loopc), _
                        screen_x, screen_y, _
                        .grh_index_list(Round(General_Random_Number(1, .grh_index_count), 0)), _
                        temp_rgb(), _
                        .alpha_blend, no_move, _
                        .x1, .y1, .angle, _
                        .vecx1, .vecx2, _
                        .vecy1, .vecy2, _
                        .life1, .life2, _
                        .fric, .spin_speedL, _
                        .gravity, .grav_strength, _
                        .bounce_strength, .x2, _
                        .y2, .XMove, _
                        .move_x1, .move_x2, _
                        .move_y1, .move_y2, _
                        .YMove, .spin_speedH, _
                        .spin, .grh_resize, .grh_resizex, .grh_resizey
        Next loopc
                
    If no_move = False Then
        'Update the group alive counter
        If .never_die = False Then
            .alive_counter = .alive_counter - 1
        End If
    End If
    
        
    End With
End Sub


Private Sub Particle_Render(ByRef temp_particle As Particle, ByVal screen_x As Integer, ByVal screen_y As Integer, _
                            ByVal grh_index As Long, ByRef rgb_list() As Long, _
                            Optional ByVal alpha_blend As Boolean, Optional ByVal no_move As Boolean, _
                            Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal angle As Integer, _
                            Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                            Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                            Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                            Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                            Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                            Optional ByVal bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                            Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                            Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                            Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
                            Optional grh_resizex As Integer, Optional grh_resizey As Integer)
'**************************************************************
'Author: Aaron Perkins
'Modified by: Ryan Cain (Onezero)
'Modified by: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/15/2003
'**************************************************************
    If no_move = False Then
        If temp_particle.alive_counter = 0 Then
            'Start new particle
            InitGrh temp_particle.Grh, grh_index, alpha_blend
            temp_particle.x = General_Random_Number(x1, x2) - 16
            temp_particle.y = General_Random_Number(y1, y2) - 16
            temp_particle.vector_x = General_Random_Number(vecx1, vecx2)
            temp_particle.vector_y = General_Random_Number(vecy1, vecy2)
            temp_particle.angle = angle
            temp_particle.alive_counter = General_Random_Number(life1, life2)
            temp_particle.friction = fric
        Else
            'Continue old particle
            'Do gravity
            If gravity = True Then
                temp_particle.vector_y = temp_particle.vector_y + grav_strength
                If temp_particle.y > 0 Then
                    'bounce
                    temp_particle.vector_y = bounce_strength
                End If
            End If
            'Do rotation
            If spin = True Then temp_particle.Grh.angle = temp_particle.Grh.angle + (General_Random_Number(spin_speedL, spin_speedH) / 100)
            If temp_particle.angle >= 360 Then
                temp_particle.angle = 0
            End If
            
            If XMove = True Then temp_particle.vector_x = General_Random_Number(move_x1, move_x2)
            If YMove = True Then temp_particle.vector_y = General_Random_Number(move_y1, move_y2)
        End If
        
        'Add in vector
        temp_particle.x = temp_particle.x + (temp_particle.vector_x \ temp_particle.friction)
        temp_particle.y = temp_particle.y + (temp_particle.vector_y \ temp_particle.friction)
    
        'decrement counter
         temp_particle.alive_counter = temp_particle.alive_counter - 1
    End If
    
    'Draw it
    If grh_resize = True Then
        If temp_particle.Grh.grhindex Then
             Grh_Render_Advance temp_particle.Grh, temp_particle.x + screen_x, temp_particle.y + screen_y, grh_resizex, grh_resizey, rgb_list(), True, True, alpha_blend
            Exit Sub
        End If
    End If
    If temp_particle.Grh.grhindex Then
        Grh_Render temp_particle.Grh, temp_particle.x + screen_x, temp_particle.y + screen_y, rgb_list(), True, True, alpha_blend
    End If
End Sub
Private Sub Particle_Group_Make(ByVal particle_group_index As Long, ByVal map_x As Integer, ByVal map_y As Integer, _
                                ByVal particle_count As Long, ByVal stream_type As Long, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
                                Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal angle As Integer, _
                                Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                Optional bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                                Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
                                Optional grh_resizex As Integer, Optional grh_resizey As Integer)
                                
'*****************************************************************
'Author: Aaron Perkins
'Modified by: Ryan Cain (Onezero)
'Last Modify Date: 5/15/2003
'Makes a new particle effect
'Modified by Juan Martín Sotuyo Dodero
'*****************************************************************
    'Update array size
    If particle_group_index > particle_group_last Then
        particle_group_last = particle_group_index
        ReDim Preserve particle_group_list(1 To particle_group_last)
    End If
    particle_group_count = particle_group_count + 1
    
    'Make active
    particle_group_list(particle_group_index).active = True
    
    'Map pos
    If (map_x <> -1) And (map_y <> -1) Then
        particle_group_list(particle_group_index).map_x = map_x
        particle_group_list(particle_group_index).map_y = map_y
    End If
    
    'Grh list
    ReDim particle_group_list(particle_group_index).grh_index_list(1 To UBound(grh_index_list))
    particle_group_list(particle_group_index).grh_index_list() = grh_index_list()
    particle_group_list(particle_group_index).grh_index_count = UBound(grh_index_list)
    
    'Sets alive vars
    If alive_counter = -1 Then
        particle_group_list(particle_group_index).alive_counter = -1
        particle_group_list(particle_group_index).liv1 = -1
        particle_group_list(particle_group_index).never_die = True
    Else
        particle_group_list(particle_group_index).alive_counter = alive_counter
        particle_group_list(particle_group_index).liv1 = alive_counter
        particle_group_list(particle_group_index).never_die = False
    End If
    
    'alpha blending
    particle_group_list(particle_group_index).alpha_blend = alpha_blend
    
    'stream type
    particle_group_list(particle_group_index).stream_type = stream_type
    
    'speed
    particle_group_list(particle_group_index).frame_speed = frame_speed
    
    particle_group_list(particle_group_index).x1 = x1
    particle_group_list(particle_group_index).y1 = y1
    particle_group_list(particle_group_index).x2 = x2
    particle_group_list(particle_group_index).y2 = y2
    particle_group_list(particle_group_index).angle = angle
    particle_group_list(particle_group_index).vecx1 = vecx1
    particle_group_list(particle_group_index).vecx2 = vecx2
    particle_group_list(particle_group_index).vecy1 = vecy1
    particle_group_list(particle_group_index).vecy2 = vecy2
    particle_group_list(particle_group_index).life1 = life1
    particle_group_list(particle_group_index).life2 = life2
    particle_group_list(particle_group_index).fric = fric
    particle_group_list(particle_group_index).spin = spin
    particle_group_list(particle_group_index).spin_speedL = spin_speedL
    particle_group_list(particle_group_index).spin_speedH = spin_speedH
    particle_group_list(particle_group_index).gravity = gravity
    particle_group_list(particle_group_index).grav_strength = grav_strength
    particle_group_list(particle_group_index).bounce_strength = bounce_strength
    particle_group_list(particle_group_index).XMove = XMove
    particle_group_list(particle_group_index).YMove = YMove
    particle_group_list(particle_group_index).move_x1 = move_x1
    particle_group_list(particle_group_index).move_x2 = move_x2
    particle_group_list(particle_group_index).move_y1 = move_y1
    particle_group_list(particle_group_index).move_y2 = move_y2
    
    particle_group_list(particle_group_index).rgb_list(0) = rgb_list(0)
    particle_group_list(particle_group_index).rgb_list(1) = rgb_list(1)
    particle_group_list(particle_group_index).rgb_list(2) = rgb_list(2)
    particle_group_list(particle_group_index).rgb_list(3) = rgb_list(3)
    
    particle_group_list(particle_group_index).grh_resize = grh_resize
    particle_group_list(particle_group_index).grh_resizex = grh_resizex
    particle_group_list(particle_group_index).grh_resizey = grh_resizey
    
    'handle
    particle_group_list(particle_group_index).id = id
    
    particle_group_list(particle_group_index).live = GetTickCount()
    
    'create particle stream
    particle_group_list(particle_group_index).particle_count = particle_count
    ReDim particle_group_list(particle_group_index).particle_stream(1 To particle_count)
    
    'plot particle group on map
    If (map_x <> -1) And (map_y <> -1) Then
        MapData(map_x, map_y).particle_group_index = particle_group_index
    End If
    
End Sub
Public Function Map_Particle_Group_Get(ByVal map_x As Integer, ByVal map_y As Integer) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/20/2003
'Checks to see if a tile position has a particle_group_index and return it
'*****************************************************************
    If InMapBounds(map_x, map_y) Then
        Map_Particle_Group_Get = MapData(map_x, map_y).particle_group_index
    Else
        Map_Particle_Group_Get = 0
    End If
End Function
Public Function General_Char_Particle_Create(ByVal ParticulaInd As Long, ByVal char_index As Integer, Optional ByVal particle_life As Long = 0) As Long

Dim rgb_list(0 To 3) As Long
rgb_list(0) = RGB(StreamData(ParticulaInd).colortint(0).r, StreamData(ParticulaInd).colortint(0).g, StreamData(ParticulaInd).colortint(0).b)
rgb_list(1) = RGB(StreamData(ParticulaInd).colortint(1).r, StreamData(ParticulaInd).colortint(1).g, StreamData(ParticulaInd).colortint(1).b)
rgb_list(2) = RGB(StreamData(ParticulaInd).colortint(2).r, StreamData(ParticulaInd).colortint(2).g, StreamData(ParticulaInd).colortint(2).b)
rgb_list(3) = RGB(StreamData(ParticulaInd).colortint(3).r, StreamData(ParticulaInd).colortint(3).g, StreamData(ParticulaInd).colortint(3).b)

General_Char_Particle_Create = Engine.Char_Particle_Group_Create(char_index, StreamData(ParticulaInd).grh_list, rgb_list(), StreamData(ParticulaInd).NumOfParticles, ParticulaInd, _
    StreamData(ParticulaInd).AlphaBlend, IIf(particle_life = 0, StreamData(ParticulaInd).life_counter, particle_life), StreamData(ParticulaInd).speed, , StreamData(ParticulaInd).x1, StreamData(ParticulaInd).y1, StreamData(ParticulaInd).angle, _
    StreamData(ParticulaInd).vecx1, StreamData(ParticulaInd).vecx2, StreamData(ParticulaInd).vecy1, StreamData(ParticulaInd).vecy2, _
    StreamData(ParticulaInd).life1, StreamData(ParticulaInd).life2, StreamData(ParticulaInd).friction, StreamData(ParticulaInd).spin_speedL, _
    StreamData(ParticulaInd).gravity, StreamData(ParticulaInd).grav_strength, StreamData(ParticulaInd).bounce_strength, StreamData(ParticulaInd).x2, _
    StreamData(ParticulaInd).y2, StreamData(ParticulaInd).XMove, StreamData(ParticulaInd).move_x1, StreamData(ParticulaInd).move_x2, StreamData(ParticulaInd).move_y1, _
    StreamData(ParticulaInd).move_y2, StreamData(ParticulaInd).YMove, StreamData(ParticulaInd).spin_speedH, StreamData(ParticulaInd).spin, _
    StreamData(ParticulaInd).grh_resize, StreamData(ParticulaInd).grh_resizex, StreamData(ParticulaInd).grh_resizey)

End Function

Public Function General_Particle_Create(ByVal ParticulaInd As Long, ByVal x As Integer, ByVal y As Integer, Optional ByVal particle_life As Long = 0) As Long

Dim rgb_list(0 To 3) As Long
rgb_list(0) = RGB(StreamData(ParticulaInd).colortint(0).r, StreamData(ParticulaInd).colortint(0).g, StreamData(ParticulaInd).colortint(0).b)
rgb_list(1) = RGB(StreamData(ParticulaInd).colortint(1).r, StreamData(ParticulaInd).colortint(1).g, StreamData(ParticulaInd).colortint(1).b)
rgb_list(2) = RGB(StreamData(ParticulaInd).colortint(2).r, StreamData(ParticulaInd).colortint(2).g, StreamData(ParticulaInd).colortint(2).b)
rgb_list(3) = RGB(StreamData(ParticulaInd).colortint(3).r, StreamData(ParticulaInd).colortint(3).g, StreamData(ParticulaInd).colortint(3).b)

General_Particle_Create = Engine.Particle_Group_Create(x, y, StreamData(ParticulaInd).grh_list, rgb_list(), StreamData(ParticulaInd).NumOfParticles, ParticulaInd, _
    StreamData(ParticulaInd).AlphaBlend, IIf(particle_life = 0, StreamData(ParticulaInd).life_counter, particle_life), StreamData(ParticulaInd).speed, , StreamData(ParticulaInd).x1, StreamData(ParticulaInd).y1, StreamData(ParticulaInd).angle, _
    StreamData(ParticulaInd).vecx1, StreamData(ParticulaInd).vecx2, StreamData(ParticulaInd).vecy1, StreamData(ParticulaInd).vecy2, _
    StreamData(ParticulaInd).life1, StreamData(ParticulaInd).life2, StreamData(ParticulaInd).friction, StreamData(ParticulaInd).spin_speedL, _
    StreamData(ParticulaInd).gravity, StreamData(ParticulaInd).grav_strength, StreamData(ParticulaInd).bounce_strength, StreamData(ParticulaInd).x2, _
    StreamData(ParticulaInd).y2, StreamData(ParticulaInd).XMove, StreamData(ParticulaInd).move_x1, StreamData(ParticulaInd).move_x2, StreamData(ParticulaInd).move_y1, _
    StreamData(ParticulaInd).move_y2, StreamData(ParticulaInd).YMove, StreamData(ParticulaInd).spin_speedH, StreamData(ParticulaInd).spin, _
    StreamData(ParticulaInd).grh_resize, StreamData(ParticulaInd).grh_resizex, StreamData(ParticulaInd).grh_resizey)

End Function
Public Function Char_Particle_Group_Create(ByVal char_index As Integer, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                        Optional ByVal particle_count As Long = 20, Optional ByVal stream_type As Long = 1, _
                                        Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                        Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
                                        Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal angle As Integer, _
                                        Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                        Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                        Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                        Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                        Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                        Optional bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                                        Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                        Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                        Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
                                        Optional grh_resizex As Integer, Optional grh_resizey As Integer)
'**************************************************************
'Author: Augusto José Rando
'**************************************************************
    Dim char_part_free_index As Integer
    
    'If Char_Particle_Group_Find(char_index, stream_type) Then Exit Function ' hay que ver si dejar o sacar esto...
    If Not Char_Check(char_index) Then Exit Function
    char_part_free_index = Char_Particle_Group_Next_Open(char_index)
    
    If char_part_free_index > 0 Then
        Char_Particle_Group_Create = Particle_Group_Next_Open
        Char_Particle_Group_Make Char_Particle_Group_Create, char_index, char_part_free_index, particle_count, stream_type, grh_index_list(), rgb_list(), alpha_blend, alive_counter, frame_speed, id, x1, y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, x2, y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin, grh_resize, grh_resizex, grh_resizey
    End If

End Function

Public Function Char_Particle_Group_Remove(ByVal char_index As Integer, ByVal stream_type As Long)
'**************************************************************
'Author: Augusto José Rando
'**************************************************************
    Dim char_part_index As Integer
    
    If Char_Check(char_index) Then
        char_part_index = Char_Particle_Group_Find(char_index, stream_type)
        If char_part_index = -1 Then Exit Function
        Call Particle_Group_Remove(char_part_index)
    End If

End Function

Public Function Char_Particle_Group_Remove_All(ByVal char_index As Integer)
'**************************************************************
'Author: Augusto José Rando
'**************************************************************
    Dim i As Integer
    
    If Char_Check(char_index) And Not charlist(char_index).particle_count = 0 Then
        For i = 1 To UBound(charlist(char_index).particle_group)
            If charlist(char_index).particle_group(i) <> 0 Then Call Particle_Group_Remove(charlist(char_index).particle_group(i))
        Next i
        Erase charlist(char_index).particle_group
        charlist(char_index).particle_count = 0
    End If
    
End Function

Private Function Char_Particle_Group_Find(ByVal char_index As Integer, ByVal stream_type As Long) As Integer
'*****************************************************************
'Author: Augusto José Rando
'Modified: returns slot or -1
'*****************************************************************

Dim i As Integer

For i = 1 To charlist(char_index).particle_count
    If particle_group_list(charlist(char_index).particle_group(i)).stream_type = stream_type Then
        Char_Particle_Group_Find = charlist(char_index).particle_group(i)
        Exit Function
    End If
Next i

Char_Particle_Group_Find = -1

End Function

Private Function Char_Particle_Group_Next_Open(ByVal char_index As Integer) As Integer
'*****************************************************************
'Author: Augusto José Rando
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    If charlist(char_index).particle_count = 0 Then
        Char_Particle_Group_Next_Open = charlist(char_index).particle_count + 1
        charlist(char_index).particle_count = Char_Particle_Group_Next_Open
        ReDim Preserve charlist(char_index).particle_group(1 To Char_Particle_Group_Next_Open) As Long
        Exit Function
    End If
    
    loopc = 1
    Do Until charlist(char_index).particle_group(loopc) = 0
        If loopc = charlist(char_index).particle_count Then
            Char_Particle_Group_Next_Open = charlist(char_index).particle_count + 1
            charlist(char_index).particle_count = Char_Particle_Group_Next_Open
            ReDim Preserve charlist(char_index).particle_group(1 To Char_Particle_Group_Next_Open) As Long
            Exit Function
        End If
        loopc = loopc + 1
    Loop
    
    Char_Particle_Group_Next_Open = loopc

Exit Function

ErrorHandler:
    charlist(char_index).particle_count = 1
    ReDim charlist(char_index).particle_group(1 To 1) As Long
    Char_Particle_Group_Next_Open = 1

End Function
Private Function Char_Check(ByVal char_index As Integer) As Boolean
'**************************************************************
'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 1/04/2003
'
'**************************************************************
    'check char_index
    If char_index > 0 And char_index <= LastChar Then
        Char_Check = (charlist(char_index).heading > 0)
    End If
    
End Function
Private Sub Char_Particle_Group_Make(ByVal particle_group_index As Long, ByVal char_index As Integer, ByVal particle_char_index As Integer, _
                                ByVal particle_count As Long, ByVal stream_type As Long, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
                                Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal angle As Integer, _
                                Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                Optional bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                                Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
                                Optional grh_resizex As Integer, Optional grh_resizey As Integer)
                                
'*****************************************************************
'Author: Aaron Perkins
'Modified by: Ryan Cain (Onezero)
'Last Modify Date: 5/15/2003
'Makes a new particle effect
'Modified by Juan Martín Sotuyo Dodero
'*****************************************************************
    'Update array size
    If particle_group_index > particle_group_last Then
        particle_group_last = particle_group_index
        ReDim Preserve particle_group_list(1 To particle_group_last)
    End If
    particle_group_count = particle_group_count + 1
    
    'Make active
    particle_group_list(particle_group_index).active = True
    
    'Char index
    particle_group_list(particle_group_index).char_index = char_index
    
    'Grh list
    ReDim particle_group_list(particle_group_index).grh_index_list(1 To UBound(grh_index_list))
    particle_group_list(particle_group_index).grh_index_list() = grh_index_list()
    particle_group_list(particle_group_index).grh_index_count = UBound(grh_index_list)
    
    'Sets alive vars
    If alive_counter = -1 Then
        particle_group_list(particle_group_index).alive_counter = -1
        particle_group_list(particle_group_index).liv1 = -1
        particle_group_list(particle_group_index).never_die = True
    Else
        particle_group_list(particle_group_index).alive_counter = alive_counter
        particle_group_list(particle_group_index).liv1 = alive_counter
        particle_group_list(particle_group_index).never_die = False
    End If
    
    'alpha blending
    particle_group_list(particle_group_index).alpha_blend = alpha_blend
    
    'stream type
    particle_group_list(particle_group_index).stream_type = stream_type
    
    'speed
    particle_group_list(particle_group_index).frame_speed = frame_speed
    
    particle_group_list(particle_group_index).x1 = x1
    particle_group_list(particle_group_index).y1 = y1
    particle_group_list(particle_group_index).x2 = x2
    particle_group_list(particle_group_index).y2 = y2
    particle_group_list(particle_group_index).angle = angle
    particle_group_list(particle_group_index).vecx1 = vecx1
    particle_group_list(particle_group_index).vecx2 = vecx2
    particle_group_list(particle_group_index).vecy1 = vecy1
    particle_group_list(particle_group_index).vecy2 = vecy2
    particle_group_list(particle_group_index).life1 = life1
    particle_group_list(particle_group_index).life2 = life2
    particle_group_list(particle_group_index).fric = fric
    particle_group_list(particle_group_index).spin = spin
    particle_group_list(particle_group_index).spin_speedL = spin_speedL
    particle_group_list(particle_group_index).spin_speedH = spin_speedH
    particle_group_list(particle_group_index).gravity = gravity
    particle_group_list(particle_group_index).grav_strength = grav_strength
    particle_group_list(particle_group_index).bounce_strength = bounce_strength
    particle_group_list(particle_group_index).XMove = XMove
    particle_group_list(particle_group_index).YMove = YMove
    particle_group_list(particle_group_index).move_x1 = move_x1
    particle_group_list(particle_group_index).move_x2 = move_x2
    particle_group_list(particle_group_index).move_y1 = move_y1
    particle_group_list(particle_group_index).move_y2 = move_y2
    
    particle_group_list(particle_group_index).rgb_list(0) = rgb_list(0)
    particle_group_list(particle_group_index).rgb_list(1) = rgb_list(1)
    particle_group_list(particle_group_index).rgb_list(2) = rgb_list(2)
    particle_group_list(particle_group_index).rgb_list(3) = rgb_list(3)
    
    particle_group_list(particle_group_index).grh_resize = grh_resize
    particle_group_list(particle_group_index).grh_resizex = grh_resizex
    particle_group_list(particle_group_index).grh_resizey = grh_resizey
    
    'handle
    particle_group_list(particle_group_index).id = id
    particle_group_list(particle_group_index).live = GetTickCount()
    
    'create particle stream
    particle_group_list(particle_group_index).particle_count = particle_count
    ReDim particle_group_list(particle_group_index).particle_stream(1 To particle_count)
    
    'plot particle group on char
    charlist(char_index).particle_group(particle_char_index) = particle_group_index
End Sub

Private Sub Grh_Render_Advance(ByRef Grh As Grh, ByVal screen_x As Integer, ByVal screen_y As Integer, ByVal height As Integer, ByVal width As Integer, ByRef rgb_list() As Long, Optional ByVal h_center As Boolean, Optional ByVal v_center As Boolean, Optional ByVal alpha_blend As Boolean = False)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'Last Modify Date: 11/19/2003
'Similar to Grh_Render, but let´s you resize the Grh
'**************************************************************
    Dim tile_width As Integer
    Dim tile_height As Integer
    Dim grh_index As Long
    
    'Animation
    If Grh.Started Then
        Grh.FrameCounter = Grh.FrameCounter + (timerTicksPerFrame * Grh.speed)
        If Grh.FrameCounter > GrhData(Grh.grhindex).NumFrames Then
            'If Grh.noloop Then
            '    Grh.FrameCounter = GrhData(Grh.GrhIndex).NumFrames
            'Else
                Grh.FrameCounter = 1
            'End If
        End If
    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    If Grh.FrameCounter = 0 Then Grh.FrameCounter = 1
    grh_index = GrhData(Grh.grhindex).Frames(Grh.FrameCounter)
    
    'Center Grh over X, Y pos
    If GrhData(Grh.grhindex).TileWidth <> 1 Then
        screen_x = screen_x - Int(GrhData(Grh.grhindex).TileWidth * 16) + 16
    End If
    
    If GrhData(Grh.grhindex).TileHeight <> 1 Then
        screen_y = screen_y - Int(GrhData(Grh.grhindex).TileHeight * 32) + 32
    End If
    
    'Draw it to device
    Device_Box_Textured_Render_Advance grh_index, _
        screen_x, screen_y, _
        GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight, _
        rgb_list, _
        GrhData(grh_index).sX, GrhData(grh_index).sY, _
        width, height, alpha_blend, Grh.angle
End Sub
Private Sub Grh_Render(ByRef Grh As Grh, ByVal screen_x As Integer, ByVal screen_y As Integer, ByRef rgb_list() As Long, Optional ByVal h_centered As Boolean = True, Optional ByVal v_centered As Boolean = True, Optional ByVal alpha_blend As Boolean = False)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/28/2003
'Modified by Juan Martín Sotuyo Dodero
'Added centering
'**************************************************************
    Dim tile_width As Integer
    Dim tile_height As Integer
    Dim grh_index As Long
    
    If Grh.grhindex = 0 Then Exit Sub
        
    'Animation
    If Grh.Started Then
        Grh.FrameCounter = Grh.FrameCounter + (timerTicksPerFrame * Grh.speed)
        If Grh.FrameCounter > GrhData(Grh.grhindex).NumFrames Then
            Grh.FrameCounter = 1
        End If
    End If


    'Figure out what frame to draw (always 1 if not animated)
    If Grh.FrameCounter = 0 Then Grh.FrameCounter = 1
    
    grh_index = GrhData(Grh.grhindex).Frames(Grh.FrameCounter)
    
    If grh_index <= 0 Then Exit Sub
    If GrhData(grh_index).FileNum = 0 Then Exit Sub
        
    'Modified by Augusto José Rando
    'Simplier function - according to basic ORE engine
    If h_centered Then
        If GrhData(Grh.grhindex).TileWidth <> 1 Then
            screen_x = screen_x - Int(GrhData(Grh.grhindex).TileWidth * (32 \ 2)) + 32 \ 2
        End If
    End If
    
    If v_centered Then
        If GrhData(Grh.grhindex).TileHeight <> 1 Then
            screen_y = screen_y - Int(GrhData(Grh.grhindex).TileHeight * 32) + 32
        End If
    End If
    
    'Draw it to device
    Device_Box_Textured_Render grh_index, _
        screen_x, screen_y, _
        GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight, _
        rgb_list(), _
        GrhData(grh_index).sX, GrhData(grh_index).sY, _
        alpha_blend, _
        Grh.angle

End Sub

Public Sub Trueno_Render()

    Static src_rect As RECT
    Static dest_rect As RECT
    Static temp_verts(3) As TLVERTEX
    Static light_value(0 To 3) As Long
    
    light_value(0) = 1179010815
    light_value(1) = 2017871615
    light_value(2) = 1179010815
    light_value(3) = 910575359

    'Set up the source rectangle
    With src_rect
        .bottom = 0 + frmMain.MainViewPic.ScaleHeight
        .Left = 0
        .Right = 0 + frmMain.MainViewPic.ScaleWidth
        .Top = 0
    End With
                
    'Set up the destination rectangle
    With dest_rect
        .bottom = 0 + frmMain.MainViewPic.ScaleHeight
        .Left = 0
        .Right = 0 + frmMain.MainViewPic.ScaleWidth
        .Top = 0
    End With
    
    'Set up the TempVerts(3) vertices
    Geometry_Create_Box temp_verts(), dest_rect, src_rect, light_value(), 0, 0, 0
    
    'Set Textures
    D3DDevice.SetTexture 0, Nothing
    
    'If alpha_blend Then
    '   'Set Rendering for alphablending
    '    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
    '    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    'End If
    
    'Draw the triangles that make up our square Textures
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0))
    
    'If alpha_blend Then
    '    'Set Rendering for colokeying
    '    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    '    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    'End If

End Sub

Public Sub DrawPJ(ByVal Index As Byte)
    Dim cColor As Long
    
    Index = Index - 1
    If cPJ(Index).nombre <> "" Then
        frmPanelAccount.lblAccData(Index + 1).Caption = cPJ(Index).nombre
        Select Case cPJ(Index).color
            Case 1 'Gris
                cColor = RGB(128, 128, 128)
            Case 2 'Azul
                cColor = RGB(0, 0, 255)
            Case 3 'Rojo
                cColor = RGB(255, 0, 0)
            Case 4 'Naranja
                cColor = RGB(255, 128, 0)
            Case 5 'Verde
                cColor = RGB(0, 225, 0)
        End Select
        frmPanelAccount.lblAccData(Index + 1).ForeColor = cColor
    End If
    If cPJ(Index).nombre = "" Then Exit Sub
    
    On Error Resume Next
    
    Dim i As Integer
    
    Dim init_x As Integer
    Dim init_y As Integer
    Dim tempito(3) As Long
    Dim grhtemp As Grh
    Static re As RECT
    
    re.Left = 0
    re.Top = 0
    re.bottom = 80
    re.Right = 76
    
    init_x = 25
    init_y = 40
    
    tempito(0) = &HFFFFFFFF
    tempito(1) = &HFFFFFFFF
    tempito(2) = &HFFFFFFFF
    tempito(3) = &HFFFFFFFF
    
    D3DDevice.BeginScene
    D3DDevice.Clear 0, ByVal 0&, D3DCLEAR_TARGET, 0, 0, 0
   
        If cPJ(Index).body = 255 Then
            Draw_GrhIndex 20206, 0, 0, &HFFFD7E
        End If
    
        If cPJ(Index).Shield = 32 Then
            Draw_GrhIndex 20128, 0, 0, &HFF
        End If
    
        If cPJ(Index).Weapon = 16 Then
            Draw_GrhIndex 20128, 0, 0, &HFFCC33
        ElseIf cPJ(Index).Weapon = 17 Then
            Draw_GrhIndex 20133, -24, -24, &HFF331
        ElseIf cPJ(Index).Weapon = 23 Then
            Draw_GrhIndex 20152, 0, 0, &HFF0000
        ElseIf cPJ(Index).Weapon = 24 Then
            Draw_GrhIndex 20185, 0, 0, -65536
        ElseIf cPJ(Index).Weapon = 25 Then
            Draw_GrhIndex 20155, 0, 0, &HFF0000
        ElseIf cPJ(Index).Weapon = 26 Then
            Draw_GrhIndex 20151, 0, 0, &HFFFF00
        ElseIf cPJ(Index).Weapon = 27 Then
            Draw_GrhIndex 20147, 0, 0, &HFF
        ElseIf cPJ(Index).Weapon = 28 Then
            Draw_GrhIndex 20146, 0, 0, &H6B1B
        ElseIf cPJ(Index).Weapon = 29 Then
            Draw_GrhIndex 20200, 0, 0, &HCCFF33
        End If
   
        If cPJ(Index).body <> 0 Then
            grhtemp.grhindex = BodyData(cPJ(Index).body).Walk(3).grhindex
            Grh_Render grhtemp, init_x, init_y, tempito
        End If
        
        If cPJ(Index).Head <> 0 Then
            grhtemp.grhindex = HeadData(cPJ(Index).Head).Head(3).grhindex
            Grh_Render grhtemp, init_x + BodyData(cPJ(Index).body).HeadOffset.x, init_y + BodyData(cPJ(Index).body).HeadOffset.y, tempito
        End If
        
        If cPJ(Index).Casco <> 0 Then
            grhtemp.grhindex = CascoAnimData(cPJ(Index).Casco).Head(3).grhindex
            Grh_Render grhtemp, init_x + BodyData(cPJ(Index).body).HeadOffset.x, init_y + BodyData(cPJ(Index).body).HeadOffset.y, tempito
        End If
        
        If cPJ(Index).Weapon <> 0 Then
            grhtemp.grhindex = WeaponAnimData(cPJ(Index).Weapon).WeaponWalk(3).grhindex
            Grh_Render grhtemp, init_x, init_y, tempito
        End If
        
        If cPJ(Index).Shield <> 0 Then
            grhtemp.grhindex = ShieldAnimData(cPJ(Index).Shield).ShieldWalk(3).grhindex
            Grh_Render grhtemp, init_x, init_y, tempito
        End If
    
    D3DDevice.EndScene
    D3DDevice.Present re, ByVal 0&, frmPanelAccount.picChar(Index).hwnd, ByVal 0&

End Sub

Private Sub RoofAlphaCalculate(roofrgb_list() As Long)
    Dim color As D3DCOLORVALUE
    Static AlphaY As Integer
    Static last_tick As Long
        If IIf(MapData(UserPos.x, UserPos.y).Trigger = 1 Or _
                MapData(UserPos.x, UserPos.y).Trigger = 2 Or _
                MapData(UserPos.x, UserPos.y).Trigger = 7 Or _
                MapData(UserPos.x, UserPos.y).Trigger = 4 Or _
                MapData(UserPos.x, UserPos.y).Trigger >= 20, True, False) Then
            If AlphaY > 0 Then
                If GetTickCount - last_tick >= 18 Then
                    AlphaY = AlphaY - 5
                    last_tick = GetTickCount
                End If
            End If
        Else
            If AlphaY < 255 Then
                If GetTickCount - last_tick >= 18 Then
                    AlphaY = AlphaY + 5
                    last_tick = GetTickCount
                End If
            End If
        End If
        
        color = AmbientColor
        color.A = AlphaY
        D3DColorToRgbList roofrgb_list(), color
End Sub
Private Sub RoofAlphaCalculateToAlpha(roofrgb_list() As Long, ByVal CharIndex As Integer)
    Dim color As D3DCOLORVALUE
    Static state As Byte
    
        If state Then
            If charlist(CharIndex).AlphaX > 0 Then
                charlist(CharIndex).AlphaX = charlist(CharIndex).AlphaX - (timerTicksPerFrame * 6.7)
                charlist(CharIndex).last_tick = GetTickCount
            Else
                state = 0
                charlist(CharIndex).AlphaX = 0
            End If
        Else
            If charlist(CharIndex).AlphaX < 255 Then
                charlist(CharIndex).AlphaX = charlist(CharIndex).AlphaX + (timerTicksPerFrame * 6.7)
                charlist(CharIndex).last_tick = GetTickCount
            Else
                state = 1
                charlist(CharIndex).AlphaX = 255
            End If
        End If
        
    color = AmbientColor
    color.A = charlist(CharIndex).AlphaX
    D3DColorToRgbList roofrgb_list(), color
End Sub
Public Sub D3DColorToRgbList(rgb_list() As Long, color As D3DCOLORVALUE)
    rgb_list(0) = D3DColorARGB(color.A, color.r, color.g, color.b)
    rgb_list(1) = rgb_list(0)
    rgb_list(2) = rgb_list(0)
    rgb_list(3) = rgb_list(0)
End Sub

Public Sub Long_To_RGB_List(rgb_list() As Long, long_color As Long)
    rgb_list(0) = long_color
    rgb_list(1) = rgb_list(0)
    rgb_list(2) = rgb_list(0)
    rgb_list(3) = rgb_list(0)
End Sub
Function textwidth(Texto As String, ByVal font_index As Integer, Optional multi As Boolean = False) As Integer

Dim A As Integer, b As Integer, d As Integer, e As Integer, f As Integer
Dim graf As Grh

If multi = False Then
    For A = 1 To Len(Texto)
        b = Asc(mid(Texto, A, 1))
        graf.grhindex = font_types(1).ascii_code(b)
        If (b <> 32) And (b <> 5) And (b <> 129) And (b <> 9) And (b <> 4) And (b <> 255) And (b <> 2) Then
            textwidth = textwidth + GrhData(GrhData(graf.grhindex).Frames(1)).pixelWidth '+ 1
        Else
            textwidth = textwidth + 4
        End If
    Next A
Else
    e = 0
    f = 0
    For A = 1 To Len(Texto)
        b = Asc(mid(Texto, A, 1))
        graf.grhindex = font_types(1).ascii_code(b)
        If b = 32 Or b = 13 Then
            If e >= 20 Then 'reemplazar por lo que os plazca
                f = f + 1
                e = 0
                d = 0
            Else
                If b = 32 Then d = d + 4
            End If
        Else
            If graf.grhindex > 12 Then
                d = d + GrhData(GrhData(graf.grhindex).Frames(1)).pixelWidth '+ 1
                If d > textwidth Then textwidth = d
            End If
        End If
        e = e + 1
    Next A
End If

End Function
Public Function Light_Remove(ByVal light_index As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    'Make sure it's a legal index
    If Light_Check(light_index) Then
        Light_Destroy light_index
        Light_Remove = True
    End If
End Function

Public Function Light_Color_Value_Get(ByVal light_index As Long, ByRef color_value As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/28/2003
'
'*****************************************************************
    'Make sure it's a legal index
    If Light_Check(light_index) Then
        color_value = light_list(light_index).color
        Light_Color_Value_Get = True
    End If
End Function

Public Function Light_Create(ByVal map_x As Integer, ByVal map_y As Integer, Optional ByVal color_value As Long = &HFFFFFF, _
                            Optional ByVal range As Byte = 1, Optional ByVal id As Long) As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Returns the light_index if successful, else 0
'Edited by Juan Martín Sotuyo Dodero
'**************************************************************
    If InMapBounds(map_x, map_y) Then
        'Make sure there is no light in the given map pos
        If Map_Light_Get(map_x, map_y) <> 0 Then
            Light_Create = 0
            Exit Function
        End If
        Light_Create = Light_Next_Open
        
        Dim r As Integer, g As Integer, b As Integer
        General_Long_Color_to_RGB color_value, r, g, b
        color_value = D3DColorXRGB(r, g, b)
        
        Light_Make Light_Create, map_x, map_y, color_value, range, id
    End If
End Function
Public Function Map_Light_Get(ByVal map_x As Integer, ByVal map_y As Integer) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/20/2003
'Checks to see if a tile position has a light_index and return it
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    'We start from the back, to get the last light to be placed on the tile first
    If light_last = 0 Then Exit Function
    
    loopc = light_last
    Do Until light_list(loopc).map_x = map_x And light_list(loopc).map_y = map_y
        If loopc = 0 Then
            Map_Light_Get = 0
            Exit Function
        End If
        loopc = loopc - 1
        If loopc = 0 Then Exit Function
    Loop
    
    Map_Light_Get = loopc
Exit Function
ErrorHandler:
    Map_Light_Get = 0
End Function
Public Function Light_Move(ByVal light_index As Long, ByVal map_x As Integer, ByVal map_y As Integer) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Returns true if successful, else false
'**************************************************************
    'Make sure it's a legal CharIndex
    If Light_Check(light_index) Then
        'Make sure it's a legal move
        If InMapBounds(map_x, map_y) Then
        
            'Move it
            Light_Erase light_index
            light_list(light_index).map_x = map_x
            light_list(light_index).map_y = map_y
    
            Light_Move = True
            
        End If
    End If
End Function

Public Function Light_Move_By_Head(ByVal light_index As Long, ByVal heading As Byte) As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 15/05/2002
'Returns true if successful, else false
'**************************************************************
    Dim map_x As Integer
    Dim map_y As Integer
    Dim nX As Integer
    Dim nY As Integer
    
    'Check for valid heading
    If heading < 1 Or heading > 8 Then
        Light_Move_By_Head = False
        Exit Function
    End If

    'Make sure it's a legal CharIndex
    If Light_Check(light_index) Then
    
        map_x = light_list(light_index).map_x
        map_y = light_list(light_index).map_y
        
        nX = map_x
        nY = map_y
        
        'Convert_Heading_to_Direction heading, nX, nY
        
        'Make sure it's a legal move
        If InMapBounds(nX, nY) Then
        
            'Move it
            Light_Erase light_index

            light_list(light_index).map_x = nX
            light_list(light_index).map_y = nY
    
            Light_Move_By_Head = True
            
        End If
    End If
End Function

Private Sub Light_Make(ByVal light_index As Long, ByVal map_x As Integer, ByVal map_y As Integer, ByVal rgb_value As Long, _
                        ByVal range As Long, Optional ByVal id As Long)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
    'Update array size
    If light_index > light_last Then
        light_last = light_index
        ReDim Preserve light_list(1 To light_last)
    End If
    light_count = light_count + 1
    
    'Make active
    light_list(light_index).active = True
    
    light_list(light_index).map_x = map_x
    light_list(light_index).map_y = map_y
    light_list(light_index).color = rgb_value
    light_list(light_index).range = range
    light_list(light_index).id = id
End Sub

Private Function Light_Check(ByVal light_index As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'**************************************************************
    'check light_index
    If light_index > 0 And light_index <= light_last Then
        If light_list(light_index).active Then
            Light_Check = True
        End If
    End If
End Function

Public Sub Light_Render_All()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim loop_counter As Long
            
    For loop_counter = 1 To light_count
        
        If light_list(loop_counter).active Then
            Light_Render loop_counter
        End If
    
    Next loop_counter
End Sub

Private Sub Light_Render(ByVal light_index As Long)
'***************************************'
'Author: Juan Martín Sotuyo Dodero
'Last modified: 11/11/2004
'Renders a light
'***************************************'
    Dim min_x As Integer
    Dim min_y As Integer
    Dim max_x As Integer
    Dim max_y As Integer
    Dim x As Integer
    Dim y As Integer
    Dim color As Long
    Dim light_trigger As Integer
    
    'Set up light borders
    min_x = light_list(light_index).map_x - light_list(light_index).range
    min_y = light_list(light_index).map_y - light_list(light_index).range
    max_x = light_list(light_index).map_x + light_list(light_index).range
    max_y = light_list(light_index).map_y + light_list(light_index).range
    
    light_trigger = MapData(light_list(light_index).map_x, light_list(light_index).map_y).Trigger
    
    'Set color
    color = light_list(light_index).color
    
    'Arrange corners
    'NE
    If InMapBounds(min_x, min_y) Then
        If MapData(min_x, min_y).Trigger = light_trigger Then _
            MapData(min_x, min_y).light_value(2) = color
    End If
    'NW
    If InMapBounds(max_x, min_y) Then
        If MapData(max_x, min_y).Trigger = light_trigger Then _
            MapData(max_x, min_y).light_value(0) = color
    End If
    'SW
    If InMapBounds(max_x, max_y) Then
        If MapData(max_x, max_y).Trigger = light_trigger Then _
            MapData(max_x, max_y).light_value(1) = color
    End If
    'SE
    If InMapBounds(min_x, max_y) Then
        If MapData(min_x, max_y).Trigger = light_trigger Then _
            MapData(min_x, max_y).light_value(3) = color
    End If
    
    'Arrange borders
    'Upper border
    For x = min_x + 1 To max_x - 1
        If InMapBounds(x, min_y) Then
            If MapData(x, min_y).Trigger = light_trigger Then
                MapData(x, min_y).light_value(0) = color
                MapData(x, min_y).light_value(2) = color
            End If
        End If
    Next x
    
    'Lower border
    For x = min_x + 1 To max_x - 1
        If InMapBounds(x, max_y) Then
            If MapData(x, max_y).Trigger = light_trigger Then
                MapData(x, max_y).light_value(1) = color
                MapData(x, max_y).light_value(3) = color
            End If
        End If
    Next x
    
    'Left border
    For y = min_y + 1 To max_y - 1
        If InMapBounds(min_x, y) Then
            If MapData(min_x, y).Trigger = light_trigger Then
                MapData(min_x, y).light_value(2) = color
                MapData(min_x, y).light_value(3) = color
            End If
        End If
    Next y
    
    'Right border
    For y = min_y + 1 To max_y - 1
        If InMapBounds(max_x, y) Then
            If MapData(max_x, y).Trigger = light_trigger Then
                MapData(max_x, y).light_value(0) = color
                MapData(max_x, y).light_value(1) = color
            End If
        End If
    Next y
    
    'Set the inner part of the light
    For x = min_x + 1 To max_x - 1
        For y = min_y + 1 To max_y - 1
            If InMapBounds(x, y) Then
                If MapData(x, y).Trigger = light_trigger Then
                    MapData(x, y).light_value(0) = color
                    MapData(x, y).light_value(1) = color
                    MapData(x, y).light_value(2) = color
                    MapData(x, y).light_value(3) = color
                End If
            End If
        Next y
    Next x
    
End Sub

Private Function Light_Next_Open() As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    If light_last = 0 Then
        Light_Next_Open = 1
        Exit Function
    End If
    
    loopc = 1
    Do Until light_list(loopc).active = False
        If loopc = light_last Then
            Light_Next_Open = light_last + 1
            Exit Function
        End If
        loopc = loopc + 1
    Loop
    
    Light_Next_Open = loopc

Exit Function

ErrorHandler:

End Function

Public Function Light_Find(ByVal id As Long) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'Find the index related to the handle
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim loopc As Long
    
    loopc = 1
    Do Until light_list(loopc).id = id
        If loopc = light_last Then
            Light_Find = 0
            Exit Function
        End If
        loopc = loopc + 1
    Loop
    
    Light_Find = loopc
Exit Function
ErrorHandler:
    Light_Find = 0
End Function

Public Function Light_Remove_All() As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    Dim Index As Long
    
    For Index = 1 To light_last
        'Make sure it's a legal index
        If Light_Check(Index) Then
            Light_Destroy Index
        End If
    Next Index
    
    Light_Remove_All = True
End Function

Private Sub Light_Destroy(ByVal light_index As Long)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim temp As Light
    
    Light_Erase light_index
    
    light_list(light_index) = temp
    
    'Update array size
    If light_index = light_last Then
        Do Until light_list(light_last).active
            light_last = light_last - 1
            If light_last = 0 Then
                light_count = 0
                Exit Sub
            End If
        Loop
        ReDim Preserve light_list(1 To light_last)
    End If
    light_count = light_count - 1
End Sub

Private Sub Light_Erase(ByVal light_index As Long)
'***************************************'
'Author: Juan Martín Sotuyo Dodero
'Last modified: 3/31/2003
'Correctly erases a light
'***************************************'
    Dim min_x As Integer
    Dim min_y As Integer
    Dim max_x As Integer
    Dim max_y As Integer
    Dim x As Integer
    Dim y As Integer
    Dim light_trigger As Integer
    
    'Set up light borders
    min_x = light_list(light_index).map_x - light_list(light_index).range
    min_y = light_list(light_index).map_y - light_list(light_index).range
    max_x = light_list(light_index).map_x + light_list(light_index).range
    max_y = light_list(light_index).map_y + light_list(light_index).range
    light_trigger = MapData(light_list(light_index).map_x, light_list(light_index).map_y).Trigger
    
    'Arrange corners
    'NE
    If InMapBounds(min_x, min_y) Then
        If MapData(min_x, min_y).Trigger = light_trigger Then _
            MapData(min_x, min_y).light_value(2) = map_base_light
    End If
    'NW
    If InMapBounds(max_x, min_y) Then
        If MapData(max_x, min_y).Trigger = light_trigger Then _
            MapData(max_x, min_y).light_value(0) = map_base_light
    End If
    'SW
    If InMapBounds(max_x, max_y) Then
        If MapData(max_x, max_y).Trigger = light_trigger Then _
            MapData(max_x, max_y).light_value(1) = map_base_light
    End If
    'SE
    If InMapBounds(min_x, max_y) Then
        If MapData(min_x, max_y).Trigger = light_trigger Then _
            MapData(min_x, max_y).light_value(3) = map_base_light
    End If
    
    'Arrange borders
    'Upper border
    For x = min_x + 1 To max_x - 1
        If InMapBounds(x, min_y) Then
            If MapData(x, min_y).Trigger = light_trigger Then
                MapData(x, min_y).light_value(0) = map_base_light
                MapData(x, min_y).light_value(2) = map_base_light
            End If
        End If
    Next x
    
    'Lower border
    For x = min_x + 1 To max_x - 1
        If InMapBounds(x, max_y) Then
            If MapData(x, max_y).Trigger = light_trigger Then
                MapData(x, max_y).light_value(1) = map_base_light
                MapData(x, max_y).light_value(3) = map_base_light
            End If
        End If
    Next x
    
    'Left border
    For y = min_y + 1 To max_y - 1
        If InMapBounds(min_x, y) Then
            If MapData(min_x, y).Trigger = light_trigger Then
                MapData(min_x, y).light_value(2) = map_base_light
                MapData(min_x, y).light_value(3) = map_base_light
            End If
        End If
    Next y
    
    'Right border
    For y = min_y + 1 To max_y - 1
        If InMapBounds(max_x, y) Then
            If MapData(max_x, y).Trigger = light_trigger Then
                MapData(max_x, y).light_value(0) = map_base_light
                MapData(max_x, y).light_value(1) = map_base_light
            End If
        End If
    Next y
    
    'Set the inner part of the light
    For x = min_x + 1 To max_x - 1
        For y = min_y + 1 To max_y - 1
            If InMapBounds(x, y) Then
                If MapData(x, y).Trigger = light_trigger Then
                    MapData(x, y).light_value(0) = map_base_light
                    MapData(x, y).light_value(1) = map_base_light
                    MapData(x, y).light_value(2) = map_base_light
                    MapData(x, y).light_value(3) = map_base_light
                End If
            End If
        Next y
    Next x
End Sub


Public Sub setup_ambient()
Select Case Hour(time)
'Noche 87, 61, 43
Case 0
frmMain.imgClima.Picture = LoadInterface("00")
    luz_dia(0).r = 233 '87
    luz_dia(0).g = 215 '61
    luz_dia(0).b = 255 '43
Exit Sub
 
Case 1
frmMain.imgClima.Picture = LoadInterface("01")
    luz_dia(1).r = 87
    luz_dia(1).g = 61
    luz_dia(1).b = 43
Exit Sub
 
Case 2
frmMain.imgClima.Picture = LoadInterface("02")
    luz_dia(2).r = 87
    luz_dia(2).g = 61
    luz_dia(2).b = 43
Exit Sub
 
Case 3
frmMain.imgClima.Picture = LoadInterface("03")
    luz_dia(3).r = 87
    luz_dia(3).g = 61
    luz_dia(3).b = 43
Exit Sub
'4 am 124,117,91
Case 4
frmMain.imgClima.Picture = LoadInterface("04")
    luz_dia(4).r = 124
    luz_dia(4).g = 127
    luz_dia(4).b = 91
Exit Sub
'5,6 am 143,137,135
Case 5
frmMain.imgClima.Picture = LoadInterface("05")
    luz_dia(5).r = 143
    luz_dia(5).g = 137
    luz_dia(5).b = 135
Exit Sub
 
Case 6
frmMain.imgClima.Picture = LoadInterface("06")
    luz_dia(6).r = 143
    luz_dia(6).g = 137
    luz_dia(6).b = 135
Exit Sub
 
'7 am 212,205,207
Case 7
frmMain.imgClima.Picture = LoadInterface("07")
    luz_dia(7).r = 212
    luz_dia(7).g = 205
    luz_dia(7).b = 207
Exit Sub
 
Case 8
frmMain.imgClima.Picture = LoadInterface("08")
    luz_dia(8).r = 212
    luz_dia(8).g = 205
    luz_dia(8).b = 207
Exit Sub
 
Case 9
frmMain.imgClima.Picture = LoadInterface("09")
    luz_dia(9).r = 212
    luz_dia(9).g = 205
    luz_dia(9).b = 207
Exit Sub
 
Case 10
frmMain.imgClima.Picture = LoadInterface("10")
    luz_dia(10).r = 212
    luz_dia(10).g = 205
    luz_dia(10).b = 207
Exit Sub
 
Case 11
frmMain.imgClima.Picture = LoadInterface("11")
    luz_dia(11).r = 212
    luz_dia(11).g = 205
    luz_dia(11).b = 207
Exit Sub
 
Case 12
frmMain.imgClima.Picture = LoadInterface("12")
    luz_dia(12).r = 212
    luz_dia(12).g = 205
    luz_dia(12).b = 207
Exit Sub
'Dia 255, 255, 255
 
Case 13
frmMain.imgClima.Picture = LoadInterface("13")
    luz_dia(13).r = 255
    luz_dia(13).g = 255
    luz_dia(13).b = 255
Exit Sub
'Medio Dia 255, 200, 255
Case 14
frmMain.imgClima.Picture = LoadInterface("14")
    luz_dia(14).r = 255
    luz_dia(14).g = 250
    luz_dia(14).b = 255
Exit Sub
 
Case 15
frmMain.imgClima.Picture = LoadInterface("15")
    luz_dia(15).r = 255
    luz_dia(15).g = 240
    luz_dia(15).b = 255
Exit Sub
 
Case 16
frmMain.imgClima.Picture = LoadInterface("16")
    luz_dia(16).r = 255
    luz_dia(16).g = 230
    luz_dia(16).b = 255
Exit Sub
 
'17/18 0, 100, 255
Case 17
frmMain.imgClima.Picture = LoadInterface("17")
    luz_dia(17).r = 230
    luz_dia(17).g = 230
    luz_dia(17).b = 255
Exit Sub
'18/19 0, 100, 255
Case 18
frmMain.imgClima.Picture = LoadInterface("18")
    luz_dia(18).r = 230
    luz_dia(18).g = 230
    luz_dia(18).b = 255
Exit Sub
'19/20 156, 142, 83
Case 19
frmMain.imgClima.Picture = LoadInterface("19")
    luz_dia(19).r = 156
    luz_dia(19).g = 142
    luz_dia(19).b = 83
Exit Sub
 
Case 20
frmMain.imgClima.Picture = LoadInterface("20")
    luz_dia(20).r = 87
    luz_dia(20).g = 61
    luz_dia(20).b = 43
Exit Sub
 
Case 21
frmMain.imgClima.Picture = LoadInterface("21")
    luz_dia(21).r = 87
    luz_dia(21).g = 61
    luz_dia(21).b = 43
Exit Sub
 
Case 22
frmMain.imgClima.Picture = LoadInterface("22")
    luz_dia(22).r = 87
    luz_dia(22).g = 61
    luz_dia(22).b = 43
Exit Sub
 
Case 23
frmMain.imgClima.Picture = LoadInterface("23")
    luz_dia(23).r = 87
    luz_dia(23).g = 61
    luz_dia(23).b = 43
Exit Sub
 
Case 24
frmMain.imgClima.Picture = LoadInterface("00")
    luz_dia(24).r = 87
    luz_dia(24).g = 61
    luz_dia(24).b = 43
Exit Sub
End Select
End Sub
 
Public Function change_day_effect(ByVal o_r As Long, ByVal og As Long, ByVal ob As Long, ByVal r_ As Long, ByVal g_ As Long, ByVal b_ As Long) As Long
Dim r As Long
Dim g As Long
Dim b As Long
r = o_r
g = og
b = ob
    If o_r > r_ Then
        r = r - 1
    ElseIf o_r < r_ Then
        r = r + 1
    End If

    If og > g_ Then
        g = g - 1
    ElseIf og < r Then
        g = g + 1
    End If

    If ob > b_ Then
        b = b - 1
    ElseIf og < r Then
        b = b + 1
    End If
If r > 255 Then r = 255
If g > 255 Then g = 255
If b > 255 Then b = 255
day_r_old = r
day_g_old = g
day_b_old = b

If r < 0 Or g < 0 Or b < 0 Then r = 1: g = 1: b = 1
change_day_effect = ARGB(r, g, b, 255)
End Function


Public Function ARGB(ByVal r As Long, ByVal g As Long, ByVal b As Long, ByVal A As Long) As Long
    Dim c As Long
    If A > 127 Then
        A = A - 128
        c = A * 2 ^ 24 Or &H80000000
        c = c Or r * 2 ^ 16
        c = c Or g * 2 ^ 8
        c = c Or b
    Else
        c = A * 2 ^ 24
        c = c Or r * 2 ^ 16
        c = c Or g * 2 ^ 8
        c = c Or b
    End If
    ARGB = c
End Function
