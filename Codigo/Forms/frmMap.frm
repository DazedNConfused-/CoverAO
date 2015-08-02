VERSION 5.00
Begin VB.Form frmMap 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Mapa"
   ClientHeight    =   8760
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   584
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox MainViewPic 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   10500
      Left            =   0
      ScaleHeight     =   700
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   700
      TabIndex        =   0
      Top             =   0
      Width           =   10500
   End
End
Attribute VB_Name = "frmMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RenderMap As Boolean
Private Sub Form_Load()
    DibujarMiniMapadx8
End Sub
Sub DibujarMiniMapadx8()
On Error Resume Next
Dim map_x As Long, map_y As Long
    Dim Rectas As RECT
    Dim grhindex As Integer
    With Rectas
        .Left = 0
        .bottom = 700
        .Top = 0
        .Right = 700
    End With
    
    D3DDevice.BeginScene
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 0, 0
    
    For map_y = 0 To 99
        For map_x = 0 To 99
            grhindex = 0
            If MapData(map_x + 1, map_y + 1).Graphic(1).grhindex > 0 Then
                grhindex = MapData(map_x + 1, map_y + 1).Graphic(1).grhindex
                engine.Device_Box_Textured_Render_Advance grhindex, map_x * 6, map_y * 6, _
                    GrhData(grhindex).pixelWidth, GrhData(grhindex).pixelHeight, _
                    Default_RGB(), _
                    GrhData(grhindex).sX, GrhData(grhindex).sY, _
                    6, 6
            End If
            If MapData(map_x + 1, map_y + 1).Graphic(2).grhindex > 0 Then
                grhindex = MapData(map_x + 1, map_y + 1).Graphic(2).grhindex
                engine.Device_Box_Textured_Render_Advance grhindex, map_x * 6, map_y * 6, _
                    GrhData(grhindex).pixelWidth, GrhData(grhindex).pixelHeight, _
                    Default_RGB(), _
                    GrhData(grhindex).sX, GrhData(grhindex).sY, _
                    6, 6
            End If
        Next map_x
    Next map_y
    
    For map_y = 0 To 99
        For map_x = 0 To 99
            grhindex = 0
            If MapData(map_x + 1, map_y + 1).Graphic(3).grhindex > 0 Then
                grhindex = MapData(map_x + 1, map_y + 1).Graphic(3).grhindex
                DrawCapa3 grhindex, map_x * 6, map_y * 6, 6, 6
            End If
        Next map_x
    Next map_y
    
    For map_y = 0 To 99
        For map_x = 0 To 99
            grhindex = 0
            If MapData(map_x + 1, map_y + 1).Graphic(4).grhindex > 0 Then
                grhindex = MapData(map_x + 1, map_y + 1).Graphic(4).grhindex
                DrawTecho grhindex, map_x * 6, map_y * 6, 6, 6
            End If
        Next map_x
    Next map_y
    
    engine.Text_Render "Tu estas aquí.", UserPos.x * 6 - (engine.textwidth("Tu estas aquí.", ActualFont) / 2), UserPos.y * 6 - 7, Default_RGB, &HFFFFFFFF
    engine.Text_Render "·", UserPos.x * 6 - (engine.textwidth(".", ActualFont) / 2), UserPos.y * 6, Default_RGB, D3DColorXRGB(255, 0, 0)
    engine.Text_Render "Doble click para actualizar.", 0, 0, Default_RGB, &HFFFFFFFF
    
    D3DDevice.EndScene
    D3DDevice.Present Rectas, ByVal 0, MainViewPic.hWnd, ByVal 0
    'frmMain.Minimap.Refresh
   
End Sub


Private Sub MainViewPic_DblClick()
    DibujarMiniMapadx8
End Sub
Sub DrawTecho(ByVal grh_index As Integer, ByVal screen_x As Integer, ByVal screen_y As Integer, ByVal width As Integer, ByVal height As Integer)

    Dim tile_width As Integer
    Dim tile_height As Integer
    
    tile_width = GrhData(grh_index).TileWidth * width
    tile_height = GrhData(grh_index).TileHeight * height
    
    engine.Device_Box_Textured_Render_Advance grh_index, _
        screen_x - IIf(GrhData(grh_index).TileWidth > 1, (GrhData(grh_index).pixelWidth / 2) / 6, 0), _
        screen_y - IIf(GrhData(grh_index).TileHeight > 1, ((GrhData(grh_index).pixelHeight / 2) / 3), 0) - 5, _
        GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight, _
        Default_RGB(), _
        GrhData(grh_index).sX, GrhData(grh_index).sY, _
        tile_width, tile_height, 0, 0

End Sub
Sub DrawCapa3(ByVal grh_index As Integer, ByVal screen_x As Integer, ByVal screen_y As Integer, ByVal width As Integer, ByVal height As Integer)

    Dim tile_width As Integer
    Dim tile_height As Integer
    
    tile_width = GrhData(grh_index).TileWidth * width
    tile_height = GrhData(grh_index).TileHeight * height
    
    engine.Device_Box_Textured_Render_Advance grh_index, _
        screen_x - IIf(GrhData(grh_index).TileWidth > 1, (GrhData(grh_index).pixelWidth / 2) / 6, 0), _
        screen_y - IIf(GrhData(grh_index).TileHeight > 1, (GrhData(grh_index).pixelHeight / 2) / 3, 0) + 4, _
        GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight, _
        Default_RGB(), _
        GrhData(grh_index).sX, GrhData(grh_index).sY, _
        tile_width, tile_height, 0, 0

End Sub

