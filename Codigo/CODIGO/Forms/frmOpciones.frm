VERSION 5.00
Begin VB.Form frmOpciones 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   8025
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7035
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H80000004&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000001&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmOpciones.frx":0000
   ScaleHeight     =   8025
   ScaleWidth      =   7035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Caption         =   "Arenas"
      Height          =   495
      Left            =   120
      TabIndex        =   36
      Top             =   4680
      Width           =   3255
   End
   Begin VB.CheckBox Check7 
      Caption         =   "Invertir botones del Mouse."
      Height          =   255
      Left            =   3840
      TabIndex        =   34
      Top             =   4680
      Width           =   2655
   End
   Begin VB.Frame Frame4 
      Caption         =   "Opciones de Mouse"
      ForeColor       =   &H8000000D&
      Height          =   2655
      Left            =   3840
      TabIndex        =   28
      Top             =   5160
      Width           =   2775
      Begin VB.CheckBox Check3 
         Caption         =   "Encendido"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txtMSens 
         Height          =   285
         Left            =   1200
         TabIndex        =   30
         Text            =   "10"
         Top             =   960
         Width           =   255
      End
      Begin VB.HScrollBar scrSens 
         Height          =   255
         Left            =   120
         Max             =   20
         Min             =   1
         TabIndex        =   29
         Top             =   1320
         Value           =   10
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Sensibilidad del Mouse"
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Apariencia y perfomance"
      ForeColor       =   &H8000000D&
      Height          =   2985
      Left            =   3600
      TabIndex        =   24
      Top             =   120
      Width           =   3285
      Begin VB.CheckBox Check4 
         Caption         =   "Deshabilitar transparencias"
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   520
         Width           =   2295
      End
      Begin VB.ListBox Interfaces 
         Height          =   1635
         ItemData        =   "frmOpciones.frx":2A6DB
         Left            =   120
         List            =   "frmOpciones.frx":2A71B
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   26
         Top             =   960
         Width           =   2895
      End
      Begin VB.CheckBox MapName 
         Caption         =   "Nombre en el Mapa"
         Height          =   285
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   2715
      End
      Begin VB.Label Label4 
         Caption         =   "Skins instalados"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   27
         Top             =   720
         Width           =   2925
      End
   End
   Begin VB.Frame frmIdioma 
      Caption         =   "Idíoma"
      ForeColor       =   &H8000000D&
      Height          =   705
      Left            =   120
      TabIndex        =   21
      Top             =   0
      Width           =   3255
      Begin VB.ComboBox Español 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmOpciones.frx":2A85F
         Left            =   180
         List            =   "frmOpciones.frx":2A869
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cerrar y Guardar"
      Height          =   360
      Left            =   120
      MousePointer    =   99  'Custom
      TabIndex        =   20
      Top             =   7560
      Width           =   3255
   End
   Begin VB.CommandButton cmdCustomKeys 
      Caption         =   "Configuracion de Controles"
      Height          =   360
      Left            =   120
      MousePointer    =   99  'Custom
      TabIndex        =   19
      Top             =   7080
      Width           =   3255
   End
   Begin VB.Frame Frame2 
      Caption         =   "Informaciòn"
      ForeColor       =   &H8000000D&
      Height          =   1695
      Left            =   120
      TabIndex        =   16
      Top             =   5280
      Width           =   3255
      Begin VB.CommandButton Command1 
         Caption         =   "Party"
         Height          =   375
         Left            =   240
         TabIndex        =   35
         Top             =   240
         Width           =   2775
      End
      Begin VB.CommandButton cmdAyuda 
         Caption         =   "¿Nesesítas Ayuda?"
         Height          =   345
         Left            =   180
         MousePointer    =   99  'Custom
         TabIndex        =   18
         Top             =   1080
         Width           =   2895
      End
      Begin VB.CommandButton cmdWeb 
         Caption         =   "Clanes"
         Height          =   345
         Index           =   1
         Left            =   180
         MousePointer    =   99  'Custom
         Picture         =   "frmOpciones.frx":2A87F
         TabIndex        =   17
         Top             =   720
         Width           =   2895
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Audio"
      ForeColor       =   &H8000000D&
      Height          =   3855
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   3255
      Begin VB.CheckBox Check2 
         Caption         =   "Efecto de navegacion"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   720
         Width           =   1935
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H80000004&
         Caption         =   "Efectos de sonido"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   1080
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H80000004&
         Caption         =   "Musica habilitada"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.TextBox txtMidi 
         Height          =   285
         Left            =   1920
         TabIndex        =   7
         Top             =   1440
         Width           =   345
      End
      Begin VB.HScrollBar Slider1 
         Height          =   315
         Index           =   0
         LargeChange     =   15
         Left            =   120
         Max             =   0
         Min             =   -4000
         SmallChange     =   2
         TabIndex        =   6
         Top             =   2160
         Width           =   2895
      End
      Begin VB.HScrollBar Slider1 
         Height          =   315
         Index           =   2
         LargeChange     =   15
         Left            =   120
         Max             =   0
         Min             =   -4000
         SmallChange     =   2
         TabIndex        =   5
         Top             =   2760
         Width           =   2895
      End
      Begin VB.HScrollBar Slideramb 
         Enabled         =   0   'False
         Height          =   315
         LargeChange     =   15
         Left            =   120
         Max             =   0
         Min             =   -4000
         SmallChange     =   2
         TabIndex        =   4
         Top             =   3360
         Width           =   2895
      End
      Begin VB.Label lblMidi 
         BackStyle       =   0  'Transparent
         Caption         =   "Reproducir Midi"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label lblBackMidi 
         Caption         =   "«"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1680
         TabIndex        =   12
         Top             =   1440
         Width           =   135
      End
      Begin VB.Label lblNextMidi 
         Caption         =   "»"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2400
         TabIndex        =   11
         Top             =   1440
         Width           =   135
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Volumen de musica"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   1920
         Width           =   2835
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Volumen de sonidos"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   2520
         Width           =   2865
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Volumen sonidos ambientales"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   3120
         Width           =   2865
      End
   End
   Begin VB.CommandButton cmdManual 
      Caption         =   "www.CoverAO.jimdo.com"
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   4200
      Width           =   2775
   End
   Begin VB.CommandButton cmdViewMap 
      Caption         =   "Mapa del Juego"
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   3720
      Width           =   2775
   End
   Begin VB.CommandButton cmdChangePassword 
      Caption         =   "Cambiar Contraseña"
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   3240
      Width           =   2775
   End
End
Attribute VB_Name = "frmOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Cover Online 0.11.6
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
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
'Kega Online is based on Baronsoft's VB6 Online RPG
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

Private Loading As Boolean
Dim MapaName As Boolean

Private Sub Check1_Click(Index As Integer)
If Check1 Then
        SwapMouseButton 1
    Else
        SwapMouseButton 0
    End If
    If Not Loading And Sounder Then _
        Call Audio.PlayWave(SND_CLICK)
    
    Select Case Index
        Case 0
            If Check1(0).value = vbUnchecked Then
                Music = 0
                Audio.MusicActivated = False
                Slider1(0).Enabled = False
            ElseIf Not Audio.MusicActivated Then  'Prevent the music from reloading
                Music = 1
                Audio.MusicActivated = True
                Slider1(0).Enabled = True
                Slider1(0).value = Audio.MusicVolume
            End If
        
        Case 1
            If Check1(1).value = vbUnchecked Then
                Sound = 0
                Audio.SoundActivated = False
                RainBufferIndex = 0
                frmMain.IsPlaying = PlayLoop.plnone
                Slider1(1).Enabled = False
            Else
                Sound = 1
                Audio.SoundActivated = True
                'Slider1(1).Enabled = True
'                Slider1(1).value = Audio.SoundVolume
            End If
            
        Case 2
            If Check1(2).value = vbUnchecked Then
                EffectSound = 0
                Audio.SoundEffectsActivated = False
            Else
                EffectSound = 1
                Audio.SoundEffectsActivated = True
            End If
    End Select
End Sub

Private Sub Check2_Click()
  'If SND_NAVEGANDO = True Then
  '          SND_NAVEGANDO = False 'esto tambien :P
  '      Else
  '          SND_NAVEGANDO = True
  '      End If
End Sub

Private Sub Check7_Click()
If Check7 Then
        SwapMouseButton 1
    Else
        SwapMouseButton 0
    End If
End Sub


Private Sub cmdAyuda_Click()
Mapa.Show
End Sub

Private Sub cmdViewMap_Click()
Mapa.Show
End Sub

Private Sub cmdWeb_Click(Index As Integer)
ComoFundarclan.Show
End Sub

Private Sub Command1_Click()
Party.Show
End Sub

Private Sub Command3_Click()
Arenas.Show
Unload Me
End Sub

Private Sub Interfaces_Click()
Select Case Interfaces
Case "Leales"
NumSkin = 1
frmMain.Picture = LoadPicture(App.Path & "\Interfaces\todo.jpg") ' Nombre de la interfaz del main.
frmMain.InvEqu.Picture = LoadPicture(App.Path & "\Interfaces\centroinventario.jpg") 'Nombre de la interfaz del inventario
frmMain.InvEqu.Picture = LoadPicture(App.Path & "\Interfaces\centrohechizos.jpg") ' Nombre de la interfaz del hechizo

Case "Zero"
NumSkin = 2
frmMain.Picture = LoadPicture(App.Path & "\Interfaces\interface.jpg") ' Nombre de la interfaz del main.
frmMain.InvEqu.Picture = LoadPicture(App.Path & "\Interfaces\e.jpg") 'Nombre de la interfaz del inventario
frmMain.InvEqu.Picture = LoadPicture(App.Path & "\Interfaces\a.jpg") ' Nombre de la interfaz del hechizo

Case "Volar de los Dragones"
NumSkin = 3
frmMain.Picture = LoadPicture(App.Path & "\Interfaces\q.jpg") ' Nombre de la interfaz del main.
frmMain.InvEqu.Picture = LoadPicture(App.Path & "\Interfaces\x.jpg") 'Nombre de la interfaz del inventario
frmMain.InvEqu.Picture = LoadPicture(App.Path & "\Interfaces\z.jpg") ' Nombre de la interfaz del hechizo

Case "Noche Roja"
NumSkin = 4
frmMain.Picture = LoadPicture(App.Path & "\Interfaces\1.jpg") ' Nombre de la interfaz del main.
frmMain.InvEqu.Picture = LoadPicture(App.Path & "\Interfaces\sa.jpg") 'Nombre de la interfaz del inventario
frmMain.InvEqu.Picture = LoadPicture(App.Path & "\Interfaces\as.jpg") ' Nombre de la interfaz del hechizo

Case "Sombra sigilosa"
NumSkin = 5
frmMain.Picture = LoadPicture(App.Path & "\Interfaces\5.jpg") ' Nombre de la interfaz del main.
frmMain.InvEqu.Picture = LoadPicture(App.Path & "\Interfaces\4.jpg") 'Nombre de la interfaz del inventario
frmMain.InvEqu.Picture = LoadPicture(App.Path & "\Interfaces\3.jpg") ' Nombre de la interfaz del hechizo

Case "Patrulla imperial"
NumSkin = 6
frmMain.Picture = LoadPicture(App.Path & "\Interfaces\8.jpg") ' Nombre de la interfaz del main.
frmMain.InvEqu.Picture = LoadPicture(App.Path & "\Interfaces\7.jpg") 'Nombre de la interfaz del inventario
frmMain.InvEqu.Picture = LoadPicture(App.Path & "\Interfaces\6.jpg") ' Nombre de la interfaz del hechizo

Case "Principal"
NumSkin = 7
frmMain.Picture = LoadPicture(App.Path & "\Interfaces\zx.jpg") ' Nombre de la interfaz del main.
frmMain.InvEqu.Picture = LoadPicture(App.Path & "\Interfaces\9.jpg") 'Nombre de la interfaz del inventario
frmMain.InvEqu.Picture = LoadPicture(App.Path & "\Interfaces\51.jpg") ' Nombre de la interfaz del hechizo

Case "Hordas del Caos"
NumSkin = 8
frmMain.Picture = LoadPicture(App.Path & "\Interfaces\zxc.jpg") ' Nombre de la interfaz del main.
frmMain.InvEqu.Picture = LoadPicture(App.Path & "\Interfaces\wq.jpg") 'Nombre de la interfaz del inventario
frmMain.InvEqu.Picture = LoadPicture(App.Path & "\Interfaces\qw.jpg") ' Nombre de la interfaz del hechizo

Case "Descarga electrica"
NumSkin = 9
frmMain.Picture = LoadPicture(App.Path & "\Interfaces\asd.jpg") ' Nombre de la interfaz del main.
frmMain.InvEqu.Picture = LoadPicture(App.Path & "\Interfaces\re.jpg") 'Nombre de la interfaz del inventario
frmMain.InvEqu.Picture = LoadPicture(App.Path & "\Interfaces\er.jpg") ' Nombre de la interfaz del hechizo

Case "Furia Orca"
NumSkin = 10
frmMain.Picture = LoadPicture(App.Path & "\Interfaces\asda.jpg") ' Nombre de la interfaz del main.
frmMain.InvEqu.Picture = LoadPicture(App.Path & "\Interfaces\77.jpg") 'Nombre de la interfaz del inventario
frmMain.InvEqu.Picture = LoadPicture(App.Path & "\Interfaces\78.jpg") ' Nombre de la interfaz del hechizo

Case "La Batalla del Arcangel"
NumSkin = 11
frmMain.Picture = LoadPicture(App.Path & "\Interfaces\asdasd.jpg") ' Nombre de la interfaz del main.
frmMain.InvEqu.Picture = LoadPicture(App.Path & "\Interfaces\qq.jpg") 'Nombre de la interfaz del inventario
frmMain.InvEqu.Picture = LoadPicture(App.Path & "\Interfaces\qwe.jpg") ' Nombre de la interfaz del hechizo

Case "La ira de demonio"
NumSkin = 12
frmMain.Picture = LoadPicture(App.Path & "\Interfaces\opo.jpg") ' Nombre de la interfaz del main.
frmMain.InvEqu.Picture = LoadPicture(App.Path & "\Interfaces\op.jpg") 'Nombre de la interfaz del inventario
frmMain.InvEqu.Picture = LoadPicture(App.Path & "\Interfaces\po.jpg") ' Nombre de la interfaz del hechizo

Case "Dungeon Dragon"
NumSkin = 13
frmMain.Picture = LoadPicture(App.Path & "\Interfaces\we.jpg") ' Nombre de la interfaz del main.
frmMain.InvEqu.Picture = LoadPicture(App.Path & "\Interfaces\asdg.jpg") 'Nombre de la interfaz del inventario
frmMain.InvEqu.Picture = LoadPicture(App.Path & "\Interfaces\asdf.jpg") ' Nombre de la interfaz del hechizo

Case "Aldemair"
NumSkin = 14
frmMain.Picture = LoadPicture(App.Path & "\Interfaces\61.jpg") ' Nombre de la interfaz del main.
frmMain.InvEqu.Picture = LoadPicture(App.Path & "\Interfaces\60.jpg") 'Nombre de la interfaz del inventario
frmMain.InvEqu.Picture = LoadPicture(App.Path & "\Interfaces\59.jpg") ' Nombre de la interfaz del hechizo

Case "Amanecer del imperio"
NumSkin = 15
frmMain.Picture = LoadPicture(App.Path & "\Interfaces\82.jpg") ' Nombre de la interfaz del main.
frmMain.InvEqu.Picture = LoadPicture(App.Path & "\Interfaces\80.jpg") 'Nombre de la interfaz del inventario
frmMain.InvEqu.Picture = LoadPicture(App.Path & "\Interfaces\81.jpg") ' Nombre de la interfaz del hechizo

Case "Arogath"
NumSkin = 16
frmMain.Picture = LoadPicture(App.Path & "\Interfaces\90.jpg") ' Nombre de la interfaz del main.
frmMain.InvEqu.Picture = LoadPicture(App.Path & "\Interfaces\92.jpg") 'Nombre de la interfaz del inventario
frmMain.InvEqu.Picture = LoadPicture(App.Path & "\Interfaces\91.jpg") ' Nombre de la interfaz del hechizo

Case "Atardecer"
NumSkin = 17
frmMain.Picture = LoadPicture(App.Path & "\Interfaces\100.jpg") ' Nombre de la interfaz del main.
frmMain.InvEqu.Picture = LoadPicture(App.Path & "\Interfaces\101.jpg") 'Nombre de la interfaz del inventario
frmMain.InvEqu.Picture = LoadPicture(App.Path & "\Interfaces\102.jpg") ' Nombre de la interfaz del hechizo

Case "La venganza del caballero republicano"
NumSkin = 18
frmMain.Picture = LoadPicture(App.Path & "\Interfaces\105.jpg") ' Nombre de la interfaz del main.
frmMain.InvEqu.Picture = LoadPicture(App.Path & "\Interfaces\104.jpg") 'Nombre de la interfaz del inventario
frmMain.InvEqu.Picture = LoadPicture(App.Path & "\Interfaces\103.jpg") ' Nombre de la interfaz del hechizo

Case "Original"
NumSkin = 19
frmMain.Picture = LoadPicture(App.Path & "\Interfaces\105.jpg") ' Nombre de la interfaz del main.
frmMain.InvEqu.Picture = LoadPicture(App.Path & "\Interfaces\108.jpg") 'Nombre de la interfaz del inventario
frmMain.InvEqu.Picture = LoadPicture(App.Path & "\Interfaces\109.jpg") ' Nombre de la interfaz del hechizo
End Select
End Sub

Private Sub Label3_Click()

End Sub

Private Sub MapName_Click()
     If MapName.value = 1 Then
            frmMain.Label2.Visible = False
            frmMain.lblMapName.Visible = True
        Else
            frmMain.Label2.Visible = True
            frmMain.lblMapName.Visible = False
        End If
End Sub

Private Sub chkop_Click(Index As Integer)

End Sub

Private Sub cmdCustomKeys_Click()
    If Not Loading And Sounder Then _
        Call Audio.PlayWave(SND_CLICK)
    Call frmCustomKeys.Show(vbModal, Me)
End Sub

Private Sub cmdManual_Click()
    If Not Loading And Sounder Then _
        Call Audio.PlayWave(SND_CLICK)
    Call ShellExecute(0, "Open", "www.CoverAO.jimdo.com", "", App.Path, 0)
End Sub

Private Sub cmdChangePassword_Click()
    Call frmNewPassword.Show(vbModal, Me)
End Sub


Private Sub Command2_Click()
Me.Visible = False
End Sub

Private Sub Form_Load()
    Loading = True      'Prevent sounds when setting check's values
    
    Slider1(0).min = -5000
    Slider1(0).max = 5000
    
    'Slider1(1).min = 0
'    Slider1(1).max = 100

    Loading = False     'Enable sounds when setting check's values
    
    
End Sub

Private Sub scrSens_Change()
MouseS = scrSens.value
Call General_Set_Mouse_Speed(MouseS)
txtMSens.Text = scrSens.value
End Sub

Private Sub Slider1_Change(Index As Integer)
    
    Select Case Index
        Case 0
            Audio.MusicVolume = Slider1(0).value
            VolumeMusic = Audio.MusicVolume
        Case 1
            Audio.SoundVolume = Slider1(1).value
            VolumeSound = Audio.SoundVolume
    End Select
End Sub

Private Sub Slider1_Scroll(Index As Integer)
    Select Case Index
        Case 0
            Audio.MusicVolume = Slider1(0).value
            VolumeMusic = Audio.MusicVolume
        Case 1
            Audio.SoundVolume = Slider1(1).value
            VolumeSound = Audio.SoundVolume
    End Select
End Sub
