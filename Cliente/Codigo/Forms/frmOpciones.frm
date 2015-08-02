VERSION 5.00
Begin VB.Form frmOpciones 
   BackColor       =   &H80000004&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   7680
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
   ScaleHeight     =   7680
   ScaleWidth      =   7035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.HScrollBar scrSens 
      Height          =   345
      Left            =   4200
      Max             =   20
      Min             =   1
      TabIndex        =   33
      Top             =   5520
      Value           =   10
      Width           =   2025
   End
   Begin VB.TextBox txtMSens 
      Height          =   315
      Left            =   4200
      TabIndex        =   32
      Text            =   "10"
      Top             =   5160
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Luz en Mous"
      Height          =   375
      Left            =   3840
      TabIndex        =   31
      Top             =   4680
      Width           =   2775
   End
   Begin VB.Frame frmIdioma 
      Caption         =   "Lenguaje (Desabilitado)"
      Height          =   705
      Left            =   120
      TabIndex        =   28
      Top             =   0
      Width           =   3255
      Begin VB.ComboBox cmbLanguage 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmOpciones.frx":0000
         Left            =   180
         List            =   "frmOpciones.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "$66"
      Height          =   2985
      Left            =   3600
      TabIndex        =   23
      Top             =   120
      Width           =   3285
      Begin VB.CheckBox MapName 
         Caption         =   "Nombre en el Mapa"
         Height          =   285
         Left            =   180
         TabIndex        =   25
         Top             =   300
         Width           =   2715
      End
      Begin VB.ListBox lstSkin 
         Height          =   1635
         Left            =   120
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   24
         Top             =   960
         Width           =   2895
      End
      Begin VB.Label Label4 
         Caption         =   "Skins"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   27
         Top             =   720
         Width           =   2925
      End
      Begin VB.Label lblSkinData 
         BackStyle       =   0  'Transparent
         Caption         =   "Autor: No habilitado"
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   2640
         Width           =   2895
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cerrar y Guardar"
      Height          =   360
      Left            =   120
      MousePointer    =   99  'Custom
      TabIndex        =   22
      Top             =   6840
      Width           =   3255
   End
   Begin VB.CommandButton cmdCustomKeys 
      Caption         =   "Configurar Teclas"
      Height          =   360
      Left            =   120
      MousePointer    =   99  'Custom
      TabIndex        =   21
      Top             =   6360
      Width           =   3255
   End
   Begin VB.Frame Frame2 
      Caption         =   "Informaciòn"
      Height          =   1575
      Left            =   120
      TabIndex        =   17
      Top             =   4680
      Width           =   3255
      Begin VB.CommandButton cmdAyuda 
         Caption         =   "Ayuda"
         Height          =   345
         Left            =   180
         MousePointer    =   99  'Custom
         TabIndex        =   20
         Top             =   1080
         Width           =   2895
      End
      Begin VB.CommandButton cmdWeb 
         Caption         =   "Pagina Web"
         Height          =   345
         Index           =   0
         Left            =   180
         MousePointer    =   99  'Custom
         TabIndex        =   19
         Top             =   300
         Width           =   2895
      End
      Begin VB.CommandButton cmdWeb 
         Caption         =   "Foro"
         Height          =   345
         Index           =   1
         Left            =   180
         MousePointer    =   99  'Custom
         TabIndex        =   18
         Top             =   690
         Width           =   2895
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sonidos"
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
         TabIndex        =   30
         Top             =   960
         Width           =   1935
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H80000004&
         Caption         =   "Efectos de sonido"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   1200
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H80000004&
         Caption         =   "Sonido habilitado"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Value           =   1  'Checked
         Width           =   1575
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
         Left            =   2400
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
         Left            =   2265
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
         Left            =   2760
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
      Caption         =   "Manual de Argentum Online"
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
'Argentum Online 0.11.6
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

Private Loading As Boolean
Dim i As Integer
Dim MapaName As Boolean

Private Sub Check1_Click(Index As Integer)
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

Private Sub cmdViewMap_Click()
Call Audio.PlayWave(SND_CLICK)
Mapa.Show
Unload Me
End Sub

Private Sub Command1_Click()
LuzMouse = Not LuzMouse
    If Not LuzMouse Then
        Engine.Light_Remove (Engine.Light_Find(20))
    Else
        Engine.Light_Create UserPos.X + frmMain.MouseX \ 32 - frmMain.MainViewPic.ScaleWidth \ 64, UserPos.Y + frmMain.MouseY / 32 - frmMain.MainViewPic.ScaleHeight \ 64, D3DColorXRGB(255, 0, 255), 2, 20
    End If
Call Audio.PlayWave(SND_CLICK)
Unload Me
End Sub

Private Sub Command3_Click()

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
    Call ShellExecute(0, "Open", "http://ao.alkon.com.ar/aomanual/", "", App.Path, 0)
End Sub

Private Sub cmdChangePassword_Click()
    Call frmNewPassword.Show(vbModal, Me)
End Sub


Private Sub Command2_Click()
Call Audio.PlayWave(SND_CLICK)
Me.Visible = False
End Sub

Private Sub Form_Load()
    Loading = True      'Prevent sounds when setting check's values
    
    Slider1(0).min = -5000
    Slider1(0).max = 5000
    
    'Slider1(1).min = 0
'    Slider1(1).max = 100

    Loading = False     'Enable sounds when setting check's values
    
    
    If Not Transparencia(Me.hwnd, 0) = 0 Then
   
    MsgBox " Esta función Api no es soportada en Versiones" _
           & "anteriores a windows 2000", vbCritical
    Me.Show
Else

    Me.Enabled = False
    Me.Show
   
    For i = 0 To 255 Step 2
        'Maycolito
        Call Transparencia(Me.hwnd, i)
        DoEvents
    Next
   
    Me.Enabled = True
     
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 
If Not Transparencia(Me.hwnd, 0) = 0 Then
    Exit Sub
Else
    For i = 255 To 0 Step -3
        DoEvents
        Call Transparencia(Me.hwnd, i)
        DoEvents
    Next
   
End If
    
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
