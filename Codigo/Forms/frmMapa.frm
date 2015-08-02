VERSION 5.00
Begin VB.Form frmMapa 
   BackColor       =   &H80000001&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8775
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   8775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label lblTexto 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmMapa.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   5040
      Width           =   8175
   End
   Begin VB.Image imgMapDungeon 
      Height          =   4935
      Left            =   0
      Top             =   0
      Width           =   8775
   End
   Begin VB.Image imgMap 
      Height          =   4935
      Left            =   0
      Top             =   0
      Width           =   8775
   End
End
Attribute VB_Name = "frmMapa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************
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
'**************************************************************************

Option Explicit

''
' This form is used to show the world map.
' It has two levels. The world map and the dungeons map.
' You can toggle between them pressing the arrows
'
' @file     frmMapa.frm
' @author Marco Vanotti (MarKoxX) marcovanotti15@gmail.com
' @version 1.0.0
' @date 20080724

''
' Checks what Key is down. If the key is const vbKeyDown or const vbKeyUp, it toggles the maps, else the form unloads.
'
' @param KeyCode Specifies the key pressed
' @param Shift Specifies if Shift Button is pressed
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'*************************************************
'Author: Marco Vanotti (MarKoxX)
'Last modified: 24/07/08
'
'*************************************************

    Select Case KeyCode
        Case vbKeyDown, vbKeyUp 'Cambiamos el "nivel" del mapa, al estilo Zelda ;D
            ToggleImgMaps
        Case Else
            Unload Me
    End Select
    
End Sub

''
' Toggle which image is visible.
'
Private Sub ToggleImgMaps()
'*************************************************
'Author: Marco Vanotti (MarKoxX)
'Last modified: 24/07/08
'
'*************************************************

    imgMap.Visible = Not imgMap.Visible
    imgMapDungeon.Visible = Not imgMapDungeon.Visible
End Sub

''
' Load the images. Resizes the form, adjusts image's left and top and set lblTexto's Top and Left.
'
Private Sub Form_Load()
'*************************************************
'Author: Marco Vanotti (MarKoxX)
'Last modified: 24/07/08
'
'*************************************************

On Error GoTo error
    
    'Cargamos las imagenes de los mapas
    imgMap.Picture = LoadPicture(DirGraficos & "mapa1")
    imgMapDungeon.Picture = LoadPicture(DirGraficos & "mapa2")
    
    
    'Ajustamos el tamaño del formulario a la imagen más grande
    If imgMap.width > imgMapDungeon.width Then
        Me.width = imgMap.width
    Else
        Me.width = imgMapDungeon.width
    End If
    
    If imgMap.height > imgMapDungeon.height Then
        Me.height = imgMap.height + lblTexto.height
    Else
        Me.height = imgMapDungeon.height + lblTexto.height
    End If
    
    'Movemos ambas imágenes al centro del formulario
    imgMap.Left = Me.width * 0.5 - imgMap.width * 0.5
    imgMap.Top = (Me.height - lblTexto.height) * 0.5 - imgMap.height * 0.5
    
    imgMapDungeon.Left = Me.width * 0.5 - imgMapDungeon.width * 0.5
    imgMapDungeon.Top = (Me.height - lblTexto.height) * 0.5 - imgMapDungeon.height * 0.5
    
    lblTexto.Top = Me.height - lblTexto.height
    lblTexto.Left = Me.width * 0.5 - lblTexto.width * 0.5
    
    imgMapDungeon.Visible = False
    Exit Sub
error:
    MsgBox Err.Description, vbInformation, "Error: " & Err.number
    Unload Me
End Sub

