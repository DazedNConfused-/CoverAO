VERSION 5.00
Begin VB.Form frmCargando 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9000
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   11970
   ControlBox      =   0   'False
   Icon            =   "frmCargando.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCargando.frx":0CCA
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   798
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox MP3Files 
      Height          =   480
      Left            =   0
      Pattern         =   "*.mp3"
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image imgProgress 
      Height          =   645
      Left            =   2205
      Picture         =   "frmCargando.frx":60A5D
      Top             =   8040
      Width           =   7575
   End
End
Attribute VB_Name = "frmCargando"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.9.0.9
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
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
Dim i As Integer
Private porcentajeActual As Integer
Private Const PROGRESS_DELAY = 10
Private Const PROGRESS_DELAY_BACKWARDS = 4
Private Const DEFAULT_PROGRESS_WIDTH = 336
Private Const DEFAULT_STEP_FORWARD = 1
Private Const DEFAULT_STEP_BACKWARDS = -3

'Extraido de WinterAO y adaptado para DX8
Private Sub Form_Load()
'Me.Picture = LoadPicture(App.Path & "\Recursos\Graficos\cargando.jpg")
If Not Transparencia(Me.hwnd, 0) = 0 Then
   
    MsgBox " Esta función Api no es soportada en Versiones" _
           & "anteriores a windows 2000", vbCritical
    Me.Show
Else
 
    ' Gradua la transparencia del formulario hasta hacerla visible _
     es decir desde el valor 0 hasta el 255
   
    'desactiva el Formulario
    Me.Enabled = False
    Me.Show
   
    For i = 0 To 255 Step 2
        ' Va aplicando los distintos valores y grados de transparencia al form
        Call Transparencia(Me.hwnd, i)
        DoEvents
    Next
   
    'reactiva la ventana
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

Public Sub progresoConDelay(ByVal porcentaje As Integer)
If porcentaje = porcentajeActual Then Exit Sub
Dim step As Integer, stepInterval As Integer, timer As Long, tickCount As Long
If (porcentaje > porcentajeActual) Then
    step = DEFAULT_STEP_FORWARD
    stepInterval = PROGRESS_DELAY
Else
    step = DEFAULT_STEP_BACKWARDS
    stepInterval = PROGRESS_DELAY_BACKWARDS
End If
Do Until CompararPorcentaje(porcentaje, porcentajeActual, step)
    Do Until (timer + stepInterval) <= GetTickCount()
        DoEvents
    Loop
    timer = GetTickCount()
    porcentajeActual = porcentajeActual + step
    Call EstablecerProgreso(porcentajeActual)
Loop
End Sub
 
 
Public Sub EstablecerProgreso(ByVal nuevoPorcentaje As Integer)
If nuevoPorcentaje >= 0 And nuevoPorcentaje <= 100 Then
    imgProgress.width = DEFAULT_PROGRESS_WIDTH * CLng(nuevoPorcentaje) / 100
ElseIf nuevoPorcentaje > 100 Then
    imgProgress.width = DEFAULT_PROGRESS_WIDTH
Else
    imgProgress.width = 0
End If
porcentajeActual = nuevoPorcentaje
End Sub
 
Private Function CompararPorcentaje(ByVal porcentajeTarget As Integer, ByVal porcentajeAct As Integer, ByVal step As Integer) As Boolean
 
If step = DEFAULT_STEP_FORWARD Then
    CompararPorcentaje = (porcentajeAct >= porcentajeTarget)
Else
    CompararPorcentaje = (porcentajeAct <= porcentajeTarget)
End If
End Function
 

