VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMeteorologic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************************
'clsMeteorologic - ImperiumAO - v1.3.0
'
'Meteorologic engine.
'
'*****************************************************************
'Respective portions copyrighted by contributors listed below.
'
'This library is free software; you can redistribute it and/or
'modify it under the terms of the GNU Lesser General Public
'License as published by the Free Software Foundation version 2.1 of
'the License
'
'This library is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
'Lesser General Public License for more details.
'
'You should have received a copy of the GNU Lesser General Public
'License along with this library; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'*****************************************************************

'*****************************************************************
'Sinuhe (sinuhe@imperiumao.com.ar)
'Augusto José Rando (barrin@imperiumao.com.ar)
'   - First Relase
'*****************************************************************

Option Explicit

Private Const DURACION_TRANS = 300

'Main Status
Private Const DIA As Byte = 1
Private Const ATARDECER As Byte = 2
Private Const NOCHE As Byte = 3
Private Const AMANECER As Byte = 4

Private Const FIN_ANIMACION As Byte = &HFF

'Secondary Status
Private Const NORMAL As Byte = 0
Private Const NUBLADO As Byte = 1
Private Const LLUVIA As Byte = 2
Private Const NIEVE As Byte = 3

Private Const Particula_Lluvia As Long = 8
Private Const Particula_Nieve As Long = 56

Private Type SYSTEMTIME
  wYear As Integer
  wMonth As Integer
  wDayOfWeek As Integer
  wDay As Integer
  wHour As Integer
  wMinute As Integer
  wSecond As Integer
  wMilliseconds As Integer
End Type

Private Type MeteoStat
    EstadoActual As Byte
    EstadoProximo As Byte
End Type

Private Type Efectos
    MainStatus As MeteoStat
    SecondaryStatus As MeteoStat
End Type

Private Type D3DCOLORVALUE
    a As Long
    r As Long
    g As Long
    b As Long
End Type

Private Meteo As Efectos

Private m_Color As D3DCOLORVALUE
Private Out_Color As Long
Private TimeElapsed As Single

Private m_Color_Dia As D3DCOLORVALUE
Private m_Color_Noche As D3DCOLORVALUE
Private m_Color_Tarde As D3DCOLORVALUE
Private m_Color_Manana As D3DCOLORVALUE

Private m_Hora_Actual As String
Private m_Last_Hora_Actual As String

Private Declare Sub GetSystemTime Lib "kernel32.dll" (lpSystemTime As SYSTEMTIME)
Private theTime As SYSTEMTIME

'Inicializa colores de los diferentes estados
Private Sub Class_Initialize()

'No animo nada al iniciar
Meteo.MainStatus.EstadoProximo = FIN_ANIMACION
Meteo.MainStatus.EstadoActual = DIA

With m_Color_Dia
    .a = 255
    .b = 255
    .r = 255
    .g = 255
End With

With m_Color_Noche
    .a = 255
    .b = 170
    .r = 170
    .g = 170
End With

With m_Color_Tarde
    .a = 255
    .b = 200
    .r = 230
    .g = 200
End With

With m_Color_Manana
    .a = 255
    .b = 230
    .r = 200
    .g = 200
End With

End Sub

Public Sub Meteo_Logic()

Dim Diferencia As Integer
Dim ParticleIndex As Long
Dim CambioValor As Single
Dim C1 As D3DCOLORVALUE
Dim C2 As D3DCOLORVALUE

If Meteo.MainStatus.EstadoProximo <> FIN_ANIMACION Then
    TimeElapsed = TimeElapsed + General_Get_Elapsed_Time
    CambioValor = TimeElapsed / DURACION_TRANS
    If CambioValor > 1 Then
        Meteo.MainStatus.EstadoActual = Meteo.MainStatus.EstadoProximo
        Meteo.MainStatus.EstadoProximo = FIN_ANIMACION
    End If
End If

Select Case Meteo.MainStatus.EstadoActual
    Case DIA
        If Meteo.MainStatus.EstadoProximo = ATARDECER Then
            C1 = m_Color_Dia
            C2 = m_Color_Tarde
            Diferencia = C1.r - C2.r
            If Diferencia > 255 Then Diferencia = 255
            m_Color.r = C1.r - Diferencia * CambioValor
            
            Diferencia = C1.g - C2.g
            If Diferencia > 255 Then Diferencia = 255
            m_Color.g = C1.g - Diferencia * CambioValor
            
            Diferencia = C1.b - C2.b
            If Diferencia > 255 Then Diferencia = 255
            m_Color.b = C1.b - Diferencia * CambioValor
        Else
            m_Color = m_Color_Dia
        End If
    Case NOCHE
        If Meteo.MainStatus.EstadoProximo = AMANECER Then
            C1 = m_Color_Noche
            C2 = m_Color_Manana
            Diferencia = C1.r - C2.r
            If Diferencia > 255 Then Diferencia = 255
            m_Color.r = C1.r - Diferencia * CambioValor
                
            Diferencia = C1.g - C2.g
            If Diferencia > 255 Then Diferencia = 255
            m_Color.g = C1.g - Diferencia * CambioValor
                
            Diferencia = C1.b - C2.b
            If Diferencia > 255 Then Diferencia = 255
            m_Color.b = C1.b - Diferencia * CambioValor
        Else
            m_Color = m_Color_Noche
        End If
    Case AMANECER
        If Meteo.MainStatus.EstadoProximo = DIA Then
            C1 = m_Color_Manana
            C2 = m_Color_Dia
            Diferencia = C1.r - C2.r
            If Diferencia > 255 Then Diferencia = 255
            m_Color.r = C1.r - Diferencia * CambioValor
                
            Diferencia = C1.g - C2.g
            If Diferencia > 255 Then Diferencia = 255
            m_Color.g = C1.g - Diferencia * CambioValor
                
            Diferencia = C1.b - C2.b
            If Diferencia > 255 Then Diferencia = 255
            m_Color.b = C1.b - Diferencia * CambioValor
        Else
            m_Color = m_Color_Manana
        End If
    Case ATARDECER
        If Meteo.MainStatus.EstadoProximo = NOCHE Then
            C1 = m_Color_Tarde
            C2 = m_Color_Noche
            Diferencia = C1.r - C2.r
            If Diferencia > 255 Then Diferencia = 255
            m_Color.r = C1.r - Diferencia * CambioValor
                
            Diferencia = C1.g - C2.g
            If Diferencia > 255 Then Diferencia = 255
            m_Color.g = C1.g - Diferencia * CambioValor
                
            Diferencia = C1.b - C2.b
            If Diferencia > 255 Then Diferencia = 255
            m_Color.b = C1.b - Diferencia * CambioValor
        Else
            m_Color = m_Color_Tarde
        End If
End Select

ParticleIndex = Engine.Engine_Meteo_Particle_Get

Select Case Meteo.SecondaryStatus.EstadoActual
    Case LLUVIA
        If (ParticleIndex = 0) And (UserMapExt = 1) Then
            Call Engine.Engine_Meteo_Particle_Set(Particula_Lluvia)
        Else
            If (UserMapExt <> 1) And (ParticleIndex <> 0) Then Call Engine.Engine_Meteo_Particle_Set(-1)
        End If
    Case NIEVE
        If (ParticleIndex = 0) And (UserMapExt = 2) Then
            Call Engine.Engine_Meteo_Particle_Set(Particula_Nieve)
        Else
            If (UserMapExt <> 2) And (ParticleIndex <> 0) Then Call Engine.Engine_Meteo_Particle_Set(-1)
        End If
    Case NORMAL
        If (ParticleIndex <> 0) Then Call Engine.Engine_Meteo_Particle_Set(-1)
End Select

Out_Color = General_RGB_Color_to_Long(m_Color.r, m_Color.g, m_Color.b, m_Color.a)
Call Engine.Map_Base_Light_Set(Out_Color)

End Sub

Public Sub Time_Logic()

GetSystemTime theTime
m_Hora_Actual = format((theTime.wHour - 3), "00")
If val(m_Hora_Actual) = -3 Then m_Hora_Actual = "21"
If val(m_Hora_Actual) = -2 Then m_Hora_Actual = "22"
If val(m_Hora_Actual) = -1 Then m_Hora_Actual = "23"

If m_Hora_Actual <> m_Last_Hora_Actual Then
    'frmMain.imgHora.Picture = General_Load_Picture_From_Resource(m_Hora_Actual & "")
    m_Last_Hora_Actual = m_Hora_Actual
End If

End Sub

Public Function Get_Time_String() As String

Dim HoraActual As Integer

GetSystemTime theTime

HoraActual = (theTime.wHour - 3)
If HoraActual = -3 Then HoraActual = 21
If HoraActual = -2 Then HoraActual = 22
If HoraActual = -1 Then HoraActual = 23

Get_Time_String = HoraActual & ":" & format(theTime.wMinute, "00") & "... "

Select Case HoraActual
    Case 5, 6, 7
        Get_Time_String = Get_Time_String & "el sol se asoma lentamente en el horizonte"
    Case 8, 9, 10, 11, 12, 13, 14, 15, 16, 17
        Get_Time_String = Get_Time_String & "¡no pierdas el tiempo!"
    Case 18, 19
        Get_Time_String = Get_Time_String & "lentamente el dia termina"
    Case Else
        Get_Time_String = Get_Time_String & "¿despierto a estas horas? ¡no olvides visitar El Mesón Hostigado!"
End Select

End Function

Public Sub SetNuevoEstado(Estado As Byte)

TimeElapsed = 0

If Meteo.MainStatus.EstadoProximo <> FIN_ANIMACION Then
    'como no termino la animación anterior, al menos le pongo que la que estaba haciendo sea desde donde sale
    Meteo.MainStatus.EstadoActual = Meteo.MainStatus.EstadoProximo
End If

Meteo.MainStatus.EstadoProximo = Estado

End Sub

Public Property Get SecondaryStatus() As Byte
SecondaryStatus = Meteo.SecondaryStatus.EstadoActual
End Property

Public Property Let SecondaryStatus(ByVal vNewValue As Byte)
Meteo.SecondaryStatus.EstadoActual = vNewValue
End Property

Public Property Get PrimaryStatus() As Byte
PrimaryStatus = Meteo.MainStatus.EstadoActual
End Property

Public Property Let PrimaryStatus(ByVal vNewValue As Byte)
Meteo.MainStatus.EstadoProximo = vNewValue
End Property

Public Sub ForzarEstado(Estado As Byte)
TimeElapsed = 0
Meteo.MainStatus.EstadoActual = Estado
Meteo.MainStatus.EstadoProximo = FIN_ANIMACION
End Sub

Public Sub NextEstado()

TimeElapsed = 0

If Meteo.MainStatus.EstadoProximo <> FIN_ANIMACION Then
    'Como no termino la animación anterior, al menos le pongo que la que estaba haciendo sea desde donde sale
    Meteo.MainStatus.EstadoActual = Meteo.MainStatus.EstadoProximo
End If

Meteo.MainStatus.EstadoProximo = Meteo.MainStatus.EstadoActual + 1
If Meteo.MainStatus.EstadoProximo > AMANECER Then Meteo.MainStatus.EstadoProximo = DIA

End Sub

