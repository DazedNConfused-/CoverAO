Attribute VB_Name = "modTorneo"
Private Const MapaTorneo As Byte = 2 ' Mapa para teletransportar

Public TorneoAct As Boolean 'Esta activado modo torneo?

Private LvlMin As Byte 'Nivel minimo para participar
Private LvlMax As Byte 'Nivel maximo  ||      ||
Private Raza As Byte ' Raza requerida para participar ( 1 Enanos , 2 Altos, 0 todas)
Private Clase As Byte ' Clase   ||      ||      || ( 0 Todas)
Private Faccion As Byte ' Faccion ||     ||      || ( 0 todas , 1 Caos,2 Armada , 3 Ciudas , 4 Criminales)

Public Sub Torneo_Crear(ByVal NivelMin As Byte, ByVal NivelMax As Byte, ByVal eRaza As Byte, ByVal eClase As Byte, ByVal eCriminal As Byte)
    
    Dim Text As String
    
    LvlMin = NivelMin
    LvlMax = NivelMax
    
    Raza = eRaza
    Clase = eClase
    Faccion = efaccion
    
    Text = "Se ha decidido organizar un torneo automatico. Para participar manden '/TORNEO'. Tengan que cuenta que necesita contar con los requisitos:" _
         & " Nivel Minimo: " & LvlMin & " ;; Nivel Maximo: " & LvlMax & " ;; Raza: "
    
    If Raza = 0 Then
        Text = Text & " Todas"
    ElseIf Raza = 1 Then
        Text = Text & " Petisos"
    ElseIf Raza = 2 Then
        Text = Text & " Altos"
    End If
    
    Text = Text & " ;; Clase : "
    
End Sub
