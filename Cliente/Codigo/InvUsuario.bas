Attribute VB_Name = "InvUsuario"
'Argentum Online 0.12.2
'Copyright (C) 2002 M�rquez Pablo Ignacio
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
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez

Option Explicit

Public Function TieneObjetosRobables(ByVal UserIndex As Integer) As Boolean

'17/09/02
'Agregue que la funci�n se asegure que el objeto no es un barco

On Error Resume Next

Dim i As Integer
Dim ObjIndex As Integer

For i = 1 To MAX_INVENTORY_SLOTS
    ObjIndex = UserList(UserIndex).Invent.Object(i).ObjIndex
    If ObjIndex > 0 Then
            If (ObjData(ObjIndex).OBJType <> eOBJType.otLlaves And _
                ObjData(ObjIndex).OBJType <> eOBJType.otMonturas And _
                ObjData(ObjIndex).OBJType <> eOBJType.otBarcos) Then
                  TieneObjetosRobables = True
                  Exit Function
            End If
    
    End If
Next i


End Function

Function ClasePuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean
On Error GoTo manejador

'Call LogTarea("ClasePuedeUsarItem")

Dim flag As Boolean
If ObjIndex = 0 Then Exit Function

'Admins can use ANYTHING!
'If UserList(UserIndex).flags.Privilegios And PlayerType.User Then
If ObjData(ObjIndex).ClaseTipo = 0 Then
    If ObjData(ObjIndex).ClaseProhibida(1) > 0 Then
        Dim i As Integer
        For i = 1 To NUMCLASES
            If ObjData(ObjIndex).ClaseProhibida(i) = UserList(UserIndex).Clase Then
                ClasePuedeUsarItem = False
                Exit Function
            End If
        Next i
    End If
Else
    If (UserList(UserIndex).Clase = Gladiador Or _
    UserList(UserIndex).Clase = Guerrero Or _
    UserList(UserIndex).Clase = Paladin Or _
    UserList(UserIndex).Clase = Cazador Or _
    UserList(UserIndex).Clase = Mercenario) Then
        If ObjData(ObjIndex).ClaseTipo = 1 Then
            ClasePuedeUsarItem = True
        Else
            ClasePuedeUsarItem = False
        End If
    End If
End If

ClasePuedeUsarItem = True

Exit Function

manejador:
    LogError ("Error en ClasePuedeUsarItem")
End Function

Sub QuitarNewbieObj(ByVal UserIndex As Integer)
Dim j As Integer
For j = 1 To MAX_INVENTORY_SLOTS
        If UserList(UserIndex).Invent.Object(j).ObjIndex > 0 Then
             
             If ObjData(UserList(UserIndex).Invent.Object(j).ObjIndex).Newbie = 1 Then _
                    Call QuitarUserInvItem(UserIndex, j, MAX_INVENTORY_OBJS)
                    Call UpdateUserInv(False, UserIndex, j)
        
        End If
Next j

'[Barrin 17-12-03] Si el usuario dej� de ser Newbie, y estaba en el Newbie Dungeon
'es transportado a su hogar de origen ;)
If UCase$(MapInfo(UserList(UserIndex).Pos.map).Restringir) = "NEWBIE" Then
    
    Dim DeDonde As WorldPos
            DeDonde = Nix
    
    Call WarpUserChar(UserIndex, DeDonde.map, DeDonde.X, DeDonde.Y, True)

End If
'[/Barrin]

End Sub

Sub LimpiarInventario(ByVal UserIndex As Integer)


Dim j As Integer
For j = 1 To MAX_INVENTORY_SLOTS
        UserList(UserIndex).Invent.Object(j).ObjIndex = 0
        UserList(UserIndex).Invent.Object(j).amount = 0
        UserList(UserIndex).Invent.Object(j).Equipped = 0
        
Next

UserList(UserIndex).Invent.NroItems = 0

UserList(UserIndex).Invent.ArmourEqpObjIndex = 0
UserList(UserIndex).Invent.ArmourEqpSlot = 0

UserList(UserIndex).Invent.WeaponEqpObjIndex = 0
UserList(UserIndex).Invent.WeaponEqpSlot = 0

UserList(UserIndex).Invent.CascoEqpObjIndex = 0
UserList(UserIndex).Invent.CascoEqpSlot = 0

UserList(UserIndex).Invent.EscudoEqpObjIndex = 0
UserList(UserIndex).Invent.EscudoEqpSlot = 0

UserList(UserIndex).Invent.AnilloEqpObjIndex = 0
UserList(UserIndex).Invent.AnilloEqpSlot = 0

UserList(UserIndex).Invent.MunicionEqpObjIndex = 0
UserList(UserIndex).Invent.MunicionEqpSlot = 0

UserList(UserIndex).Invent.BarcoObjIndex = 0
UserList(UserIndex).Invent.BarcoSlot = 0

UserList(UserIndex).Invent.MonturaObjIndex = 0
UserList(UserIndex).Invent.MonturaSlot = 0
End Sub

Sub TirarOro(ByVal Cantidad As Long, ByVal UserIndex As Integer)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 23/01/2007
'23/01/2007 -> Pablo (ToxicWaste): Billetera invertida y explotar oro en el agua.
'***************************************************
On Error GoTo Errhandler

'If Cantidad > 100000 Then Exit Sub

'SI EL Pjta TIENE ORO LO TIRAMOS
If (Cantidad > 0) And (Cantidad <= UserList(UserIndex).Stats.GLD) Then
        Dim i As Byte
        Dim MiObj As Obj
        'info debug
        Dim loops As Integer
        
        'Seguridad Alkon (guardo el oro tirado si supera los 50k)
        If Cantidad > 50000 Then
            Dim j As Integer
            Dim k As Integer
            Dim M As Integer
            Dim Cercanos As String
            M = UserList(UserIndex).Pos.map
            For j = UserList(UserIndex).Pos.X - 10 To UserList(UserIndex).Pos.X + 10
                For k = UserList(UserIndex).Pos.Y - 10 To UserList(UserIndex).Pos.Y + 10
                    If InMapBounds(M, j, k) Then
                        If MapData(M, j, k).UserIndex > 0 Then
                            Cercanos = Cercanos & UserList(MapData(M, j, k).UserIndex).name & ","
                        End If
                    End If
                Next k
            Next j
            Call LogDesarrollo(UserList(UserIndex).name & " tira oro. Cercanos: " & Cercanos)
        End If
        '/Seguridad
        Dim Extra As Long
        Dim TeniaOro As Long
        TeniaOro = UserList(UserIndex).Stats.GLD
        If Cantidad > 500000 Then 'Para evitar explotar demasiado
            Extra = Cantidad - 500000
            Cantidad = 500000
        End If
        
        Do While (Cantidad > 0)
            
            If Cantidad > MAX_INVENTORY_OBJS And UserList(UserIndex).Stats.GLD > MAX_INVENTORY_OBJS Then
                MiObj.amount = MAX_INVENTORY_OBJS
                Cantidad = Cantidad - MiObj.amount
            Else
                MiObj.amount = Cantidad
                Cantidad = Cantidad - MiObj.amount
            End If

            MiObj.ObjIndex = iORO
            
            If EsGM(UserIndex) Then Call LogGM(UserList(UserIndex).name, "Tiro cantidad:" & MiObj.amount & " Objeto:" & ObjData(MiObj.ObjIndex).name)
            Dim AuxPos As WorldPos
            
            If UserList(UserIndex).Clase = eClass.Mercenario And UserList(UserIndex).Invent.BarcoObjIndex = 476 Then
                AuxPos = TirarItemAlPiso(UserList(UserIndex).Pos, MiObj, False)
                If AuxPos.X <> 0 And AuxPos.Y <> 0 Then
                    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - MiObj.amount
                End If
            Else
                AuxPos = TirarItemAlPiso(UserList(UserIndex).Pos, MiObj, True)
                If AuxPos.X <> 0 And AuxPos.Y <> 0 Then
                    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - MiObj.amount
                End If
            End If
            
            'info debug
            loops = loops + 1
            If loops > 100 Then
                LogError ("Error en tiraroro")
                Exit Sub
            End If
            
        Loop
        If TeniaOro = UserList(UserIndex).Stats.GLD Then Extra = 0
        If Extra > 0 Then
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - Extra
        End If
    
End If

Exit Sub

Errhandler:

End Sub

Sub QuitarUserInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Cantidad As Integer)
    If Slot < 1 Or Slot > MAX_INVENTORY_SLOTS Then Exit Sub
    
    With UserList(UserIndex).Invent.Object(Slot)
        If .amount <= Cantidad Then
            If .Equipped = 1 Then
                Call Desequipar(UserIndex, Slot)
            ElseIf UserList(UserIndex).flags.Montando = 1 Then
                If Slot = UserList(UserIndex).Invent.MonturaSlot Then
                    DoEquita UserIndex, .ObjIndex, Slot
                End If
            End If
        End If
        
        'Quita un objeto
        .amount = .amount - Cantidad
        '�Quedan mas?
        If .amount <= 0 Then
            UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems - 1
            .ObjIndex = 0
            .amount = 0
        End If
    End With
End Sub

Sub UpdateUserInv(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal Slot As Byte)

Dim NullObj As UserObj
Dim LoopC As Long

'Actualiza un solo slot
If Not UpdateAll Then

    'Actualiza el inventario
    If UserList(UserIndex).Invent.Object(Slot).ObjIndex > 0 Then
        Call ChangeUserInv(UserIndex, Slot, UserList(UserIndex).Invent.Object(Slot))
    Else
        Call ChangeUserInv(UserIndex, Slot, NullObj)
    End If

Else

'Actualiza todos los slots
    For LoopC = 1 To MAX_INVENTORY_SLOTS
        'Actualiza el inventario
        If UserList(UserIndex).Invent.Object(LoopC).ObjIndex > 0 Then
            Call ChangeUserInv(UserIndex, LoopC, UserList(UserIndex).Invent.Object(LoopC))
        Else
            Call ChangeUserInv(UserIndex, LoopC, NullObj)
        End If
    Next LoopC
End If

End Sub

Sub DropObj(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal num As Integer, ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer)

Dim Obj As Obj

If num > 0 Then
  
  If num > UserList(UserIndex).Invent.Object(Slot).amount Then num = UserList(UserIndex).Invent.Object(Slot).amount
  
  'Check objeto en el suelo
  If MapData(UserList(UserIndex).Pos.map, X, Y).ObjInfo.ObjIndex = 0 Or MapData(UserList(UserIndex).Pos.map, X, Y).ObjInfo.ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex Then
        Obj.ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
        
        If num + MapData(UserList(UserIndex).Pos.map, X, Y).ObjInfo.amount > MAX_INVENTORY_OBJS Then
            num = MAX_INVENTORY_OBJS - MapData(UserList(UserIndex).Pos.map, X, Y).ObjInfo.amount
        End If
        
        Obj.amount = num
        
        Call MakeObj(Obj, map, X, Y)
        Call QuitarUserInvItem(UserIndex, Slot, num)
        Call UpdateUserInv(False, UserIndex, Slot)
        
        If ObjData(Obj.ObjIndex).OBJType = eOBJType.otBarcos Then
            Call WriteConsoleMsg(1, UserIndex, "", FontTypeNames.FONTTYPE_TALK)
        End If
        
        If Not UserList(UserIndex).flags.Privilegios And PlayerType.User Then Call LogGM(UserList(UserIndex).name, "Tiro cantidad:" & num & " Objeto:" & ObjData(Obj.ObjIndex).name)
        
        'Log de Objetos que se tiran al piso. Pablo (ToxicWaste) 07/09/07
        'Es un Objeto que tenemos que loguear?
        If ObjData(Obj.ObjIndex).Log = 1 Then
            Call LogDesarrollo(UserList(UserIndex).name & " tir� al piso " & Obj.amount & " " & ObjData(Obj.ObjIndex).name & " Mapa: " & map & " X: " & X & " Y: " & Y)
        ElseIf Obj.amount > 5000 Then 'Es mucha cantidad? > Sub� a 5000 el minimo porque si no se llenaba el log de cosas al pedo. (NicoNZ)
        'Si no es de los prohibidos de loguear, lo logueamos.
            If ObjData(Obj.ObjIndex).NoLog <> 1 Then
                Call LogDesarrollo(UserList(UserIndex).name & " tir� al piso " & Obj.amount & " " & ObjData(Obj.ObjIndex).name & " Mapa: " & map & " X: " & X & " Y: " & Y)
            End If
        End If
  Else
    Call WriteConsoleMsg(1, UserIndex, "No hay espacio en el piso.", FontTypeNames.FONTTYPE_INFO)
  End If
    
End If

End Sub

Sub EraseObj(ByVal num As Integer, ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer)

MapData(map, X, Y).ObjInfo.amount = MapData(map, X, Y).ObjInfo.amount - num

If MapData(map, X, Y).ObjInfo.amount <= 0 Then
    MapData(map, X, Y).ObjInfo.ObjIndex = 0
    MapData(map, X, Y).ObjInfo.amount = 0
    
    Call modSendData.SendToAreaByPos(map, X, Y, PrepareMessageObjectDelete(X, Y))
End If

End Sub

Sub MakeObj(ByRef Obj As Obj, ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer)

If Obj.ObjIndex > 0 And Obj.ObjIndex <= UBound(ObjData) Then

    If MapData(map, X, Y).ObjInfo.ObjIndex = Obj.ObjIndex Then
        MapData(map, X, Y).ObjInfo.amount = MapData(map, X, Y).ObjInfo.amount + Obj.amount
    Else
        MapData(map, X, Y).ObjInfo = Obj
        
        Call modSendData.SendToAreaByPos(map, X, Y, PrepareMessageObjectCreate(ObjData(Obj.ObjIndex).GrhIndex, X, Y, Obj.ObjIndex))
    End If
End If

End Sub

Function MeterItemEnInventario(ByVal UserIndex As Integer, ByRef MiObj As Obj) As Boolean
On Error GoTo Errhandler

'Call LogTarea("MeterItemEnInventario")
 
Dim X As Integer
Dim Y As Integer
Dim Slot As Byte

'�el user ya tiene un objeto del mismo tipo?
Slot = 1
Do Until UserList(UserIndex).Invent.Object(Slot).ObjIndex = MiObj.ObjIndex And _
         UserList(UserIndex).Invent.Object(Slot).amount + MiObj.amount <= MAX_INVENTORY_OBJS
   Slot = Slot + 1
   If Slot > MAX_INVENTORY_SLOTS Then
         Exit Do
   End If
Loop
    
'Sino busca un slot vacio
If Slot > MAX_INVENTORY_SLOTS Then
   Slot = 1
   Do Until UserList(UserIndex).Invent.Object(Slot).ObjIndex = 0
       Slot = Slot + 1
       If Slot > MAX_INVENTORY_SLOTS Then
           Call WriteConsoleMsg(1, UserIndex, "No podes cargar mas objetos.", FontTypeNames.FONTTYPE_FIGHT)
           MeterItemEnInventario = False
           Exit Function
       End If
   Loop
   UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems + 1
End If
    
'Mete el objeto
If UserList(UserIndex).Invent.Object(Slot).amount + MiObj.amount <= MAX_INVENTORY_OBJS Then
   'Menor que MAX_INV_OBJS
   UserList(UserIndex).Invent.Object(Slot).ObjIndex = MiObj.ObjIndex
   UserList(UserIndex).Invent.Object(Slot).amount = UserList(UserIndex).Invent.Object(Slot).amount + MiObj.amount
Else
   UserList(UserIndex).Invent.Object(Slot).amount = MAX_INVENTORY_OBJS
End If
    
MeterItemEnInventario = True
       
Call UpdateUserInv(False, UserIndex, Slot)


Exit Function
Errhandler:

End Function


Sub GetObj(ByVal UserIndex As Integer)

Dim Obj As ObjData
Dim MiObj As Obj
Dim ObjPos As String

'�Hay algun obj?
If MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).ObjInfo.ObjIndex > 0 Then
    '�Esta permitido agarrar este obj?
    If ObjData(MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).ObjInfo.ObjIndex).Agarrable <> 1 Then
        Dim X As Integer
        Dim Y As Integer
        Dim Slot As Byte
        
        X = UserList(UserIndex).Pos.X
        Y = UserList(UserIndex).Pos.Y
        Obj = ObjData(MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).ObjInfo.ObjIndex)
        MiObj.amount = MapData(UserList(UserIndex).Pos.map, X, Y).ObjInfo.amount
        MiObj.ObjIndex = MapData(UserList(UserIndex).Pos.map, X, Y).ObjInfo.ObjIndex
        
               If MiObj.ObjIndex = 12 Then
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + MiObj.amount
        Call EraseObj(MapData(UserList(UserIndex).Pos.map, X, Y).ObjInfo.amount, UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
        Call WriteUpdateGold(UserIndex)
        Exit Sub
        End If
        
        If MeterItemEnInventario(UserIndex, MiObj) Then
            'Quitamos el objeto
            Call EraseObj(MapData(UserList(UserIndex).Pos.map, X, Y).ObjInfo.amount, UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
            If Not UserList(UserIndex).flags.Privilegios And PlayerType.User Then Call LogGM(UserList(UserIndex).name, "Agarro:" & MiObj.amount & " Objeto:" & ObjData(MiObj.ObjIndex).name)
            
            'Log de Objetos que se agarran del piso. Pablo (ToxicWaste) 07/09/07
            'Es un Objeto que tenemos que loguear?
            If ObjData(MiObj.ObjIndex).Log = 1 Then
                ObjPos = " Mapa: " & UserList(UserIndex).Pos.map & " X: " & UserList(UserIndex).Pos.X & " Y: " & UserList(UserIndex).Pos.Y
                Call LogDesarrollo(UserList(UserIndex).name & " junt� del piso " & MiObj.amount & " " & ObjData(MiObj.ObjIndex).name & ObjPos)
            ElseIf MiObj.amount > MAX_INVENTORY_OBJS - 1000 Then 'Es mucha cantidad?
                'Si no es de los prohibidos de loguear, lo logueamos.
                If ObjData(MiObj.ObjIndex).NoLog <> 1 Then
                    ObjPos = " Mapa: " & UserList(UserIndex).Pos.map & " X: " & UserList(UserIndex).Pos.X & " Y: " & UserList(UserIndex).Pos.Y
                    Call LogDesarrollo(UserList(UserIndex).name & " junt� del piso " & MiObj.amount & " " & ObjData(MiObj.ObjIndex).name & ObjPos)
                End If
            End If
        End If
        
    End If
Else
    Call WriteConsoleMsg(1, UserIndex, "No hay nada aqu�.", FontTypeNames.FONTTYPE_INFO)
End If

End Sub

Sub Desequipar(ByVal UserIndex As Integer, ByVal Slot As Byte)
'Desequipa el item slot del inventario
Dim Obj As ObjData


If (Slot < LBound(UserList(UserIndex).Invent.Object)) Or (Slot > UBound(UserList(UserIndex).Invent.Object)) Then
    Exit Sub
ElseIf UserList(UserIndex).Invent.Object(Slot).ObjIndex = 0 Then
    Exit Sub
End If

Obj = ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex)

Select Case Obj.OBJType
    Case eOBJType.otWeapon
        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        UserList(UserIndex).Invent.WeaponEqpObjIndex = 0
        UserList(UserIndex).Invent.WeaponEqpSlot = 0
        UserList(UserIndex).Char.WeaponAnim = NingunArma
        Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                
    Case eOBJType.otNudillos
        UserList(UserIndex).Invent.NudiEqpIndex = 0
        UserList(UserIndex).Invent.NudiEqpSlot = 0
        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        UserList(UserIndex).Char.WeaponAnim = NingunArma
        Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                
    Case eOBJType.otFlechas
        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        UserList(UserIndex).Invent.MunicionEqpObjIndex = 0
        UserList(UserIndex).Invent.MunicionEqpSlot = 0
        
    Case eOBJType.otAnillo
        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        UserList(UserIndex).Invent.AnilloEqpObjIndex = 0
        UserList(UserIndex).Invent.AnilloEqpSlot = 0
        
    Case eOBJType.otArmadura ' Puede ser un escudo, casco , o vestimenta
        UserList(UserIndex).Invent.Object(Slot).Equipped = 0
        
        Select Case ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).SubTipo
            Case 0
                UserList(UserIndex).Invent.ArmourEqpObjIndex = 0
                UserList(UserIndex).Invent.ArmourEqpSlot = 0
                Call DarCuerpoDesnudo(UserIndex)
                If Not UserList(UserIndex).flags.Montando = 1 Then Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                
            Case 1
                UserList(UserIndex).Invent.Object(Slot).Equipped = 0
                UserList(UserIndex).Invent.CascoEqpObjIndex = 0
                UserList(UserIndex).Invent.CascoEqpSlot = 0
                
                UserList(UserIndex).Char.CascoAnim = NingunCasco
                Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)

                
            Case 2
                UserList(UserIndex).Invent.Object(Slot).Equipped = 0
                UserList(UserIndex).Invent.EscudoEqpObjIndex = 0
                UserList(UserIndex).Invent.EscudoEqpSlot = 0
                
                UserList(UserIndex).Char.ShieldAnim = NingunEscudo
                Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                
        End Select
End Select

Call WriteUpdateUserStats(UserIndex)
Call UpdateUserInv(False, UserIndex, Slot)

End Sub

Function SexoPuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean
On Error GoTo Errhandler

If Not ObjIndex <> 0 Then Exit Function

If ObjData(ObjIndex).Mujer = 1 Then
    SexoPuedeUsarItem = UserList(UserIndex).genero <> eGenero.Hombre
ElseIf ObjData(ObjIndex).Hombre = 1 Then
    SexoPuedeUsarItem = UserList(UserIndex).genero <> eGenero.Mujer
Else
    SexoPuedeUsarItem = True
End If

Exit Function
Errhandler:
    Call LogError("SexoPuedeUsarItem")
End Function


Function FaccionPuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean

If Not ObjIndex <> 0 Then Exit Function

If ObjData(ObjIndex).Real = 1 Then
    FaccionPuedeUsarItem = esArmada(UserIndex)
ElseIf ObjData(ObjIndex).Caos = 1 Then
    FaccionPuedeUsarItem = esCaos(UserIndex)
ElseIf ObjData(ObjIndex).Milicia Then
    FaccionPuedeUsarItem = esMili(UserIndex)
Else
    FaccionPuedeUsarItem = True
End If

End Function

Sub EquiparInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte)
On Error GoTo Errhandler

'Equipa un item del inventario
Dim Obj As ObjData
Dim ObjIndex As Integer

ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
Obj = ObjData(ObjIndex)

If Not Obj.MinELV = 0 Then
    If Obj.MinELV > UserList(UserIndex).Stats.ELV Then
        Call WriteConsoleMsg(1, UserIndex, "Debes ser nivel " & Obj.MinELV & " para poder utilizar este objeto.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
End If

If Obj.Newbie = 1 And Not EsNewbie(UserIndex) Then
     Call WriteConsoleMsg(1, UserIndex, "Solo los newbies pueden usar este objeto.", FontTypeNames.FONTTYPE_INFO)
     Exit Sub
End If
        
Select Case Obj.OBJType

    Case eOBJType.otFlechas
       If ClasePuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) And _
          FaccionPuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) Then
                
                'Si esta equipado lo quita
                If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                    'Quitamos del inv el item
                    Call Desequipar(UserIndex, Slot)
                    Exit Sub
                End If
                
                'Quitamos el elemento anterior
                If UserList(UserIndex).Invent.MunicionEqpObjIndex > 0 Then
                    Call Desequipar(UserIndex, UserList(UserIndex).Invent.MunicionEqpSlot)
                End If
        
                UserList(UserIndex).Invent.Object(Slot).Equipped = 1
                UserList(UserIndex).Invent.MunicionEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
                UserList(UserIndex).Invent.MunicionEqpSlot = Slot
                
       Else
            Call WriteConsoleMsg(1, UserIndex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
       End If
    
    Case eOBJType.otArmadura
        If ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).SubTipo = 0 Then
            If UserList(UserIndex).flags.Navegando = 1 Then Exit Sub
            'Nos aseguramos que puede usarla
            If ClasePuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) And _
               SexoPuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) And _
               CheckRazaUsaRopa(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) And _
               FaccionPuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) Then
               
               'Si esta equipado lo quita
                If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                    Call Desequipar(UserIndex, Slot)
                    Call DarCuerpoDesnudo(UserIndex)
                    If Not UserList(UserIndex).flags.Montando = 1 Then Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                    Exit Sub
                End If
        
                'Quita el anterior
                If UserList(UserIndex).Invent.ArmourEqpObjIndex > 0 Then
                    Call Desequipar(UserIndex, UserList(UserIndex).Invent.ArmourEqpSlot)
                End If
        
                'Lo equipa
                UserList(UserIndex).Invent.Object(Slot).Equipped = 1
                UserList(UserIndex).Invent.ArmourEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
                UserList(UserIndex).Invent.ArmourEqpSlot = Slot
                
                UserList(UserIndex).Char.Body = Obj.Ropaje
                If Not UserList(UserIndex).flags.Montando = 1 Then Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)

                UserList(UserIndex).flags.Desnudo = 0
            Else
                Call WriteConsoleMsg(1, UserIndex, "Tu clase,genero o raza no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
            End If
        ElseIf ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).SubTipo = 1 Then
            If UserList(UserIndex).flags.Navegando = 1 Then Exit Sub
            If ClasePuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) Then
                'Si esta equipado lo quita
                If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                    Call Desequipar(UserIndex, Slot)
                    
                    UserList(UserIndex).Char.CascoAnim = NingunCasco
                    Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                    Exit Sub
                End If
        
                'Quita el anterior
                If UserList(UserIndex).Invent.CascoEqpObjIndex > 0 Then
                    Call Desequipar(UserIndex, UserList(UserIndex).Invent.CascoEqpSlot)
                End If
        
                'Lo equipa
                
                UserList(UserIndex).Invent.Object(Slot).Equipped = 1
                UserList(UserIndex).Invent.CascoEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
                UserList(UserIndex).Invent.CascoEqpSlot = Slot

                UserList(UserIndex).Char.CascoAnim = Obj.CascoAnim
                Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)

            Else
                Call WriteConsoleMsg(1, UserIndex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
            End If
        ElseIf ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex).SubTipo = 2 Then
            If UserList(UserIndex).flags.Navegando = 1 Then Exit Sub
                If ClasePuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) And FaccionPuedeUsarItem(UserIndex, UserList(UserIndex).Invent.Object(Slot).ObjIndex) Then
                    
                 'Si esta equipado lo quita
                 If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                     Call Desequipar(UserIndex, Slot)
                     
                     UserList(UserIndex).Char.ShieldAnim = NingunEscudo
                    Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)

                     Exit Sub
                 End If
         
                 'Quita el anterior
                 If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then
                     Call Desequipar(UserIndex, UserList(UserIndex).Invent.EscudoEqpSlot)
                 End If
         
                 'Lo equipa
                 
                 UserList(UserIndex).Invent.Object(Slot).Equipped = 1
                 UserList(UserIndex).Invent.EscudoEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
                 UserList(UserIndex).Invent.EscudoEqpSlot = Slot
                 
                 UserList(UserIndex).Char.ShieldAnim = Obj.ShieldAnim
                 Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)

             Else
                 Call WriteConsoleMsg(1, UserIndex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
             End If
        End If
    Case eOBJType.otWeapon
       If ClasePuedeUsarItem(UserIndex, ObjIndex) And _
          FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
            'Si esta equipado lo quita
            If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                'Quitamos del inv el item
                Call Desequipar(UserIndex, Slot)
                Exit Sub
            End If
            
            If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then
    If ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).MiniEscu <> 1 Then
    If ObjData(ObjIndex).DosManos = 1 Then
        Call WriteConsoleMsg(1, UserIndex, "Debes desequipar ese escudo para poder equipar esta espada.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
End If
End If
End If
            
            'Quitamos el elemento anterior
            If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                Call Desequipar(UserIndex, UserList(UserIndex).Invent.WeaponEqpSlot)
            End If
            
            If UserList(UserIndex).Invent.NudiEqpIndex > 0 Then
                Call Desequipar(UserIndex, UserList(UserIndex).Invent.NudiEqpSlot)
            End If
            
            UserList(UserIndex).Invent.Object(Slot).Equipped = 1
            UserList(UserIndex).Invent.WeaponEqpObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
            UserList(UserIndex).Invent.WeaponEqpSlot = Slot
            
            'Sonido
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_SACARARMA, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
            
            UserList(UserIndex).Char.WeaponAnim = Obj.WeaponAnim
            Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
       Else
            Call WriteConsoleMsg(1, UserIndex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
       End If
       
       Case eOBJType.otAnillo
       If ClasePuedeUsarItem(UserIndex, ObjIndex) And _
          FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
                'Si esta equipado lo quita
                If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                    'Quitamos del inv el item
                    Call Desequipar(UserIndex, Slot)
                    Exit Sub
                End If
               
                'Quitamos el elemento anterior
                If UserList(UserIndex).Invent.AnilloEqpObjIndex > 0 Then
                    Call Desequipar(UserIndex, UserList(UserIndex).Invent.AnilloEqpSlot)
                End If
       
                UserList(UserIndex).Invent.Object(Slot).Equipped = 1
                UserList(UserIndex).Invent.AnilloEqpObjIndex = ObjIndex
                UserList(UserIndex).Invent.AnilloEqpSlot = Slot
               
       End If
       

  Case eOBJType.otRuna
               
               If MapInfo(UserList(UserIndex).Pos.map).Pk = True Then
               Call WriteConsoleMsg(1, UserIndex, "Debes estar en Zona Segura.", FontTypeNames.FONTTYPE_INFO)
               Exit Sub
               End If
               
                Dim Ciudad As WorldPos
                Dim Destino As String
               
                Select Case (0)
               
                Case 0
                Destino = "Nix"
                Ciudad.map = 34 'Nix
                Ciudad.X = 40
                Ciudad.Y = 87

                
            End Select
            Call WarpUserChar(UserIndex, Ciudad.map, Ciudad.X, Ciudad.Y, True)
        
        Exit Sub
                
    Case eOBJType.otNudillos
        If ClasePuedeUsarItem(UserIndex, ObjIndex) Then
            'Si esta equipado lo quita
            If UserList(UserIndex).Invent.Object(Slot).Equipped Then
                'Quitamos del inv el item
                Call Desequipar(UserIndex, Slot)
                'Animacion por defecto
                UserList(UserIndex).Char.WeaponAnim = NingunArma
                Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
                Exit Sub
            End If
            
            'Quitamos el arma si tiene alguna equipada
            If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                Call Desequipar(UserIndex, UserList(UserIndex).Invent.WeaponEqpSlot)
            End If
            
            If UserList(UserIndex).Invent.NudiEqpIndex > 0 Then
                Call Desequipar(UserIndex, UserList(UserIndex).Invent.NudiEqpSlot)
            End If
            
            UserList(UserIndex).Invent.Object(Slot).Equipped = 1
            UserList(UserIndex).Invent.NudiEqpIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
            UserList(UserIndex).Invent.NudiEqpSlot = Slot

            UserList(UserIndex).Char.WeaponAnim = Obj.WeaponAnim
            Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
       Else
            Call WriteConsoleMsg(1, UserIndex, "Tu clase no puede usar este objeto.", FontTypeNames.FONTTYPE_INFO)
       End If
        
End Select

'Actualiza
Call UpdateUserInv(False, UserIndex, Slot)

Exit Sub
Errhandler:
Call LogError("EquiparInvItem Slot:" & Slot & " - Error: " & Err.Number & " - Error Description : " & Err.description)
End Sub

 Function CheckRazaUsaRopa(ByVal UserIndex As Integer, ItemIndex As Integer) As Boolean
On Error GoTo Errhandler

If ObjData(ItemIndex).RazaTipo > 0 Then
    'Verifica si la raza puede usar la ropa
    If UserList(UserIndex).Raza = eRaza.Humano Or _
       UserList(UserIndex).Raza = eRaza.Elfo Or _
       UserList(UserIndex).Raza = eRaza.Orco Or _
       UserList(UserIndex).Raza = eRaza.Drow Then
            CheckRazaUsaRopa = (ObjData(ItemIndex).RazaTipo = 1)
    ElseIf UserList(UserIndex).Raza = eRaza.Orco Then
        CheckRazaUsaRopa = (ObjData(ItemIndex).RazaTipo = 3)
    Else
        CheckRazaUsaRopa = (ObjData(ItemIndex).RazaTipo = 2)
    End If
Else
    'Verifica si la raza puede usar la ropa
    If UserList(UserIndex).Raza = eRaza.Humano Or _
       UserList(UserIndex).Raza = eRaza.Elfo Or _
       UserList(UserIndex).Raza = eRaza.Drow Or _
       UserList(UserIndex).Raza = eRaza.Orco Then
            CheckRazaUsaRopa = (ObjData(ItemIndex).RazaEnana = 0)
    Else
        CheckRazaUsaRopa = (ObjData(ItemIndex).RazaEnana = 1)
    End If
End If

Exit Function
Errhandler:
    Call LogError("Error CheckRazaUsaRopa ItemIndex:" & ItemIndex)

End Function

Sub UseInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte)
'*************************************************
'Author: Unknown
'Last modified: 24/01/2007
'Handels the usage of items from inventory box.
'24/01/2007 Pablo (ToxicWaste) - Agrego el Cuerno de la Armada y la Legi�n.
'24/01/2007 Pablo (ToxicWaste) - Utilizaci�n nueva de Barco en lvl 20 por clase Pirata y Pescador.
'*************************************************

Dim Obj As ObjData
Dim ObjIndex As Integer
Dim TargObj As ObjData
Dim MiObj As Obj

With UserList(UserIndex)

If .Invent.Object(Slot).amount = 0 Then Exit Sub

Obj = ObjData(.Invent.Object(Slot).ObjIndex)

If Not Obj.MinELV = 0 Then
    If Obj.MinELV > .Stats.ELV Then
        Call WriteConsoleMsg(1, UserIndex, "Debes ser nivel " & Obj.MinELV & " para poder utilizar este objeto.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
End If

If Obj.Newbie = 1 And Not EsNewbie(UserIndex) Then
    Call WriteConsoleMsg(1, UserIndex, "Solo los newbies pueden usar estos objetos.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

If Obj.OBJType = eOBJType.otWeapon Then
    If Obj.proyectil = 1 Then
        If Not .flags.ModoCombate Then
            Call WriteConsoleMsg(1, UserIndex, "No est�s en modo de combate, presiona la tecla ""C"" para pasar al modo combate.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'valido para evitar el flood pero no bloqueo. El bloqueo se hace en WLC con proyectiles.
        If Not IntervaloPermiteUsar(UserIndex, False) Then Exit Sub
    Else
        'dagas
        If Not IntervaloPermiteUsar(UserIndex) Then Exit Sub
    End If
Else
    If Not IntervaloPermiteUsar(UserIndex) Then Exit Sub
End If

ObjIndex = .Invent.Object(Slot).ObjIndex
.flags.TargetObjInvIndex = ObjIndex
.flags.TargetObjInvSlot = Slot

Select Case Obj.OBJType
    Case eOBJType.otUseOnce
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(1, UserIndex, "��Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        Case eOBJType.OtPOTSINVI
If UserList(UserIndex).flags.Muerto = 1 Then
Call WriteConsoleMsg(1, UserIndex, "Solo puedes usarla si te encuentras vivo", FontTypeNames.FONTTYPE_GUILD)
Exit Sub
End If
If UserList(UserIndex).flags.invisible = 1 Then
Call WriteConsoleMsg(1, UserIndex, "��Ya estas Invisible joder!!", FontTypeNames.FONTTYPE_GUILD)
Exit Sub
Else
UserList(UserIndex).flags.invisible = 0
UserList(UserIndex).flags.invisible = 1
Call WriteConsoleMsg(1, UserIndex, "Te has echo invisible", FontTypeNames.FONTTYPE_GUILD)
        Call QuitarUserInvItem(UserIndex, Slot, 1)
        Call UpdateUserInv(False, UserIndex, Slot)
End If 'Mayco Goy
        
        'Usa el item
        .Stats.MinHam = .Stats.MinHam + Obj.MinHam
        If .Stats.MinHam > .Stats.MaxHam Then _
            .Stats.MinHam = .Stats.MaxHam
        .flags.Hambre = 0
        Call WriteUpdateHungerAndThirst(UserIndex)
        'Sonido
        
        If ObjIndex = e_ObjetosCriticos.Manzana Or ObjIndex = e_ObjetosCriticos.Manzana2 Or ObjIndex = e_ObjetosCriticos.ManzanaNewbie Then
            Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, e_SoundIndex.MORFAR_MANZANA)
        Else
            Call SonidosMapas.ReproducirSonido(SendTarget.ToPCArea, UserIndex, e_SoundIndex.SOUND_COMIDA)
        End If
        
        'Quitamos del inv el item
        Call QuitarUserInvItem(UserIndex, Slot, 1)
        
        Call UpdateUserInv(False, UserIndex, Slot)

    Case eOBJType.otGuita
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(1, UserIndex, "��Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        .Stats.GLD = .Stats.GLD + .Invent.Object(Slot).amount
        .Invent.Object(Slot).amount = 0
        .Invent.Object(Slot).ObjIndex = 0
        .Invent.NroItems = .Invent.NroItems - 1
        
        Call UpdateUserInv(False, UserIndex, Slot)
        Call WriteUpdateGold(UserIndex)
        
    Case eOBJType.otWeapon
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(1, UserIndex, "��Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Not .Stats.MinSta > 0 Then
            If .genero = eGenero.Hombre Then
                Call WriteConsoleMsg(1, UserIndex, "Estas muy cansado", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(1, UserIndex, "Estas muy cansada", FontTypeNames.FONTTYPE_INFO)
            End If
            Exit Sub
        End If
        
        
        If ObjData(ObjIndex).proyectil = 1 Then
            If .Invent.Object(Slot).Equipped = 0 Then
                Call WriteConsoleMsg(1, UserIndex, "Antes de usar la herramienta deberias equipartela.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            'liquid: muevo esto aca adentro, para que solo pida modo combate si estamos por usar el arco
            If Not .flags.ModoCombate Then
                Call WriteConsoleMsg(1, UserIndex, "No est�s en modo de combate, presiona la tecla ""C"" para pasar al modo combate.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            Call WriteWorkRequestTarget(UserIndex, Proyectiles)
        Else
            If .flags.TargetObj = Le�a Then
                If .Invent.Object(Slot).ObjIndex = DAGA Then
                    If .Invent.Object(Slot).Equipped = 0 Then
                        Call WriteConsoleMsg(1, UserIndex, "Antes de usar la herramienta deberias equipartela.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    Call TratarDeHacerFogata(.flags.TargetObjMap, _
                         .flags.TargetObjX, .flags.TargetObjY, UserIndex)
                End If
            End If
        End If

        
        Select Case ObjIndex
            Case CA�A_PESCA, RED_PESCA
                Call WriteWorkRequestTarget(UserIndex, eSkill.Pesca)
            Case HACHA_LE�ADOR
                Call WriteWorkRequestTarget(UserIndex, eSkill.Talar)
            Case PIQUETE_MINERO
                Call WriteWorkRequestTarget(UserIndex, eSkill.Mineria)
            Case MARTILLO_HERRERO
                Call WriteWorkRequestTarget(UserIndex, eSkill.Herreria)
            Case SERRUCHO_CARPINTERO
                Call EnivarObjConstruibles(UserIndex)
                Call WriteShowCarpenterForm(UserIndex)
        End Select
        
    
    Case eOBJType.otPociones
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(1, UserIndex, "��Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Not IntervaloPermiteGolpeUsar(UserIndex, False) Then
            Call WriteConsoleMsg(1, UserIndex, "��Debes esperar unos momentos para tomar otra poci�n!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        .flags.TomoPocion = True
        .flags.TipoPocion = Obj.TipoPocion
                
        Select Case .flags.TipoPocion
        
            Case 1 'Modif la agilidad
                .flags.DuracionEfecto = Obj.DuracionEfecto
        
                'Usa el item
                .Stats.UserAtributos(eAtributos.Agilidad) = .Stats.UserAtributos(eAtributos.Agilidad) + RandomNumber(Obj.MinModificador, Obj.MaxModificador)
                If .Stats.UserAtributos(eAtributos.Agilidad) > MAXATRIBUTOS Then _
                    .Stats.UserAtributos(eAtributos.Agilidad) = MAXATRIBUTOS
  
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                
                Call WriteAgilidad(UserIndex)
                
            Case 2 'Modif la fuerza
                .flags.DuracionEfecto = Obj.DuracionEfecto
        
                'Usa el item
                .Stats.UserAtributos(eAtributos.Fuerza) = .Stats.UserAtributos(eAtributos.Fuerza) + RandomNumber(Obj.MinModificador, Obj.MaxModificador)
                If .Stats.UserAtributos(eAtributos.Fuerza) > MAXATRIBUTOS Then _
                    .Stats.UserAtributos(eAtributos.Fuerza) = MAXATRIBUTOS
   
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                
                Call WriteFuerza(UserIndex)
                
            Case 3 'Pocion roja, restaura HP
                'Usa el item
                .Stats.MinHP = .Stats.MinHP + RandomNumber(Obj.MinModificador, Obj.MaxModificador)
                If .Stats.MinHP > .Stats.MaxHP Then _
                    .Stats.MinHP = .Stats.MaxHP
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
            
            Case 4 'Pocion azul, restaura MANA
                'Usa el item
                'nuevo calculo para recargar mana
                .Stats.MinMAN = .Stats.MinMAN + Porcentaje(.Stats.MaxMAN, 4) + .Stats.ELV \ 2 + 40 / .Stats.ELV
                If .Stats.MinMAN > .Stats.MaxMAN Then _
                    .Stats.MinMAN = .Stats.MaxMAN
                
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
                
            Case 5  'Pocion Remover Paralizis
If .flags.Paralizado = 1 Then
.flags.Paralizado = 0
Call WriteParalizeOK(UserIndex)
Call WriteConsoleMsg(1, UserIndex, "Pierdes el estado de paralisis", FontTypeNames.FONTTYPE_INFO)
'Sacamos el item del inventario
Call QuitarUserInvItem(UserIndex, Slot, 1)
Call UpdateUserInv(False, UserIndex, Slot)
Else
Call WriteConsoleMsg(1, UserIndex, "Solo puedes utilizarla si estas paralizado", FontTypeNames.FONTTYPE_INFO)
End If
            
            Case 6  ' Pocion Negra
                 UserList(UserIndex).Stats.UserSkills(eSkill.Liderazgo) = 100
                    Call QuitarUserInvItem(UserIndex, Slot, 1)
                    
            Case 7  'Pocion Remover Paralizis
                 Call WriteConsoleMsg(1, UserIndex, "Solo puedes usarla si te encuentras vivo", FontTypeNames.FONTTYPE_GUILD)
Exit Sub
If .flags.Paralizado = 1 Then
.flags.Paralizado = 0
Call WriteParalizeOK(UserIndex)
Call WriteConsoleMsg(1, UserIndex, "Pierdes el estado de paralisis", FontTypeNames.FONTTYPE_INFO)
'Sacamos el item del inventario
Call QuitarUserInvItem(UserIndex, Slot, 1)
Call UpdateUserInv(False, UserIndex, Slot)
Else
Call WriteConsoleMsg(1, UserIndex, "Solo puedes utilizarla si estas paralizado", FontTypeNames.FONTTYPE_INFO)
End If
                    
                    
                 End Select
       Call WriteUpdateUserStats(UserIndex)
       Call UpdateUserInv(False, UserIndex, Slot)

     Case eOBJType.otBebidas
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(1, UserIndex, "��Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        .Stats.MinAGU = .Stats.MinAGU + Obj.MinSed
        If .Stats.MinAGU > .Stats.MaxAGU Then _
            .Stats.MinAGU = .Stats.MaxAGU
        .flags.Sed = 0
        Call WriteUpdateHungerAndThirst(UserIndex)
        
        'Quitamos del inv el item
        Call QuitarUserInvItem(UserIndex, Slot, 1)
        
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
        
        Call UpdateUserInv(False, UserIndex, Slot)
    
    Case eOBJType.otLlaves
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(1, UserIndex, "��Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If .flags.TargetObj = 0 Then Exit Sub
        TargObj = ObjData(.flags.TargetObj)
        '�El objeto clickeado es una puerta?
        If TargObj.OBJType = eOBJType.otPuertas Then
            '�Esta cerrada?
            If TargObj.Cerrada = 1 Then
                  '�Cerrada con llave?
                  If TargObj.Llave > 0 Then
                     If TargObj.clave = Obj.clave Then
         
                        MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex _
                        = ObjData(MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex).IndexCerrada
                        .flags.TargetObj = MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex
                        Call WriteConsoleMsg(1, UserIndex, "Has abierto la puerta.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                     Else
                        Call WriteConsoleMsg(1, UserIndex, "La llave no sirve.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                     End If
                  Else
                     If TargObj.clave = Obj.clave Then
                        MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex _
                        = ObjData(MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex).IndexCerradaLlave
                        Call WriteConsoleMsg(1, UserIndex, "Has cerrado con llave la puerta.", FontTypeNames.FONTTYPE_INFO)
                        .flags.TargetObj = MapData(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY).ObjInfo.ObjIndex
                        Exit Sub
                     Else
                        Call WriteConsoleMsg(1, UserIndex, "La llave no sirve.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                     End If
                  End If
            Else
                  Call WriteConsoleMsg(1, UserIndex, "No esta cerrada.", FontTypeNames.FONTTYPE_INFO)
                  Exit Sub
            End If
        End If
    
    Case eOBJType.otBotellaVacia
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(1, UserIndex, "��Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If Not HayAgua(.Pos.map, .flags.TargetX, .flags.TargetY) Then
            Call WriteConsoleMsg(1, UserIndex, "No hay agua all�.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        MiObj.amount = 1
        MiObj.ObjIndex = ObjData(.Invent.Object(Slot).ObjIndex).IndexAbierta
        Call QuitarUserInvItem(UserIndex, Slot, 1)
        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(.Pos, MiObj)
        End If
        
        Call UpdateUserInv(False, UserIndex, Slot)
    
    Case eOBJType.otBotellaLlena
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(1, UserIndex, "��Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        .Stats.MinAGU = .Stats.MinAGU + Obj.MinSed
        If .Stats.MinAGU > .Stats.MaxAGU Then _
            .Stats.MinAGU = .Stats.MaxAGU
        .flags.Sed = 0
        Call WriteUpdateHungerAndThirst(UserIndex)
        MiObj.amount = 1
        MiObj.ObjIndex = ObjData(.Invent.Object(Slot).ObjIndex).IndexCerrada
        Call QuitarUserInvItem(UserIndex, Slot, 1)
        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(.Pos, MiObj)
        End If
        
        Call UpdateUserInv(False, UserIndex, Slot)
    
    Case eOBJType.otPergaminos
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(1, UserIndex, "��Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If .Stats.MaxMAN > 0 Then
            If .flags.Hambre = 0 And _
                .flags.Sed = 0 Then
                Call AgregarHechizo(UserIndex, Slot)
                Call UpdateUserInv(False, UserIndex, Slot)
            Else
                Call WriteConsoleMsg(1, UserIndex, "Estas demasiado hambriento y sediento.", FontTypeNames.FONTTYPE_INFO)
            End If
        Else
            Call WriteConsoleMsg(1, UserIndex, "No tienes conocimientos de las Artes Arcanas.", FontTypeNames.FONTTYPE_INFO)
        End If
    Case eOBJType.otMinerales
        If .flags.Muerto = 1 Then
             Call WriteConsoleMsg(1, UserIndex, "��Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
             Exit Sub
        End If
        Call WriteWorkRequestTarget(UserIndex, FundirMetal)
       
    Case eOBJType.otInstrumentos
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(1, UserIndex, "��Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Obj.Real Then '�Es el Cuerno Real?
            If FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
                If MapInfo(.Pos.map).Pk = False Then
                    Call WriteConsoleMsg(1, UserIndex, "No hay Peligro aqu�. Es Zona Segura ", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                Call SendData(SendTarget.toMap, .Pos.map, PrepareMessagePlayWave(Obj.Snd1, .Pos.X, .Pos.Y))
                Exit Sub
            Else
                Call WriteConsoleMsg(1, UserIndex, "Solo Miembros de la Armada Real pueden usar este cuerno.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        ElseIf Obj.Caos Then '�Es el Cuerno Legi�n?
            If FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
                If MapInfo(.Pos.map).Pk = False Then
                    Call WriteConsoleMsg(1, UserIndex, "No hay Peligro aqu�. Es Zona Segura ", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                Call SendData(SendTarget.toMap, .Pos.map, PrepareMessagePlayWave(Obj.Snd1, .Pos.X, .Pos.Y))
                Exit Sub
            Else
                Call WriteConsoleMsg(1, UserIndex, "Solo Miembros de la Legi�n Oscura pueden usar este cuerno.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        'Si llega aca es porque es o Laud o Tambor o Flauta
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Obj.Snd1, .Pos.X, .Pos.Y))
       
    Case eOBJType.otMonturas
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call WriteConsoleMsg(1, UserIndex, "��Estas muerto!! Los muertos no dominan los animales. ", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            Call DoEquita(UserIndex, ObjIndex, Slot)
            
    Case eOBJType.otBarcos
        If ((LegalPos(.Pos.map, .Pos.X - 1, .Pos.Y, True, False) _
                Or LegalPos(.Pos.map, .Pos.X, .Pos.Y - 1, True, False) _
                Or LegalPos(.Pos.map, .Pos.X + 1, .Pos.Y, True, False) _
                Or LegalPos(.Pos.map, .Pos.X, .Pos.Y + 1, True, False)) _
                And .flags.Navegando = 0) _
                Or .flags.Navegando = 1 Then
            Call DoNavega(UserIndex, Obj, Slot)
        Else
            Call WriteConsoleMsg(1, UserIndex, "�Debes aproximarte al agua para usar el barco!", FontTypeNames.FONTTYPE_INFO)
        End If
        
    Case eOBJType.otPasajes
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteConsoleMsg(1, UserIndex, "��Estas muerto!! Solo podes usar items cuando estas vivo. ", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
       
        If UserList(UserIndex).flags.TargetNPC <> 0 Then
            If Left$(Npclist(UserList(UserIndex).flags.TargetNPC).name, 6) <> "Pirata" Then
                Call WriteConsoleMsg(1, UserIndex, "Primero debes hacer click sobre el pirata.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        Else
            Call WriteConsoleMsg(1, UserIndex, "Primero debes hacer click sobre el pirata.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
       
        If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 3 Then
            Call WriteConsoleMsg(1, UserIndex, "Est�s demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        If Not MapaValido(Obj.HastaMap) Then
            Call WriteConsoleMsg(1, UserIndex, "El pasaje lleva hacia un mapa que ya no esta disponible! Disculpa las molestias.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
       
        If UserList(UserIndex).Stats.UserSkills(5) < Obj.CantidadSkill Then
            Call WriteConsoleMsg(1, UserIndex, "Debido a la peligrosidad del viaje, no puedo llevarte, ya que al menos necesitas saber manejar una barca.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
       
        Call WarpUserChar(UserIndex, Obj.HastaMap, Obj.HastaX, Obj.HastaY, True)
        Call WriteConsoleMsg(1, UserIndex, "Has viajado por varios d�as, te sientes exhausto!", FontTypeNames.FONTTYPE_CENTINELA)
        UserList(UserIndex).Stats.MinAGU = 0
        UserList(UserIndex).Stats.MinHam = 0
        UserList(UserIndex).flags.Sed = 1
        UserList(UserIndex).flags.Hambre = 1
        Call WriteUpdateHungerAndThirst(UserIndex)
        Call QuitarUserInvItem(UserIndex, Slot, 1)
        Call UpdateUserInv(False, UserIndex, Slot)
End Select

End With

End Sub

Sub EnivarArmasConstruibles(ByVal UserIndex As Integer)

Call WriteBlacksmithWeapons(UserIndex)

End Sub
 
Sub EnivarObjConstruibles(ByVal UserIndex As Integer)

Call WriteCarpenterObjects(UserIndex)

End Sub

Sub EnivarArmadurasConstruibles(ByVal UserIndex As Integer)

Call WriteBlacksmithArmors(UserIndex)

End Sub

Sub TirarTodo(ByVal UserIndex As Integer)
On Error Resume Next

If MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).Trigger = 6 Then Exit Sub
If MapInfo(UserList(UserIndex).Pos.map).Seguro = 1 Then Exit Sub

Call TirarTodosLosItems(UserIndex)

Dim Cantidad As Long
Cantidad = UserList(UserIndex).Stats.GLD

If UserList(UserIndex).Stats.GLD < 100001 Then _
    Call TirarOro(UserList(UserIndex).Stats.GLD, UserIndex)

End Sub
Public Function ItemSeCae(ByVal index As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
 
    With ObjData(index)
        ItemSeCae = (.Real <> 1 Or .NoSeCae = 0) And _
                    (.Caos <> 1 Or .NoSeCae = 0) And _
                    .OBJType <> eOBJType.otLlaves And _
                    .OBJType <> eOBJType.otBarcos And _
                    .NoSeCae = 0
    End With
 
End Function

Sub TirarTodosLosItems(ByVal UserIndex As Integer)
    Dim i As Byte
    Dim NuevaPos As WorldPos
    Dim MiObj As Obj
    Dim ItemIndex As Integer
    
    For i = 1 To MAX_INVENTORY_SLOTS
        ItemIndex = UserList(UserIndex).Invent.Object(i).ObjIndex
        If ItemIndex > 0 Then
             If ItemSeCae(ItemIndex) Then
                NuevaPos.X = 0
                NuevaPos.Y = 0
                
                'Creo el Obj
                MiObj.amount = UserList(UserIndex).Invent.Object(i).amount
                MiObj.ObjIndex = ItemIndex
                'Pablo (ToxicWaste) 24/01/2007
                'Si es pirata y usa un Gale�n entonces no explota los items. (en el agua)
                If UserList(UserIndex).Clase = eClass.Mercenario And UserList(UserIndex).Invent.BarcoObjIndex = 476 Then
                    Tilelibre UserList(UserIndex).Pos, NuevaPos, MiObj, False, True
                Else
                    Tilelibre UserList(UserIndex).Pos, NuevaPos, MiObj, True, True
             
             
Dim Sacrificio As Obj
Sacrificio.ObjIndex = SACRI 'no tocar.
Sacrificio.amount = 1 'modificarp or la cantidad a tirar.
If UserList(UserIndex).Invent.AnilloEqpObjIndex = SACRI Then
Call WriteConsoleMsg(1, UserIndex, "Pierdes tu pendiente, pero tu inventario sigue intacto", FontTypeNames.FONTTYPE_INFO)
Call TirarItemAlPiso(UserList(UserIndex).Pos, Sacrificio)
Call QuitarObjetos(SACRI, 1, UserIndex)
Exit Sub
End If
                    
                End If
                
                If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                    Call DropObj(UserIndex, i, MAX_INVENTORY_OBJS, NuevaPos.map, NuevaPos.X, NuevaPos.Y)
                End If
             End If
        End If
    Next i
End Sub

Function ItemNewbie(ByVal ItemIndex As Integer) As Boolean

ItemNewbie = ObjData(ItemIndex).Newbie = 1

End Function

Sub TirarTodosLosItemsNoNewbies(ByVal UserIndex As Integer)
Dim i As Byte
Dim NuevaPos As WorldPos
Dim MiObj As Obj
Dim ItemIndex As Integer

If MapData(UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).Trigger = 6 Then Exit Sub

For i = 1 To MAX_INVENTORY_SLOTS
    ItemIndex = UserList(UserIndex).Invent.Object(i).ObjIndex
    If ItemIndex > 0 Then
        If ItemSeCae(ItemIndex) And Not ItemNewbie(ItemIndex) Then
            NuevaPos.X = 0
            NuevaPos.Y = 0
            
            'Creo MiObj
            MiObj.amount = UserList(UserIndex).Invent.Object(i).ObjIndex
            MiObj.ObjIndex = ItemIndex
            'Pablo (ToxicWaste) 24/01/2007
            'Tira los Items no newbies en todos lados.
            Tilelibre UserList(UserIndex).Pos, NuevaPos, MiObj, True, True
            If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                If MapData(NuevaPos.map, NuevaPos.X, NuevaPos.Y).ObjInfo.ObjIndex = 0 Then Call DropObj(UserIndex, i, MAX_INVENTORY_OBJS, NuevaPos.map, NuevaPos.X, NuevaPos.Y)
            End If
        End If
    End If
Next i

End Sub