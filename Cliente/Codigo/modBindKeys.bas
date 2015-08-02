Attribute VB_Name = "modBindKeys"
'*****************************************************************
'modBindKeys - ImperiumAO - v1.4.5 R5
'
'User input functions.
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
'Augusto José Rando (barrin@imperiumao.com.ar)
'   - First Relase
'*****************************************************************

Option Explicit

Type tBoton
    TipoAccion As Integer
    SendString As String
    hlist As Integer
    invslot As Integer
End Type

Type tBindedKey
    KeyCode As Integer
    VirtualKey As Long
    name As String
End Type

Public NUMBOTONES As Integer
Public NUMBINDS As Integer

Public MacroKeys() As tBoton
Public BindKeys() As tBindedKey
Public BotonElegido As Integer

Public Function Accionar() As Boolean

    Accionar = True

    If frmMain.Input_Keyboard_KeyDOWN(DIK_MULTIPLY) Then
        frmMain.Engine.Engine_Stats_Show_Toggle
    
    '84 = PrintScreen = vbKeySnapshot = DIK_SYSRQ
    ElseIf frmMain.Engine.Input_Keyboard_KeyDOWN(DIK_SYSRQ) Then
        Call frmMain.Client_Screenshot(frmMain.hdc, 800, 600)

    ElseIf frmMain.Engine.Input_Keyboard_KeyDOWN(DIK_F12) Then
        Call frmMain.Client_Screenshot(frmMain.MainViewPic.hdc, frmMain.MainViewPic.width, frmMain.MainViewPic.height)

    'SINU-RECORDER
    ElseIf frmMain.Engine.Input_Keyboard_KeyDOWN(DIK_PAUSE) Then

    ElseIf (frmMain.Engine.Input_Keyboard_Last_KeyDOWN(BindKeys(1).VirtualKey) And frmMain.Engine.Input_Keyboard_KeyUP(BindKeys(1).VirtualKey)) Then
           If (Not CurrentUser.Descansando) And _
           (Not CurrentUser.Meditando) Then
                If ClientTCP.DeadCheck Then Exit Function
                If ClientTCP.CombateCheck Then Exit Function
                If IntervaloPermiteAtacar Then Call ClientTCP.Send_Data(Attack, Integer_To_String(CurrentUser.UserMinSTA))
        End If
    
    ElseIf frmMain.Engine.Input_Keyboard_KeyDOWN(BindKeys(2).VirtualKey) Then
        If Not CurrentUser.Comerciando Then
            If ClientTCP.DeadCheck Then Exit Function
            Call AgarrarItem
        Else
            Call PrintToConsole(Locale_GUI_Frase(236), 255, 0, 32, False, False, False)
        End If
    
    ElseIf frmMain.Engine.Input_Keyboard_KeyDOWN(BindKeys(3).VirtualKey) Then
        If Not CurrentUser.Comerciando Then
            If ClientTCP.DeadCheck Then Exit Function
            Call TirarItem
        Else
            Call PrintToConsole(Locale_GUI_Frase(236), 255, 0, 32, False, False, False)
        End If
    
    ElseIf frmMain.Engine.Input_Keyboard_KeyDOWN(BindKeys(6).VirtualKey) Then
        Call ClientTCP.Send_Data(Safe_Mode)
        CurrentUser.Seguro = Not CurrentUser.Seguro
        frmMain.modoseguro.Visible = Not frmMain.modoseguro.Visible
        frmMain.nomodoseguro.Visible = Not frmMain.nomodoseguro.Visible
    
    ElseIf frmMain.Engine.Input_Keyboard_KeyDOWN(BindKeys(12).VirtualKey) Then
        Call ClientTCP.Send_Data(Combat_Mode)
        CurrentUser.Combate = Not CurrentUser.Combate
        frmMain.modocombate.Visible = Not frmMain.modocombate.Visible
        frmMain.nomodocombate.Visible = Not frmMain.nomodocombate.Visible

    ElseIf frmMain.Engine.Input_Keyboard_KeyDOWN(BindKeys(7).VirtualKey) Then
        frmMain.Engine.Engine_Label_Render_Set
    
    ElseIf frmMain.Engine.Input_Keyboard_KeyDOWN(BindKeys(8).VirtualKey) Then
        If ClientTCP.DeadCheck Then Exit Function
        Call ClientTCP.Send_Data(Working_Click, Byte_To_String(Domar) & Integer_To_String(CurrentUser.UserMinSTA))
        
    ElseIf frmMain.Engine.Input_Keyboard_KeyDOWN(BindKeys(9).VirtualKey) Then
        If ClientTCP.DeadCheck Then Exit Function
        If ClientTCP.StealCheck Then Exit Function
        Call ClientTCP.Send_Data(Working_Click, Byte_To_String(Robar) & Integer_To_String(CurrentUser.UserMinSTA))
    
    ElseIf frmMain.Engine.Input_Keyboard_KeyDOWN(BindKeys(5).VirtualKey) Then
        If ClientTCP.DeadCheck Then Exit Function
        Call EquiparItem
    
    ElseIf frmMain.Engine.Input_Keyboard_KeyDOWN(BindKeys(4).VirtualKey) Then
        If IntervaloPermiteUsar Then Call UsarItem
    
    ElseIf frmMain.Engine.Input_Keyboard_KeyDOWN(BindKeys(10).VirtualKey) Then
        If IntervaloPermiteRefrescar Then Call ClientTCP.Send_Data(Request_Pos)
    
    ElseIf frmMain.Engine.Input_Keyboard_KeyDOWN(BindKeys(11).VirtualKey) Then
        Call ClientTCP.Send_Data(Working_Click, Byte_To_String(Ocultarse) & Integer_To_String(CurrentUser.UserMinSTA))
        
    ElseIf frmMain.Engine.Input_Keyboard_KeyDOWN(BindKeys(13).VirtualKey) Then
        Call ClientTCP.Send_Data(Role_Mode)
        CurrentUser.Rol = Not CurrentUser.Rol
        frmMain.modorol.Visible = Not frmMain.modorol.Visible
        frmMain.nomodorol.Visible = Not frmMain.nomodorol.Visible
    Else
        Accionar = False
        Exit Function
    End If

End Function

Sub TirarItem()
    If (ItemElegido > 0 And ItemElegido <= MAX_INVENTORY_SLOTS) Or (ItemElegido = FLAGORO) Then
        frmCantidad.Show vbModeless, frmMain
    End If
End Sub

Sub AgarrarItem()
    Call ClientTCP.Send_Data(Get_Item)
End Sub

Sub UsarItem()
    If (ItemElegido > 0) And (ItemElegido <= MAX_INVENTORY_SLOTS) And Not ClientTCP.MeditarCheck() Then Call ClientTCP.Send_Data(Use_Item, Integer_To_String(ItemElegido))
End Sub

Sub EquiparItem()
    If (ItemElegido > 0) And (ItemElegido <= MAX_INVENTORY_SLOTS) Then
        If Not ClientTCP.MeditarCheck() And Not ClientTCP.DeadCheck() Then Call ClientTCP.Send_Data(Equip_Item, Integer_To_String(ItemElegido))
    End If
End Sub

Sub LoadDefaultBinds()

Dim Arch As String, lc As Integer
Arch = App.Path & "\init\" & "ImpAoInit.bnd"

NUMBINDS = Val(General_Var_Get(Arch, "INIT", "NumBinds"))
ReDim BindKeys(1 To NUMBINDS) As tBindedKey

For lc = 1 To NUMBINDS
    BindKeys(lc).KeyCode = Val(General_Field_Read(1, General_Var_Get(Arch, "DEFAULTS", str(lc)), ","))
    BindKeys(lc).name = General_Field_Read(2, General_Var_Get(Arch, "DEFAULTS", str(lc)), ",")
    BindKeys(lc).VirtualKey = MapVirtualKey(BindKeys(lc).KeyCode, 0)
    
    If BindKeys(lc).VirtualKey = DIK_NUMPAD4 Then
        BindKeys(lc).VirtualKey = DIK_LEFTARROW
    ElseIf BindKeys(lc).VirtualKey = DIK_NUMPAD6 Then
        BindKeys(lc).VirtualKey = DIK_RIGHTARROW
    ElseIf BindKeys(lc).VirtualKey = DIK_NUMPAD8 Then
        BindKeys(lc).VirtualKey = DIK_UPARROW
    ElseIf BindKeys(lc).VirtualKey = DIK_NUMPAD2 Then
        BindKeys(lc).VirtualKey = DIK_DOWNARROW
    End If
    
Next lc

End Sub

Public Sub MouseLeftClick(ByVal tX As Integer, ByVal tY As Integer)

Dim char_index As Integer
Dim char_name As String

If GetKeyState(vbKeyShift) < 0 Then
    Call ClientTCP.Send_Data_Command_GM(cmdTeleploc, Integer_To_String(tX) & Integer_To_String(tY))
    Exit Sub
End If

If CurrentUser.UsingSkill = 0 Then
    If Not ClientTCP.DeadCheck Then Call ClientTCP.Send_Data(Left_Click, Integer_To_String(tX) & Integer_To_String(tY))
Else
    Select Case CurrentUser.UsingSkill
        Case Magia
            If Not IntervaloPermiteLanzarSpell Then Exit Sub
        Case Proyectiles, Arrojadizas
            If Not IntervaloPermiteAtacar Then Exit Sub
        Case Domar
            If Not IntervaloPermiteTrabajar Then Exit Sub
        Case GM_SELECT
            char_index = frmMain.Engine.Map_Char_Get(tX, tY)
            
            If char_index > 0 Then
                If frmPanelGm.Visible Then
                    char_name = frmMain.Engine.Char_Name_Get(char_index)
                    char_name = IIf(InStr(1, char_name, "<") > 0, RTrim$(General_Field_Read(1, char_name, "<")), char_name)
                    frmPanelGm.cboListaUsus.Text = char_name
                End If
            End If
        Case Else
            If Not IntervaloPermiteTrabajar Then Exit Sub
    End Select
    
    Call FormParser.Parse_Form(frmMain)
    
    Call ClientTCP.Send_Data(Working_Left_Click, Integer_To_String(tX) & Integer_To_String(tY) & Byte_To_String(CurrentUser.UsingSkill) & Integer_To_String(CurrentUser.UserMinSTA))
    CurrentUser.UsingSkill = 0

End If

End Sub

Public Sub MouseRightClick(ByVal tX As Integer, ByVal tY As Integer)

If CurrentUser.UsingSkill > 0 Then
    Call FormParser.Parse_Form(frmMain)
    CurrentUser.UsingSkill = 0
Else
    Call ClientTCP.Send_Data(Right_Click, Integer_To_String(tX) & Integer_To_String(tY) & Integer_To_String(CurrentUser.UserMinSTA))
End If

End Sub

