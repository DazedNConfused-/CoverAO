VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGrapchicalInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************
'Graphical inventory for ORE 0.4.1
'Designed and implemented by Juan Martín Sotuyo Dodero (Maraxus)
'juansotuyo@hotmail.com
'*****************************************************

'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'(at your option) any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'**************************************************************************

Option Explicit

Dim InventoryOffset As Long             'Number of lines we scrolled down from topmost
Dim InvSelectedItem As Long             'Currently selected item
Dim TempItem As Integer
'Dim InvSurface As DirectDrawSurface7            'DD Surface used to render everything

'Dim UserInventory(1 To MAX_INVENTORY_SLOTS) As Inventory    'User's inventory

Dim WithEvents InventoryWindow As PictureBox    'Placeholder where to render the inventory
Attribute InventoryWindow.VB_VarHelpID = -1

Private Sub Class_Initialize()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'
'***************************************************
End Sub

Private Sub Class_Terminate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'
'***************************************************
End Sub

Public Sub ScrollInventory(ByVal up As Boolean)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Scrolls the graphical inventory up / down
'***************************************************
    'Only move if there is something to show up /down
    If up Then
        If (InventoryOffset + (InventoryWindow.ScaleHeight \ TilePixelHeight)) * (InventoryWindow.ScaleWidth \ TilePixelWidth) < MAX_INVENTORY_SLOTS Then _
            InventoryOffset = InventoryOffset + 1
    Else
        If InventoryOffset > 0 Then _
            InventoryOffset = InventoryOffset - 1
    End If
End Sub

Public Sub SelectGold()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Sets the gold as the currently selected item
'***************************************************
    Dim prevSelectedItem As Long
    
    'Store preivously selected item
    prevSelectedItem = InvSelectedItem
    
    'Select the gold
    InvSelectedItem = FLAGORO
    
End Sub

Public Property Get SelectedItem() As Long
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item index
'***************************************************
    SelectedItem = InvSelectedItem
End Property

Public Property Get MaxHit(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the max hit of the selected item
'***************************************************
    MaxHit = UserInventory(slot).MaxHit
End Property

Public Property Get MinHit(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the min hit of the selected item
'***************************************************
    MinHit = UserInventory(slot).MinHit
End Property

Public Property Get Def(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the defense of the selected item
'***************************************************
    Def = UserInventory(slot).Def
End Property

Public Property Get grhindex(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the grh index of the selected item
'***************************************************
If slot > 20 Or slot < 1 Then Exit Property
    grhindex = UserInventory(slot).grhindex
End Property

Public Property Get Valor(ByVal slot As Byte) As Single
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the value of the selected item
'***************************************************
    Valor = UserInventory(slot).Valor
End Property

Public Property Get Amount(ByVal slot As Byte) As Long
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's amount
'***************************************************
    If slot = FLAGORO Then
        Amount = UserGLD
    ElseIf slot >= LBound(UserInventory) And slot <= UBound(UserInventory) Then
        Amount = UserInventory(slot).Amount
    End If
End Property

Public Property Get OBJIndex(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's object index
'***************************************************
    OBJIndex = UserInventory(slot).OBJIndex
End Property

Public Property Get OBJType(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's object type
'***************************************************
    OBJType = UserInventory(slot).OBJType
End Property

Public Property Get ItemName(ByVal slot As Byte) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's name
'***************************************************
    If Not slot = 0 Then ItemName = UserInventory(slot).name
End Property

Public Property Get Equipped(ByVal slot As Byte) As Boolean
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves True if the item at the given pos is eqiupped
'***************************************************
    Equipped = UserInventory(slot).Equipped
End Property

Public Sub SetItem(ByVal slot As Byte, ByVal eOBJIndex As Integer, ByVal eAmount As Integer, ByVal eEquipped As Byte, _
                        ByVal eGrhIndex As Integer, ByVal eObjType As Integer, ByVal eMaxHit As Integer, ByVal eMinHit As Integer, _
                        ByVal eDef As Integer, ByVal eValor As Single, ByVal eName As String, ByVal Puede As Byte)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/04/06
'Sets all data for a given inventory slot
'***************************************************
    If slot < 1 Or slot > MAX_INVENTORY_SLOTS Then Exit Sub
    
    With UserInventory(slot)
        .Amount = eAmount
        .Def = eDef
        .Equipped = eEquipped
        .grhindex = eGrhIndex
        .MaxHit = eMaxHit
        .MinHit = eMinHit
        .name = eName
        .OBJIndex = eOBJIndex
        .OBJType = eObjType
        .Valor = eValor
        .PuedeUsar = Puede
    End With

End Sub

Private Function ClickItem(ByVal x As Long, ByVal y As Long) As Long
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Selects the item clicked if it's valid and return's it's index
'***************************************************
    Dim TempItem As Long
    Dim temp_x As Long
    Dim temp_y As Long
    
    temp_x = x \ 32
    temp_y = y \ 32
    
    TempItem = temp_x + (temp_y + InventoryOffset) * (InventoryWindow.ScaleWidth \ 32) + 1
    
    'Make sure it's within limits
    If TempItem <= MAX_INVENTORY_SLOTS Then
        'Make sure slot isn't empty
        If UserInventory(TempItem).grhindex Then
            ClickItem = TempItem
        Else
            ClickItem = 0
        End If
    End If
    
    RenderInv = True
End Function

Public Sub DrawInventory()
    Dim i As Byte
    Dim x As Integer
    Dim y As Integer
    Dim temp_array(3) As Long
    
    For i = 1 To UBound(UserInventory)
        If UserInventory(i).grhindex Then
            x = ((i - 1) Mod (InventoryWindow.width / 32)) * 32
            y = ((i - 1) \ (InventoryWindow.width / 32)) * 32
            
            Engine.Long_To_RGB_List temp_array, -1
            
            If Not UserInventory(i).PuedeUsar = 1 Then
                Call Engine.Draw_GrhIndex(UserInventory(i).grhindex, x - 1, y - 1, -65536)
            Else
                Call Engine.Draw_GrhIndex(UserInventory(i).grhindex, x - 1, y - 1, -1)
            End If

            Call Engine.Text_Render(str(UserInventory(i).Amount), x - 5, y - 2, temp_array, 1)
            
            Engine.Long_To_RGB_List temp_array, D3DColorXRGB(255, 0, 0)
            If UserInventory(i).Equipped Then
                Call Engine.Text_Render("+", x + 20, y, temp_array)
            End If
        End If
    Next i
    
    If InvSelectedItem > 0 Then
        Engine.Draw_GrhIndex 2, (((InvSelectedItem - 1) Mod (InventoryWindow.width / 32)) * 32) - 1, (((InvSelectedItem - 1) \ (InventoryWindow.width / 32)) * 32) - 1, -1
    End If
    
    RenderInv = False
End Sub

Public Sub Initialize(ByRef InvPic As PictureBox)
    Set InventoryWindow = InvPic

    InvSelectedItem = ClickItem(1, 1)   'If there is anything there we select the top left item

End Sub

Private Sub InventoryWindow_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Implements the mouse move event of the inventory picture box
'***************************************************
    Dim temp_x As Integer
    Dim temp_y As Integer
    
    temp_x = x \ 32
    temp_y = y \ 32
    
    TempItem = temp_x + (temp_y + InventoryOffset) * (InventoryWindow.ScaleWidth \ 32) + 1
    
    If TempItem <= 25 And TempItem >= 1 Then
        frmMain.lblInvInfo.Visible = True
        If UserInventory(TempItem).name = "" Then
            frmMain.lblInvInfo.Caption = "Nada"
        Else
            frmMain.lblInvInfo.Caption = UserInventory(TempItem).name '& " || " & "Cantidad: " & UserInventory(TempItem).Amount & vbNewLine & "Def:" & UserInventory(TempItem).Def & " MinHit:" & UserInventory(TempItem).MinHit & " MaxHit:" & UserInventory(TempItem).MaxHit
        End If
    End If
End Sub

Private Sub InventoryWindow_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Implements the mouse up event of the inventory picture box
'Check outs which item was clicked
'***************************************************
    'Store previously selected item
    Dim prevSelItem As Long
    
    'Exit if it got outside the control's area
    If x < 0 Or y < 0 Or x > InventoryWindow.width Or y > InventoryWindow.height Then _
        Exit Sub
    
    prevSelItem = InvSelectedItem

    'Get the currently clickced item
    InvSelectedItem = ClickItem(CInt(x), CInt(y))

    RenderInv = True
End Sub

Public Function GetPos(ByVal x As Integer, ByVal y As Integer) As Byte
    Dim temp_x As Integer
    Dim temp_y As Integer
    
    temp_x = x \ 32
    temp_y = y \ 32
    
    GetPos = temp_x + (temp_y + InventoryOffset) * (InventoryWindow.ScaleWidth \ 32) + 1
End Function


