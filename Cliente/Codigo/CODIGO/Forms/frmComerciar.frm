VERSION 5.00
Begin VB.Form frmComerciar 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   7665
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6930
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmComerciar.frx":0000
   ScaleHeight     =   511
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   462
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox cantidad 
      Alignment       =   2  'Center
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   8
      Text            =   "1"
      Top             =   6960
      Width           =   510
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Left            =   900
      ScaleHeight     =   480
      ScaleWidth      =   495
      TabIndex        =   2
      Top             =   1620
      Width           =   495
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3960
      Index           =   1
      Left            =   3720
      TabIndex        =   1
      Top             =   2580
      Width           =   2490
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3960
      Index           =   0
      Left            =   735
      TabIndex        =   0
      Top             =   2580
      Width           =   2490
   End
   Begin VB.Image cmdMasMenos 
      Height          =   420
      Index           =   0
      Left            =   2955
      Tag             =   "1"
      Top             =   6885
      Width           =   195
   End
   Begin VB.Image cmdMasMenos 
      Height          =   420
      Index           =   1
      Left            =   3840
      Tag             =   "1"
      Top             =   6885
      Width           =   195
   End
   Begin VB.Image cmdCerrar 
      Height          =   345
      Index           =   1
      Left            =   6480
      Top             =   180
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   1560
      TabIndex        =   7
      Top             =   1560
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Precio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   5400
      TabIndex        =   6
      Top             =   1920
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   5520
      TabIndex        =   5
      Top             =   1560
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Estadisticas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   3000
      TabIndex        =   4
      Top             =   1800
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Estadisticas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   1560
      TabIndex        =   3
      Top             =   1800
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Image Image1 
      Height          =   450
      Index           =   1
      Left            =   4230
      MouseIcon       =   "frmComerciar.frx":2FD54
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   6855
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   450
      Index           =   0
      Left            =   585
      MouseIcon       =   "frmComerciar.frx":2FEA6
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   6855
      Width           =   2175
   End
End
Attribute VB_Name = "frmComerciar"
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

Public LastIndex1 As Integer
Public LastIndex2 As Integer
Public LasActionBuy As Boolean
Private lIndex As Byte

Private Sub cantidad_Change()
    If Val(cantidad.Text) < 1 Then
        cantidad.Text = 1
    End If
    
    If Val(cantidad.Text) > MAX_INVENTORY_OBJS Then
        cantidad.Text = 1
    End If
    
    If lIndex = 0 Then
        If List1(0).ListIndex <> -1 Then
            'El precio, cuando nos venden algo, lo tenemos que redondear para arriba.
            Label1(1).Caption = CalculateSellPrice(NPCInventory(List1(0).ListIndex + 1).Valor, Val(cantidad.Text)) 'No mostramos numeros reales
        End If
    Else
        If List1(1).ListIndex <> -1 Then
            Label1(1).Caption = CalculateBuyPrice(Inventario.Valor(List1(1).ListIndex + 1), Val(cantidad.Text)) 'No mostramos numeros reales
        End If
    End If
End Sub

Private Sub cantidad_KeyPress(KeyAscii As Integer)
If (KeyAscii <> 8) Then
    If (KeyAscii <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
        KeyAscii = 0
    End If
End If
End Sub

Private Sub cmdMasMenos_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Index
    Case 0
        cmdMasMenos(Index).Picture = LoadInterface("menos-down")
        cmdMasMenos(Index).Tag = "1"
    Case 1
        cmdMasMenos(Index).Picture = LoadInterface("mas-down")
        cmdMasMenos(Index).Tag = "1"
        cantidad.Text = Val(cantidad.Text) + 1
End Select
End Sub

Private Sub cmdMasMenos_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Index
    Case 0
        If cmdMasMenos(Index).Tag = "0" Then
            cmdMasMenos(Index).Picture = LoadInterface("menos-over")
            cmdMasMenos(Index).Tag = "1"
        End If
    Case 1
        If cmdMasMenos(Index).Tag = "0" Then
            cmdMasMenos(Index).Picture = LoadInterface("mas-over")
            cmdMasMenos(Index).Tag = "1"
            cantidad.Text = Val(cantidad.Text) - 1
        End If
End Select
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If image1(0).Tag = "1" Then
    image1(0).Picture = Nothing
    image1(0).Tag = "0"
End If

If image1(1).Tag = "1" Then
    image1(1).Picture = Nothing
    image1(1).Tag = "0"
End If

If cmdCerrar(1).Tag = "1" Then
    cmdCerrar(1).Picture = Nothing
    cmdCerrar(1).Tag = "0"
    End If
    
If cmdMasMenos(0).Tag = "1" Then
    cmdMasMenos(0).Picture = Nothing
    cmdMasMenos(0).Tag = "0"
  End If

 If cmdMasMenos(1).Tag = "1" Then
    cmdMasMenos(1).Picture = Nothing
    cmdMasMenos(1).Tag = "0"
End If
End Sub

Private Sub Form_Load()
'Cargamos la interfase
Me.Picture = LoadInterface("comerciar")
image1(0).Picture = LoadInterface("comprar-over")
image1(1).Picture = LoadInterface("vender-over")
cmdCerrar(1).Picture = LoadInterface("salir-over")
Picture1.Cls
End Sub
''
' Calculates the selling price of an item (The price that a merchant will sell you the item)
'
' @param objValue Specifies value of the item.
' @param objAmount Specifies amount of items that you want to buy
' @return   The price of the item.

Private Function CalculateSellPrice(ByRef objValue As Single, ByVal objAmount As Long) As Long
'*************************************************
'Author: Marco Vanotti (MarKoxX)
'Last modified: 19/08/2008
'Last modify by: Franco Zeoli (Noich)
'*************************************************
    On Error GoTo error
    'We get a Single value from the server, when vb uses it, by approaching, it can diff with the server value, so we do (Value * 100000) and get the entire part, to discard the unwanted floating values.
    CalculateSellPrice = CCur(objValue * 1000000) / 1000000 * objAmount + 0.5
    
    Exit Function
error:
    MsgBox Err.Description, vbExclamation, "Error: " & Err.number
End Function
''
' Calculates the buying price of an item (The price that a merchant will buy you the item)
'
' @param objValue Specifies value of the item.
' @param objAmount Specifies amount of items that you want to buy
' @return   The price of the item.
Private Function CalculateBuyPrice(ByRef objValue As Single, ByVal objAmount As Long) As Long
'*************************************************
'Author: Marco Vanotti (MarKoxX)
'Last modified: 19/08/2008
'Last modify by: Franco Zeoli (Noich)
'*************************************************
    On Error GoTo error
    'We get a Single value from the server, when vb uses it, by approaching, it can diff with the server value, so we do (Value * 100000) and get the entire part, to discard the unwanted floating values.
    CalculateBuyPrice = Fix(CCur(objValue * 1000000) / 1000000 * objAmount)
    
    Exit Function
error:
    MsgBox Err.Description, vbExclamation, "Error: " & Err.number
End Function

Private Sub Image1_Click(Index As Integer)

Call Audio.PlayWave(SND_CLICK)

If List1(Index).List(List1(Index).ListIndex) = "" Or _
   List1(Index).ListIndex < 0 Then Exit Sub

If Not IsNumeric(cantidad.Text) Or cantidad.Text = 0 Then Exit Sub

Select Case Index
    Case 0
        frmComerciar.List1(0).SetFocus
        LastIndex1 = List1(0).ListIndex
        LasActionBuy = True
        If UserGLD >= CalculateSellPrice(NPCInventory(List1(0).ListIndex + 1).Valor, Val(cantidad.Text)) Then
            Call WriteCommerceBuy(List1(0).ListIndex + 1, cantidad.Text)
        Else
            AddtoRichTextBox frmMain.RecChat, "No tenés suficiente oro.", 2, 51, 223, 1, 1
            Exit Sub
        End If
   
   Case 1
        LastIndex2 = List1(1).ListIndex
        LasActionBuy = False
        
        Call WriteCommerceSell(List1(1).ListIndex + 1, cantidad.Text)
End Select

End Sub

Private Sub list1_Click(Index As Integer)

lIndex = Index

Select Case Index
    Case 0
        
        Label1(0).Caption = NPCInventory(List1(0).ListIndex + 1).name
        Label1(1).Caption = CalculateSellPrice(NPCInventory(List1(0).ListIndex + 1).Valor, Val(cantidad.Text)) 'No mostramos numeros reales
        Label1(2).Caption = NPCInventory(List1(0).ListIndex + 1).Amount
        
        If Label1(2).Caption <> 0 Then
        
        Select Case NPCInventory(List1(0).ListIndex + 1).OBJType
            Case eObjType.otWeapon
                Label1(3).Caption = "Max Golpe:" & NPCInventory(List1(0).ListIndex + 1).MaxHit
                Label1(4).Caption = "Min Golpe:" & NPCInventory(List1(0).ListIndex + 1).MinHit
                Label1(3).Visible = True
                Label1(4).Visible = True
            Case eObjType.otArmadura
                Label1(3).Visible = False
                Label1(4).Caption = "Defensa:" & NPCInventory(List1(0).ListIndex + 1).Def
                Label1(4).Visible = True
            Case Else
                Label1(3).Visible = False
                Label1(4).Visible = False
        End Select
        
        Picture1.Cls
        Call Engine.DrawGrhToHdc(Picture1.hdc, NPCInventory(List1(0).ListIndex + 1).grhindex, 3, 5)
        
        End If
    
    Case 1
        Label1(0).Caption = Inventario.ItemName(List1(1).ListIndex + 1)
        Label1(1).Caption = CalculateBuyPrice(Inventario.Valor(List1(1).ListIndex + 1), Val(cantidad.Text)) 'No mostramos numeros reales
        Label1(2).Caption = Inventario.Amount(List1(1).ListIndex + 1)
        
        If Label1(2).Caption <> 0 Then
        
        Select Case Inventario.OBJType(List1(1).ListIndex + 1)
            Case eObjType.otWeapon
                Label1(3).Caption = "Max Golpe:" & Inventario.MaxHit(List1(1).ListIndex + 1)
                Label1(4).Caption = "Min Golpe:" & Inventario.MinHit(List1(1).ListIndex + 1)
                Label1(3).Visible = True
                Label1(4).Visible = True
            Case eObjType.otArmadura
                Label1(3).Visible = False
                Label1(4).Caption = "Defensa:" & Inventario.Def(List1(1).ListIndex + 1)
                Label1(4).Visible = True
            Case Else
                Label1(3).Visible = False
                Label1(4).Visible = False
        End Select
        
        Picture1.Cls
        Call Engine.DrawGrhToHdc(Picture1.hdc, Inventario.grhindex(List1(1).ListIndex + 1), 3, 5)
        
        End If
        
End Select

If Label1(2).Caption = 0 Then ' 27/08/2006 - GS > No mostrar imagen ni nada, cuando no ahi nada que mostrar.
    Label1(3).Visible = False
    Label1(4).Visible = False
    Picture1.Visible = False
Else
    Picture1.Visible = True
    Picture1.Refresh
End If

End Sub

'<-------------------------NUEVO-------------------------->
'<-------------------------NUEVO-------------------------->
'<-------------------------NUEVO-------------------------->

Private Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 0 Then
    image1(Index).Picture = LoadInterface("comprar-down")
    image1(Index).Tag = "1"
ElseIf Index = 1 Then
    image1(Index).Picture = LoadInterface("vender-down")
    image1(Index).Tag = "1"
End If
End Sub

Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 0 Then
    If image1(Index).Tag = "0" Then
        image1(Index).Picture = LoadInterface("comprar-over")
        image1(Index).Tag = "1"
    End If
ElseIf Index = 1 Then
    If image1(Index).Tag = "0" Then
        image1(Index).Picture = LoadInterface("vender-over")
        image1(Index).Tag = "1"
    End If
End If
End Sub

Private Sub cmdCerrar_Mouseup(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call WriteCommerceEnd
End Sub

Private Sub cmdCerrar_MouseDown(Button As Integer, Index As Integer, Shift As Integer, X As Single, Y As Single)
cmdCerrar(1).Picture = LoadInterface("salir-down")
cmdCerrar(1).Tag = "1"
End Sub

Private Sub cmdCerrar_MouseMove(Button As Integer, Shift As Integer, Index As Integer, X As Single, Y As Single)

If cmdCerrar(1).Tag = "0" Then
    cmdCerrar(1).Picture = LoadInterface("salir-over")
    cmdCerrar(1).Tag = "1"
End If
End Sub
