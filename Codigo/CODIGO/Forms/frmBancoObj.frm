VERSION 5.00
Begin VB.Form frmBancoObj 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   7620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6945
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmBancoObj.frx":0000
   ScaleHeight     =   508
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   463
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Left            =   900
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   7
      Top             =   1620
      Width           =   480
   End
   Begin VB.Timer tmrNumber 
      Left            =   120
      Top             =   240
   End
   Begin VB.TextBox Cantidad 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3240
      TabIndex        =   6
      Text            =   "1"
      Top             =   6960
      Width           =   510
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
      Left            =   780
      TabIndex        =   0
      Top             =   2580
      Width           =   2490
   End
   Begin VB.Image cmdMasMenos 
      Height          =   420
      Index           =   0
      Left            =   2955
      Tag             =   "1"
      Top             =   6870
      Width           =   195
   End
   Begin VB.Image cmdMasMenos 
      Height          =   420
      Index           =   1
      Left            =   3855
      Tag             =   "1"
      Top             =   6870
      Width           =   195
   End
   Begin VB.Image Image2 
      Height          =   345
      Left            =   6480
      Top             =   180
      Width           =   345
   End
   Begin VB.Image Image1 
      Height          =   450
      Index           =   1
      Left            =   4230
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   6855
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   450
      Index           =   0
      Left            =   585
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   6855
      Width           =   2175
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   3
      Left            =   1560
      TabIndex        =   5
      Top             =   1800
      Visible         =   0   'False
      Width           =   4605
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   1560
      TabIndex        =   4
      Top             =   1995
      Visible         =   0   'False
      Width           =   4020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   5520
      TabIndex        =   3
      Top             =   1560
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   1560
      TabIndex        =   2
      Top             =   1560
      Width           =   2925
   End
End
Attribute VB_Name = "frmBancoObj"
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

Public LasActionBuy As Boolean
Public LastIndex1 As Integer
Public LastIndex2 As Integer

Private m_Number As Integer
Private m_Increment As Integer
Private m_Interval As Integer


Private Sub cantidad_Change()

If Val(cantidad.Text) < 1 Then
    cantidad.Text = 1
End If
If Val(cantidad.Text) > MAX_INVENTORY_OBJS Then
    cantidad.Text = 1
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
        m_Increment = -1
    Case 1
        cmdMasMenos(Index).Picture = LoadInterface("mas-down")
        cmdMasMenos(Index).Tag = "1"
        m_Increment = 1
End Select
tmrNumber.Interval = 30
tmrNumber.Enabled = True
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
        End If
End Select

End Sub

Private Sub cmdMasMenos_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Form_MouseMove(Button, Shift, X, Y)
tmrNumber.Enabled = False
End Sub

Private Sub Form_Deactivate()
'Me.SetFocus
End Sub


Private Sub Form_load()
Me.Picture = LoadInterface("boveda")
End Sub

Private Sub image1_Click(Index As Integer)

Call Audio.PlayWave(SND_CLICK)

If List1(Index).List(List1(Index).ListIndex) = "" Or _
   List1(Index).ListIndex < 0 Then Exit Sub

If Not IsNumeric(cantidad.Text) Then Exit Sub

Select Case Index
    Case 0
        frmBancoObj.List1(0).SetFocus
        LastIndex1 = List1(0).ListIndex
        LasActionBuy = True
        Call WriteBankExtractItem(List1(0).ListIndex + 1, cantidad.Text)
        
   Case 1
        LastIndex2 = List1(1).ListIndex
        LasActionBuy = False
        Call WriteBankDeposit(List1(1).ListIndex + 1, cantidad.Text)
End Select
End Sub

Private Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 0 Then
    Image1(Index).Picture = LoadInterface("retirar-down")
    Image1(Index).Tag = "1"
ElseIf Index = 1 Then
    Image1(Index).Picture = LoadInterface("depositar-down")
    Image1(Index).Tag = "1"
End If
End Sub

Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 0 Then
    If Image1(Index).Tag = "0" Then
        Image1(Index).Picture = LoadInterface("retirar-over")
        Image1(Index).Tag = "1"
    End If
ElseIf Index = 1 Then
    If Image1(Index).Tag = "0" Then
        Image1(Index).Picture = LoadInterface("depositar-over")
        Image1(Index).Tag = "1"
    End If
End If
End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call WriteBankEnd
End Sub

Private Sub Command2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Picture = LoadInterface("salir-down")
Image2.Tag = "1"
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Image2.Tag = "0" Then
    Image2.Picture = LoadInterface("salir-over")
    Image2.Tag = "1"
End If

End Sub

Private Sub list1_Click(Index As Integer)

Select Case Index
    Case 0
        Label1(0).Caption = UserBancoInventory(List1(0).ListIndex + 1).name
        Label1(2).Caption = UserBancoInventory(List1(0).ListIndex + 1).Amount
        Select Case UserBancoInventory(List1(0).ListIndex + 1).OBJType
            Case 2
                Label1(3).Caption = "Max Golpe:" & UserBancoInventory(List1(0).ListIndex + 1).MaxHit
                Label1(4).Caption = "Min Golpe:" & UserBancoInventory(List1(0).ListIndex + 1).MinHit
                Label1(3).Visible = True
                Label1(4).Visible = True
            Case 3, 17
                Label1(3).Visible = False
                Label1(4).Caption = "Defensa:" & UserBancoInventory(List1(0).ListIndex + 1).Def
                Label1(4).Visible = True
            Case Else
                Label1(3).Visible = False
                Label1(4).Visible = False
        End Select
        
        Picture1.Cls
        If UserBancoInventory(List1(0).ListIndex + 1).Amount <> 0 Then _
            Call Engine.DrawGrhToHdc(Picture1.hdc, NPCInventory(List1(0).ListIndex + 1).grhindex, 3, 5)
    Case 1
        Label1(0).Caption = Inventario.ItemName(List1(1).ListIndex + 1)
        Label1(2).Caption = Inventario.Amount(List1(1).ListIndex + 1)
        Select Case Inventario.OBJType(List1(1).ListIndex + 1)
            Case 2
                Label1(3).Caption = "Max Golpe:" & Inventario.MaxHit(List1(1).ListIndex + 1)
                Label1(4).Caption = "Min Golpe:" & Inventario.MinHit(List1(1).ListIndex + 1)
                Label1(3).Visible = True
                Label1(4).Visible = True
            Case 3, 17
                Label1(3).Visible = False
                Label1(4).Caption = "Defensa:" & Inventario.Def(List1(1).ListIndex + 1)
                Label1(4).Visible = True
            Case Else
                Label1(3).Visible = False
                Label1(4).Visible = False
        End Select
        
        Picture1.Cls
        If Inventario.Amount(List1(1).ListIndex + 1) <> 0 Then _
            Call Engine.DrawGrhToHdc(Picture1.hdc, Inventario.grhindex(List1(1).ListIndex + 1), 3, 5)
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

Private Sub tmrNumber_Timer()
Const MIN_NUMBER = 1
Const MAX_NUMBER = 10000

    m_Number = m_Number + m_Increment
    If m_Number < MIN_NUMBER Then
        m_Number = MIN_NUMBER
    ElseIf m_Number > MAX_NUMBER Then
        m_Number = MAX_NUMBER
    End If
    cantidad.Text = Format$(m_Number)
    If m_Interval > 1 Then
        m_Interval = m_Interval - 1
        tmrNumber.Interval = m_Interval
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Image1(0).Tag = "1" Then
    Image1(0).Picture = Nothing
    Image1(0).Tag = "0"
End If

If Image1(1).Tag = "1" Then
    Image1(1).Picture = Nothing
    Image1(1).Tag = "0"
End If

If cmdMasMenos(0).Tag = "1" Then
    cmdMasMenos(0).Picture = Nothing
    cmdMasMenos(0).Tag = "0"
End If

If cmdMasMenos(1).Tag = "1" Then
    cmdMasMenos(1).Picture = Nothing
    cmdMasMenos(1).Tag = "0"
End If

If Image2.Tag = "1" Then
    Image2.Picture = Nothing
    Image2.Tag = "0"
End If

End Sub
