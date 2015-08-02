VERSION 5.00
Begin VB.Form frmCantidad 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   1335
   ClientLeft      =   1635
   ClientTop       =   4410
   ClientWidth     =   2220
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H80000014&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCantidad.frx":0000
   ScaleHeight     =   89
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   148
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
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
      ForeColor       =   &H80000005&
      Height          =   195
      Left            =   300
      MaxLength       =   5
      TabIndex        =   0
      Top             =   555
      Width           =   1485
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   135
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.Image Command1 
      Height          =   255
      Left            =   240
      Top             =   960
      Width           =   855
   End
   Begin VB.Image Command2 
      Height          =   255
      Left            =   1200
      Top             =   960
      Width           =   855
   End
End
Attribute VB_Name = "frmCantidad"
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

Private Sub Command1_Click()
    If LenB(frmCantidad.Text1.Text) > 0 Then
        If Not IsNumeric(frmCantidad.Text1.Text) Then Exit Sub  'Should never happen
        Call WriteDrop(Inventario.SelectedItem, frmCantidad.Text1.Text)
        frmCantidad.Text1.Text = ""
    End If
    
    Unload Me
End Sub


  Private Sub Command2_Click()
      If Inventario.SelectedItem = 0 Then Exit Sub
     
        If Inventario.SelectedItem <> FLAGORO Then
           Call WriteDrop(Inventario.SelectedItem, Inventario.Amount(Inventario.SelectedItem))
           Unload Me
       Else
           If UserGLD > 100000 Then
               Call WriteDrop(Inventario.SelectedItem, 100000)
              Unload Me
          Else
               Call WriteDrop(Inventario.SelectedItem, UserGLD)
               Unload Me
           End If
       End If
   
       frmCantidad.Text1.Text = ""
   End Sub

Private Sub Form_Load()
    Me.Picture = LoadInterface("Cantidad")
End Sub

Private Sub Label1_Click()
Unload Me
End Sub

Private Sub Text1_Change()
    On Error GoTo ErrHandler
     If Val(Text1.Text) < 0 Then
           Text1.Text = "1"
       End If
     
      If Val(Text1.Text) > UserGLD Then
            Text1.Text = UserGLD
        End If
     
      Exit Sub
     
ErrHandler:
       'If we got here the user may have pasted (Shift + Insert) a REALLY large number, causing an overflow, so we set amount back to 1
       Text1.Text = "1"
   End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If (KeyAscii <> 8) Then
        If (KeyAscii < 48 Or KeyAscii > 57) Then
            KeyAscii = 0
        End If
    End If
End Sub
