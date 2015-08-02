VERSION 5.00
Begin VB.Form frmCorreo 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Correo"
   ClientHeight    =   6900
   ClientLeft      =   120
   ClientTop       =   390
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Enviar Mensaje"
      Height          =   3615
      Left            =   120
      TabIndex        =   5
      Top             =   3240
      Width           =   6975
      Begin VB.CommandButton Command4 
         Caption         =   "Enviar"
         Height          =   495
         Left            =   4920
         TabIndex        =   16
         Top             =   2880
         Width           =   1575
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Limpiar"
         Height          =   495
         Left            =   4920
         TabIndex        =   15
         Top             =   2280
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   4800
         TabIndex        =   14
         Top             =   1800
         Width           =   1935
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00000000&
         Height          =   615
         Left            =   4800
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   12
         Top             =   600
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Adjuntar Item"
         Height          =   255
         Left            =   2400
         TabIndex        =   11
         Top             =   240
         Width           =   1815
      End
      Begin VB.ListBox List3 
         Enabled         =   0   'False
         Height          =   2790
         Left            =   2160
         TabIndex        =   10
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Height          =   2175
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Cantidad:"
         Height          =   255
         Left            =   4800
         TabIndex        =   13
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Mensaje:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Para:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Mensaje"
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6975
      Begin VB.CommandButton Command2 
         Caption         =   "Guardar Item"
         Height          =   495
         Left            =   5040
         TabIndex        =   4
         Top             =   2520
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Borrar"
         Height          =   495
         Left            =   5040
         TabIndex        =   3
         Top             =   1920
         Width           =   1695
      End
      Begin VB.ListBox List2 
         Height          =   1425
         Left            =   2520
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   4215
      End
      Begin VB.ListBox List1 
         Height          =   2790
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmCorreo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
If List3.Enabled = False Then
Check1.value = 1
List3.Enabled = True
Else
List3.Enabled = False
End If
End Sub

Private Sub Form_Load()
'Picture1.Initialize D3DX, Picture1, MAX_INVENTORY_SLOTS
Dim loopX As Long
For loopX = 1 To MAX_INVENTORY_SLOTS
     With Inventario
            If .OBJIndex(loopX) <> 0 Then
               'Picture1.loopX , .OBJIndex(loopX), .Amount(loopX), .grhindex(loopX), .OBJType(loopX)
            End If
     End With
Next loopX
End Sub

Public Sub DibujaGrh(Grh As Integer)
    Call Engine.DrawGrhToHdc(Picture1.hdc, Grh, 0, 0)
End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub
