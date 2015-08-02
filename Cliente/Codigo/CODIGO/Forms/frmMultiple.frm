VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4890
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   7920
   LinkTopic       =   "Arenas"
   ScaleHeight     =   4890
   ScaleWidth      =   7920
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "Sagrada Orden"
      Height          =   495
      Left            =   4560
      TabIndex        =   5
      Top             =   4200
      Width           =   2295
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Republicano"
      Height          =   495
      Left            =   4560
      TabIndex        =   4
      Top             =   3480
      Width           =   2295
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Caos"
      Height          =   495
      Left            =   4560
      TabIndex        =   3
      Top             =   2640
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Ver mapa de Eventos"
      Height          =   975
      Left            =   240
      TabIndex        =   2
      Top             =   3240
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir de arenas"
      Height          =   975
      Left            =   4560
      TabIndex        =   1
      Top             =   720
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ir Arenas"
      Height          =   975
      Left            =   240
      Picture         =   "Multiple.frx":0000
      TabIndex        =   0
      Top             =   720
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   11520
      Left            =   -720
      Picture         =   "Multiple.frx":3ED81
      Top             =   -3600
      Width           =   15360
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
If MapInfo(UserList(UserIndex).Pos.map).Pk = True Then
               Call WriteConsoleMsg(1, UserIndex, "Debes estar en Zona Segura.", FontTypeNames.FONTTYPE_INFO)
               Exit Sub
               End If
               
                Dim Ciudad As WorldPos
                Dim Destino As String
               
                Select Case (0)
               
                Case 0
                Destino = "Mapa evento"
                Ciudad.map = 751 'Mapa Evento
                Ciudad.X = 24
                Ciudad.Y = 19

                
            End Select
            Call WarpUserChar(UserIndex, Ciudad.map, Ciudad.X, Ciudad.Y, True)
        
        Exit Sub
End Sub

Private Sub Command2_Click()
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
End Sub
