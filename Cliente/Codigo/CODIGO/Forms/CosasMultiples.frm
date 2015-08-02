VERSION 5.00
Begin VB.Form CosasMultiples 
   Caption         =   "Form1"
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   13350
   LinkTopic       =   "Form1"
   Picture         =   "CosasMultiples.frx":0000
   ScaleHeight     =   6780
   ScaleWidth      =   13350
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command12 
      Caption         =   "Gm de Consultas"
      Height          =   1095
      Left            =   8160
      TabIndex        =   11
      Top             =   5280
      Width           =   3135
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Equipos "
      Height          =   1095
      Left            =   8160
      TabIndex        =   10
      Top             =   3720
      Width           =   3135
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Mapa de Dungeons"
      Height          =   1095
      Left            =   8160
      TabIndex        =   9
      Top             =   2040
      Width           =   3135
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Como entrenar rapido"
      Height          =   1095
      Left            =   8280
      TabIndex        =   8
      Top             =   360
      Width           =   3015
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Clan Caos"
      Height          =   1095
      Left            =   4320
      TabIndex        =   7
      Top             =   5280
      Width           =   3015
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Clan Sagrada Orden"
      Height          =   1095
      Left            =   4440
      TabIndex        =   6
      Top             =   3720
      Width           =   3015
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Clan Milicia"
      Height          =   1095
      Left            =   4440
      TabIndex        =   5
      Top             =   2040
      Width           =   3015
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Como Fundar clan"
      Height          =   1095
      Left            =   4440
      TabIndex        =   4
      Top             =   360
      Width           =   3015
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Como veo cuantos hay en la Party?"
      Height          =   1095
      Left            =   360
      TabIndex        =   3
      Top             =   5280
      Width           =   3015
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Como elimino a alguien de la Party?"
      Height          =   1095
      Left            =   360
      TabIndex        =   2
      Top             =   3720
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Como aceptar a alguien en la Party?"
      Height          =   1095
      Left            =   360
      TabIndex        =   1
      Top             =   2040
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Como crear una Party?"
      Height          =   1095
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   3015
   End
End
Attribute VB_Name = "CosasMultiples"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
frmOpciones.Show
End Sub
