VERSION 5.00
Begin VB.Form EcharParty 
   Caption         =   "Eliminar Party"
   ClientHeight    =   2280
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   ScaleHeight     =   2280
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Caption         =   "Para eliminar a alguien de la party usa el comando: /ECHARPARTY (nombre del Personaje)"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   6615
   End
   Begin VB.Image Image1 
      Height          =   11520
      Left            =   -5160
      Picture         =   "Comoeliminara alguiendelpary.frx":0000
      Top             =   0
      Width           =   15360
   End
End
Attribute VB_Name = "EcharParty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

