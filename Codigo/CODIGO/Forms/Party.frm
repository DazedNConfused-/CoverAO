VERSION 5.00
Begin VB.Form Party 
   Caption         =   "CoverAO - Party"
   ClientHeight    =   5820
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   4740
   LinkTopic       =   "Form1"
   Palette         =   "Party.frx":0000
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5820
   ScaleWidth      =   4740
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Expulsar del Party"
      Height          =   735
      Left            =   1080
      TabIndex        =   5
      Top             =   3120
      Width           =   2535
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Salir  Party"
      Height          =   735
      Left            =   1080
      TabIndex        =   3
      Top             =   4080
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Online Party"
      Height          =   735
      Left            =   1080
      TabIndex        =   2
      Top             =   2160
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Agregar Pj al Party"
      Height          =   735
      Left            =   1080
      TabIndex        =   1
      Top             =   1200
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Crear Party"
      Height          =   735
      Left            =   1080
      Picture         =   "Party.frx":32C6
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "La Experiencia Ganada se entrega al Terminar la Party"
      Height          =   495
      Left            =   1320
      TabIndex        =   4
      Top             =   5160
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   24780
      Left            =   0
      Picture         =   "Party.frx":104A11
      Top             =   0
      Width           =   33045
   End
End
Attribute VB_Name = "Party"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Call WritePartyCreate
Unload Me
End Sub

Private Sub Command2_Click()
PartyAcept.Show
Unload Me
End Sub

Private Sub Command3_Click()
Call WritePartyOnline
Unload Me
End Sub

Private Sub Command4_Click()
Call WritePartyLeave
Unload Me
End Sub

Private Sub Command5_Click()
EcharParty.Show
Unload Me
End Sub
