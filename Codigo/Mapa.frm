VERSION 5.00
Begin VB.Form Mapa 
   BorderStyle     =   0  'None
   Caption         =   "Sherasd"
   ClientHeight    =   8970
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12570
   LinkTopic       =   "Form1"
   Picture         =   "Mapa.frx":0000
   ScaleHeight     =   8970
   ScaleWidth      =   12570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   120
      Picture         =   "Mapa.frx":3ED81
      Top             =   3480
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   120
      Picture         =   "Mapa.frx":3F9C3
      Top             =   5760
      Width           =   480
   End
   Begin VB.Image Inframundo 
      Height          =   9000
      Left            =   600
      Picture         =   "Mapa.frx":40605
      Top             =   0
      Width           =   11970
   End
   Begin VB.Image Mundo 
      Height          =   8970
      Left            =   600
      Picture         =   "Mapa.frx":57017
      Top             =   0
      Width           =   11970
   End
End
Attribute VB_Name = "Mapa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Image3_Click()
Inframundo.Visible = True
Mundo.Visible = False
End Sub

Private Sub Image4_Click()
Inframundo.Visible = False
Mundo.Visible = True
End Sub
