VERSION 5.00
Begin VB.Form Arenas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Arenas"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Salir Arenas"
      Height          =   855
      Left            =   3120
      TabIndex        =   1
      Top             =   2160
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ir Arenas"
      Height          =   855
      Left            =   480
      TabIndex        =   0
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   11520
      Left            =   -5040
      Picture         =   "Arenas.frx":0000
      Top             =   -4800
      Width           =   15360
   End
End
Attribute VB_Name = "Arenas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Call WriteDeath
Unload Me
End Sub

Private Sub Command2_Click()
Call WriteSinDeath
Unload Me
End Sub
