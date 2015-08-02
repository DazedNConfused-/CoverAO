VERSION 5.00
Begin VB.Form PartyAcept 
   Caption         =   "Form1"
   ClientHeight    =   1845
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   ScaleHeight     =   1845
   ScaleWidth      =   8325
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label2 
      Caption         =   "Recuerden que la Experiencia y el oro ganan al eliminar la party"
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   960
      Width           =   4575
   End
   Begin VB.Label Label1 
      Caption         =   "Para haceptar a alguien en la party usa el comando: /ACCEPTPARTY (nombre del Personaje)"
      Height          =   255
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Width           =   6735
   End
   Begin VB.Image Pr 
      Height          =   11520
      Left            =   -6000
      Picture         =   "Aceptar a alguien en la party.frx":0000
      Top             =   -720
      Width           =   15360
   End
End
Attribute VB_Name = "PartyAcept"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

