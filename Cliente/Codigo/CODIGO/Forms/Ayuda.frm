VERSION 5.00
Begin VB.Form Ayuda 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Kega-AO"
   ClientHeight    =   6045
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   ScaleHeight     =   6045
   ScaleWidth      =   8160
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton frmayuda 
      Caption         =   "¿Dondé Entreno?"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Atencion al Cliente de Kega-AO"
      BeginProperty Font 
         Name            =   "Morpheus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "Ayuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
frmdondeleveo.Show
End Sub

Private Sub frmayuda_Click()
frmdondeleveo.Show
End Sub

