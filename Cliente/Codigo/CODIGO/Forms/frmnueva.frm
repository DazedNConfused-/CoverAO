VERSION 5.00
Begin VB.Form frmnueva 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Formulario de mensaje a adminitradores"
   ClientHeight    =   4155
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   5325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Accede al Control de cuentas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   840
      TabIndex        =   6
      Top             =   3120
      Width           =   3375
   End
   Begin VB.OptionButton Option5 
      Caption         =   "Quiero denunciar a otro personaje"
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   2400
      Width           =   2775
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Nesesito ayuda general sobre el juego o como jugar"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   2040
      Width           =   4095
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Quiero reportar un error en el juego"
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   1560
      Width           =   2775
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Tengo problemas con mi cuenta"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   1200
      Width           =   2655
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Tengo problema con uno de mis personajes"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   840
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Por favor, Seleccione su tipo de consulta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   4215
   End
End
Attribute VB_Name = "frmnueva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Option2_Click()
cuentawe.Show
Unload Me
End Sub

Private Sub Option3_Click()
errores.Show
Unload Me
End Sub

Private Sub Option4_Click()
ayudajuego.Show
Unload Me
End Sub

Private Sub Option5_Click()
denuncias.Show
Unload Me
End Sub
