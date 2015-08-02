VERSION 5.00
Begin VB.Form frmBot 
   Caption         =   "Control Bot"
   ClientHeight    =   1545
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2130
   LinkTopic       =   "Form1"
   ScaleHeight     =   1545
   ScaleWidth      =   2130
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPala 
      Caption         =   "Pala"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   2055
   End
   Begin VB.CommandButton cmdMago 
      Caption         =   "Mago"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   600
      TabIndex        =   0
      Text            =   "1"
      Top             =   90
      Width           =   735
   End
End
Attribute VB_Name = "frmBot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdMago_Click()
Dim I As Byte
For I = 1 To val(Text1.Text)
    Call WriteCreateNPC(581)
Next I
End Sub

Private Sub cmdPala_Click()
Dim I As Byte
For I = 1 To val(Text1.Text)
    Call WriteCreateNPC(563)
Next I
End Sub

