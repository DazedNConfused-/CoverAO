VERSION 5.00
Begin VB.Form frmInterface 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "CoverAO-By Eter"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmInterface.frx":0000
   ScaleHeight     =   2610
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   855
      Left            =   840
      TabIndex        =   2
      Top             =   1440
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   735
      Left            =   2520
      TabIndex        =   1
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton skin1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "frmInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
frmMain.skin1 = 1
frmMain.skin2 = 0
frmMain.skin3 = 0
     frmMain.Picture = LoadPicture(App.Path & _
    "\Interfaces\Imagen main 1")
       
    frmMain.InvEqu.Picture = LoadPicture(App.Path & _
    "\Interfaces\imagen inventario 1")
   
    frmMain.InvEqu.Picture = LoadPicture(App.Path & _
    "\Interfaces\imagen inventario completo 1 ")
 
End Sub
 
Private Sub Command2_Click()
frmMain.skin1 = 0
frmMain.skin2 = 1
frmMain.skin3 = 0
 
     frmMain.Picture = LoadPicture(App.Path & _
    "\Interfaces\imagen main 2")
       
   
    frmMain.InvEqu.Picture = LoadPicture(App.Path & _
    "\Interfaces\imagen inventario 2")
   
    frmMain.InvEqu.Picture = LoadPicture(App.Path & _
    "\Interfaces\imagen inventario completo 2")
End Sub
Private Sub Command3_Click()
frmMain.skin1 = 0
frmMain.skin2 = 0
frmMain.skin3 = 1
 
     frmMain.Picture = LoadPicture(App.Path & _
    "\Interfaces\imagen main 3")
       
   
    frmMain.InvEqu.Picture = LoadPicture(App.Path & _
    "\Interfaces\imagen inventario 3")
   
    frmMain.InvEqu.Picture = LoadPicture(App.Path & _
    "\Interfaces\imagen inventario completo 3")
 
End Sub
 
Private Sub Image1_Click()
Unload Me
End Sub

