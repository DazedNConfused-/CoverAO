VERSION 5.00
Begin VB.Form frmCrearCuenta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Crear nueva cuenta"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   2895
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox mailTxt 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   255
      IMEMode         =   3  'DISABLE
      Left            =   360
      TabIndex        =   10
      Top             =   2160
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   3960
      Width           =   2655
   End
   Begin VB.ComboBox CuentQuestions 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmCrearCuenta.frx":0000
      Left            =   360
      List            =   "frmCrearCuenta.frx":0010
      TabIndex        =   8
      Top             =   2640
      Width           =   2175
   End
   Begin VB.TextBox answerTxt 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   3240
      Width           =   2175
   End
   Begin VB.TextBox pass1Txt 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   255
      IMEMode         =   3  'DISABLE
      Left            =   360
      PasswordChar    =   "x"
      TabIndex        =   4
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox passTxt 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   255
      IMEMode         =   3  'DISABLE
      Left            =   360
      PasswordChar    =   "x"
      TabIndex        =   2
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox nameTxt 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Correo Electronico"
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Respuesta"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Confirmar contraseña"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Contraseña"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Nombre de cuenta:"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmCrearCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    UserAccount = nameTxt.Text
    UserPassword = passTxt.Text
    UserEmail = mailTxt.Text
    
    If Not UserPassword = pass1Txt.Text Then
        MsgBox "Las contraseñas no coinciden."
        Exit Sub
    End If
    
    If Not CheckMailString(UserEmail) Then
        MsgBox "Direccion de mail invalida."
        Exit Sub
    End If
    
    UserAnswer = answerTxt.Text
    UserQuestion = CuentQuestions.ListIndex
    
    If Len(UserAnswer) < 11 Then
        MsgBox "Respuesta muy corta"
        Exit Sub
    End If

    EstadoLogin = CrearNuevaCuenta
    If frmMain.Socket1.Connected Then
        frmMain.Socket1.Disconnect
        frmMain.Socket1.Cleanup
        DoEvents
    End If
    
    frmMain.Socket1.HostName = CurServerIp
    frmMain.Socket1.RemotePort = CurServerPort
    frmMain.Socket1.Connect
    
    Unload Me
End Sub

Private Sub Form_load()
    Me.Icon = frmMain.Icon
End Sub
