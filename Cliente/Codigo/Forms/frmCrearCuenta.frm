VERSION 5.00
Begin VB.Form frmCrearCuenta 
   BorderStyle     =   0  'None
   Caption         =   "Crear nueva cuenta"
   ClientHeight    =   7650
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCrearCuenta.frx":0000
   ScaleHeight     =   7650
   ScaleWidth      =   8550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   375
      Left            =   7320
      TabIndex        =   6
      Top             =   7080
      Width           =   975
   End
   Begin VB.TextBox mailTxt 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1680
      TabIndex        =   5
      Top             =   4320
      Width           =   3135
   End
   Begin VB.ComboBox CuentQuestions 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmCrearCuenta.frx":271CD
      Left            =   720
      List            =   "frmCrearCuenta.frx":271DD
      TabIndex        =   4
      Top             =   5880
      Width           =   2175
   End
   Begin VB.TextBox answerTxt 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   6240
      Width           =   2175
   End
   Begin VB.TextBox pass1Txt 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3000
      PasswordChar    =   "x"
      TabIndex        =   2
      Top             =   3640
      Width           =   2895
   End
   Begin VB.TextBox passTxt 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "x"
      TabIndex        =   1
      Top             =   3000
      Width           =   2895
   End
   Begin VB.TextBox nameTxt 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   2300
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   3120
      TabIndex        =   7
      Top             =   6480
      Width           =   2415
   End
End
Attribute VB_Name = "frmCrearCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = frmMain.Icon
End Sub

Private Sub Label1_Click()
UserAccount = nameTxt.Text
    UserPassword = passTxt.Text
    UserEmail = mailTxt.Text
    
    If Not UserPassword = pass1Txt.Text Then
        MsgBox "Las contraseñas no coinciden."
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
