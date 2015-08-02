VERSION 5.00
Begin VB.Form frmLauncher 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Launcher"
   ClientHeight    =   6900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   Picture         =   "frmLauncher.frx":0000
   ScaleHeight     =   6900
   ScaleWidth      =   7710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      CausesValidation=   0   'False
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmLauncher.frx":39BC0
      Left            =   480
      List            =   "frmLauncher.frx":39BC2
      TabIndex        =   3
      Top             =   2880
      Width           =   4510
   End
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   120
   End
   Begin VB.ComboBox Combo2 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmLauncher.frx":39BC4
      Left            =   480
      List            =   "frmLauncher.frx":39BCB
      TabIndex        =   1
      Text            =   "800x600 @ 16"
      Top             =   3600
      Width           =   2655
   End
   Begin VB.Image Image3 
      Height          =   9585
      Left            =   -720
      Picture         =   "frmLauncher.frx":39BDD
      Top             =   1320
      Visible         =   0   'False
      Width           =   14550
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   2
      Top             =   3660
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   255
      Index           =   2
      Left            =   720
      Picture         =   "frmLauncher.frx":51A74
      Stretch         =   -1  'True
      Top             =   4005
      Width           =   300
   End
   Begin VB.Image Image2 
      Height          =   240
      Index           =   1
      Left            =   5295
      Picture         =   "frmLauncher.frx":51F2E
      Top             =   2595
      Width           =   300
   End
   Begin VB.Image Image2 
      Height          =   255
      Index           =   0
      Left            =   5290
      Picture         =   "frmLauncher.frx":54DAD
      Top             =   2895
      Width           =   300
   End
   Begin VB.Image image1 
      Height          =   615
      Index           =   1
      Left            =   5640
      MouseIcon       =   "frmLauncher.frx":57D8F
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   6120
      Width           =   1755
   End
   Begin VB.Image image1 
      Height          =   615
      Index           =   0
      Left            =   290
      MouseIcon       =   "frmLauncher.frx":57EE1
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   6120
      Width           =   1755
   End
   Begin VB.Image image1 
      Height          =   600
      Index           =   6
      Left            =   6075
      MouseIcon       =   "frmLauncher.frx":58033
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   5190
      Width           =   1350
   End
   Begin VB.Image image1 
      Height          =   600
      Index           =   5
      Left            =   4640
      MouseIcon       =   "frmLauncher.frx":58185
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   5190
      Width           =   1350
   End
   Begin VB.Image image1 
      Height          =   600
      Index           =   4
      Left            =   3180
      MouseIcon       =   "frmLauncher.frx":582D7
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   5190
      Width           =   1350
   End
   Begin VB.Image image1 
      Height          =   600
      Index           =   3
      Left            =   1740
      MouseIcon       =   "frmLauncher.frx":58429
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   5190
      Width           =   1350
   End
   Begin VB.Image image1 
      Height          =   600
      Index           =   2
      Left            =   285
      MouseIcon       =   "frmLauncher.frx":5857B
      MousePointer    =   99  'Custom
      Tag             =   "1"
      Top             =   5190
      Width           =   1350
   End
   Begin VB.Image Mas 
      Height          =   135
      Left            =   3790
      Top             =   3630
      Width           =   195
   End
   Begin VB.Image Menos 
      Height          =   135
      Left            =   3790
      Top             =   3760
      Width           =   195
   End
   Begin VB.Label INFO 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Bienvenidos de RunekAO, has click en inicias el juego para jugar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   615
      Left            =   2160
      TabIndex        =   0
      Top             =   6120
      Width           =   3375
   End
End
Attribute VB_Name = "frmLauncher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************************************************
'Autor: Juan Beccaceci(Wildem
'Fecha: 08/03/12
'Launcher para IAO CLON :)
'**************************************************************************************************************************
Option Explicit
Public Resul As Boolean

Private Sub Form_load()
Set m_objWMINameSpace = GetObject("winmgmts:")
Set m_objCPUSet = m_objWMINameSpace.InstancesOf("Win32_Processor")
For Each oCpu In m_objCPUSet
     With oCpu
     Combo1.Text = .name
     End With
Next
Exit Sub
End Sub

Private Sub Image2_Click(Index As Integer)
Dim CambiarResolucion As Boolean

 Select Case Index
 Case 0
If Audio.SoundActivated = True Then
Music = 0
Image2(0).Picture = LoadInterface("xd")
Audio.SoundActivated = False
Else
Music = 1
Image2(0).Picture = LoadInterface("TikeSILaun")
Audio.SoundActivated = True
End If

 Case 1
If Audio.SoundActivated = True Then
Sound = 0
Image2(1).Picture = LoadInterface("efectosoff")
Audio.SoundActivated = False
Else
Image2(1).Picture = LoadInterface("efectoon")
Sound = 1
Audio.SoundActivated = True
End If

 Case 2
If Resul = True Then
Resul = False
Image2(2).Picture = LoadInterface("correroff")
Else
Resul = True
Image2(2).Picture = LoadInterface("correron")
End If
End Select

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Image1(0).Tag = "1" Then
        Image1(0).Picture = Nothing
        Image1(0).Tag = "0"
    End If
    
    If Image1(1).Tag = "1" Then
        Image1(1).Picture = Nothing
        Image1(1).Tag = "0"
    End If
    
    If Image1(2).Tag = "1" Then
        Image1(2).Picture = Nothing
        Image1(2).Tag = "0"
    End If
    
    If Image1(3).Tag = "1" Then
        Image1(3).Picture = Nothing
        Image1(3).Tag = "0"
    End If
    
    If Image1(4).Tag = "1" Then
        Image1(4).Picture = Nothing
        Image1(4).Tag = "0"
    End If
    
    If Image1(5).Tag = "1" Then
        Image1(5).Picture = Nothing
        Image1(5).Tag = "0"
    End If
    
    If Image1(6).Tag = "1" Then
        Image1(6).Picture = Nothing
        Image1(6).Tag = "0"
    End If
End Sub


Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

If Index = 0 Then
    If Image1(Index).Tag = "0" Then
            Image1(0).Picture = LoadInterface("salirover")
            Image1(0).Tag = "1"
            INFO.Caption = "Volver al escritorio de Windows"
        End If
     ElseIf Index = 1 Then
    If Image1(Index).Tag = "0" Then
            Image1(1).Picture = LoadInterface("iniciarover")
            Image1(1).Tag = "1"
            INFO.Caption = "¡Inicia RunekAO¡"
        End If
    ElseIf Index = 2 Then
    If Image1(Index).Tag = "0" Then
            Image1(2).Picture = LoadInterface("sitioover")
            Image1(2).Tag = "1"
            INFO.Caption = "Visità www.Runek-AO.com.ar"
        End If
    ElseIf Index = 3 Then
    If Image1(Index).Tag = "0" Then
            Image1(3).Picture = LoadInterface("foroover")
            Image1(3).Tag = "1"
            INFO.Caption = "Visita los foros de discucion donde prodràs, opinal, pedir ayuda o simplemente relajarte"
        End If
   ElseIf Index = 4 Then
    If Image1(Index).Tag = "0" Then
            Image1(4).Picture = LoadInterface("manualover")
            Image1(4).Tag = "1"
            INFO.Caption = "Manual del Juego: Seguramente la mayorìa de tus dudas pueden ser respondidas aquì"
            INFO.Tag = "1"
        End If
    ElseIf Index = 5 Then
    If Image1(Index).Tag = "0" Then
            Image1(5).Picture = LoadInterface("faqover")
            Image1(5).Tag = "1"
            INFO.Caption = "Accede a la seccion preguntas frecuentes del sitio"
            INFO.Tag = "1"
        End If
    ElseIf Index = 6 Then
    If Image1(Index).Tag = "0" Then
            Image1(6).Picture = LoadInterface("notasover")
            Image1(6).Tag = "1"
            INFO.Caption = "Informaciòn de Desarrollo"
            INFO.Tag = "1"
        End If
        End If
End Sub

Private Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

If Index = 0 Then
        Image1(0).Picture = LoadInterface("salirdown")
        Image1(0).Tag = "1"
    ElseIf Index = 1 Then
        Image1(1).Picture = LoadInterface("iniciardown")
        Image1(1).Tag = "1"
    ElseIf Index = 2 Then
        Image1(2).Picture = LoadInterface("sitiodown")
        Image1(2).Tag = "1"
    ElseIf Index = 3 Then
        Image1(3).Picture = LoadInterface("forodown")
        Image1(3).Tag = "1"
    ElseIf Index = 4 Then
        Image1(4).Picture = LoadInterface("manualdown")
        Image1(4).Tag = "1"
    ElseIf Index = 5 Then
        Image1(5).Picture = LoadInterface("faqdown")
        Image1(5).Tag = "1"
    ElseIf Index = 6 Then
        Image1(6).Picture = LoadInterface("notasdown")
        Image1(6).Tag = "1"
        End If
End Sub

Private Sub Image1_Click(Index As Integer)

Select Case Index
  Case 0
If MsgBox("Los datos del Launcher de han modificado, ¿Desea guardar los datos?", vbYesNo) = vbYes Then
frmLauncher.SetFocus
Unload Me
End If

  Case 1
  
Combo1.Visible = False
Combo2.Visible = False
Image3.Picture = LoadInterface("iniciando")
Image3.Visible = True

'Detectar si hay programas habiertos que perjudican el sv
If Detected("sXe Injected.exe") Then
MsgBox "ATENCIÒN: Se detecto el programa: sXe Injected, puede causar el mal funcionamiento del juego, cierrelo"
End
End If
'Se termino :)

If Image Then
Image2(2).Picture = LoadInterface("correron")
Combo2.Text = "800x600 @ 32"
Else
Image2(2).Picture = LoadInterface("correroff")
Combo2.Text = "800x600 @ 16"
End If
Call Main
frmLauncher.Hide

  Case 2
Dim x
x = ShellExecute(Me.hwnd, "Open", "http://www.imperiumao.com.ar", &O0, &O0, SW_Normal)
Case 3
x = ShellExecute(Me.hwnd, "Open", "http://www.imperiumgames.com.ar/foro/f8/", &O0, &O0, SW_Normal)
Case 4
x = ShellExecute(Me.hwnd, "Open", "http://wiki.imperiumao.com.ar/index.php?title=Portada", &O0, &O0, SW_Normal)
Case 5
x = ShellExecute(Me.hwnd, "Open", "http://wiki.imperiumao.com.ar/index.php?title=Portada", &O0, &O0, SW_Normal)
Case 6
Shell ("C:\Users\Carlos\Desktop\Launcher\Cherlong.txt")
End Select
End Sub

Private Sub Timer1_Timer()
Static inter As Long
inter = inter + 1

If inter = 1 Then
    Me.Picture = LoadInterface("presentacion")
End If
End Sub
