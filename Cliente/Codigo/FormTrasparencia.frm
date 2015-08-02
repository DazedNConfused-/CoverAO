VERSION 5.00
Begin VB.Form FormTrasparencia 
   Caption         =   "Form1"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picMacro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   480
      Index           =   0
      Left            =   1200
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   960
      Width           =   480
   End
End
Attribute VB_Name = "FormTrasparencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Aca declaramos
Dim i As Integer
 
Private Sub Form_Load()
 
If Not Transparencia(frmOpciones.hwnd, 0) = 0 Then
   
    MsgBox " Esta función Api no es soportada en Versiones" _
           & "anteriores a windows 2000", vbCritical
    Me.Show
Else
 
    ' Gradua la transparencia del formulario hasta hacerla visible _
     es decir desde el valor 0 hasta el 255
   
    'desactiva el Formulario
    Me.Enabled = False
    Me.Show
   
    For i = 0 To 255 Step 2
        ' Va aplicando los distintos valores y grados de transparencia al form
        Call Transparencia(frmOpciones.hwnd, i)
        DoEvents
    Next
   
    'reactiva la ventana
    Me.Enabled = True
     
End If
 
End Sub
 
 
'Al descargar la ventana hace el efecto FadeOut, osea cuando el formulario desaparece
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 
If Not Transparencia(frmOpciones.hwnd, 0) = 0 Then
    Exit Sub
Else
    ' Gradua la transparencia del formulario hasta hacerla invisible y luego se descarga, desde el valor 255 hasta el 0
    For i = 255 To 0 Step -3
        DoEvents
        Call Transparencia(frmOpciones.hwnd, i)
        DoEvents
    Next
   
End If
 
End Sub
