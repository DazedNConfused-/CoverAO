Attribute VB_Name = "Module1"
Option Explicit
 
'Función Api para aplicar la transparencia a la ventana
Private Declare Function SetLayeredWindowAttributes Lib "user32" _
    (ByVal hwnd As Long, _
     ByVal crKey As Long, _
     ByVal bAlpha As Byte, _
     ByVal dwFlags As Long) As Long
 
'Funciones api para los estilos de la ventana
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
    (ByVal hwnd As Long, _
     ByVal nIndex As Long) As Long
 
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hwnd As Long, _
     ByVal nIndex As Long, _
     ByVal dwNewLong As Long) As Long
 
'Constantes
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const WS_EX_LAYERED = &H80000
'Función que recibe el handle de la ventana y el valor para aplciar la _
 transparencia
Public Function Transparencia(ByVal hwnd As Long, Valor As Integer) As Long
 
On Local Error GoTo ErrSub
 
Dim Estilo As Long
 
If Valor < 0 Or Valor > 255 Then
    Transparencia = 1
Else
 
    Estilo = GetWindowLong(hwnd, GWL_EXSTYLE)
    Estilo = Estilo Or WS_EX_LAYERED
   
    SetWindowLong hwnd, GWL_EXSTYLE, Estilo
   
    'Aplica el nuevo estilo con la transparencia
    SetLayeredWindowAttributes hwnd, 0, Valor, LWA_ALPHA
   
    Transparencia = 0
End If
 
If Err Then
    Transparencia = 2
End If
   
Exit Function
 
'Error
ErrSub:
   
   MsgBox Err.Description, vbCritical, "Error"
 
End Function

