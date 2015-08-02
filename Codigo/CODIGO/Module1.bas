Attribute VB_Name = "Module1"
'**************Autor: Nait(Nicolás Pedetti)**************************
Option Explicit
Public Const NumItemsLuz As Byte = 2 'Numeros de items con luz
Public GrhItemLuz(1 To NumItemsLuz) As Integer
 
Function TieneLuz(ByVal X As Byte, ByVal Y As Byte)
'**************Autor: Nait(Nicolás Pedetti)**************************
Dim i As Byte
For i = 1 To NumItemsLuz
   If GrhItemLuz(i) = MapData(X, Y).ObjGrh.grhindex And MapData(X, Y).OBJInfo.TieneLuz = 0 Then
   MapData(X, Y).OBJInfo.TieneLuz = 1
   Light.Create_Light_To_Map X, Y, 3, 255, 255, 255
   End If
Next i
End Function
 
Function DeletLuz(ByVal X As Byte, ByVal Y As Byte)
'**************Autor: Nait(Nicolás Pedetti)**************************
Dim i As Byte
For i = 1 To NumItemsLuz
    If GrhItemLuz(i) = MapData(X, Y).ObjGrh.grhindex And MapData(X, Y).OBJInfo.TieneLuz = 1 Then
        Light.Delete_Light_To_Map X, Y
    End If
Next i
End Function
 
Sub ObjLuz()
'**************Autor: Nait(Nicolás Pedetti)**************************
GrhItemLuz(1) = 1521 'Fogata
GrhItemLuz(2) = 912 'Daga comun
End Sub
