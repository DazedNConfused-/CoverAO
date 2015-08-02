Attribute VB_Name = "modFunction"
Option Explicit
Public CurServerIp As String
Public CurServerPort As Integer

Function AttactMsg(ByVal chat As String, ByVal color As Long) As Boolean
    If Left(chat, 1) = "¡" And Right$(chat, 1) = "!" Then
        AttactMsg = False
        Exit Function
    End If
    
    If Len(chat) = 2 Then
        If IsNumeric(chat) Then
            If color = vbRed Then
                AttactMsg = False
                Exit Function
            End If
        End If
    End If
    
        
End Function
Public Function RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
    'Initialize randomizer
    Randomize timer
    
    'Generate random number
    RandomNumber = (UpperBound - LowerBound) * Rnd + LowerBound
End Function

Public Function CheckMailString(ByVal sString As String) As Boolean
On Error GoTo errHnd
    Dim lPos  As Long
    Dim lX    As Long
    Dim iAsc  As Integer
    
    '1er test: Busca un simbolo @
    lPos = InStr(sString, "@")
    If (lPos <> 0) Then
        '2do test: Busca un simbolo . después de @ + 1
        If Not (InStr(lPos, sString, ".", vbBinaryCompare) > lPos + 1) Then _
            Exit Function
        
        '3er test: Recorre todos los caracteres y los valída
        For lX = 0 To Len(sString) - 1
            If Not (lX = (lPos - 1)) Then   'No chequeamos la '@'
                iAsc = Asc(mid$(sString, (lX + 1), 1))
                If Not CMSValidateChar_(iAsc) Then _
                    Exit Function
            End If
        Next lX
        
        'Finale
        CheckMailString = True
    End If
errHnd:
End Function
Private Function CMSValidateChar_(ByVal iAsc As Integer) As Boolean
    CMSValidateChar_ = (iAsc >= 48 And iAsc <= 57) Or _
                        (iAsc >= 65 And iAsc <= 90) Or _
                        (iAsc >= 97 And iAsc <= 122) Or _
                        (iAsc = 95) Or (iAsc = 45) Or (iAsc = 46)
End Function

Public Function Field_Read(ByVal field_pos As Long, ByVal Text As String, ByVal delimiter As String) As String
    Dim i As Long
    Dim LastPos As Long
    Dim CurrentPos As Long
    
    LastPos = 0
    CurrentPos = 0
    
    For i = 1 To field_pos
        LastPos = CurrentPos
        CurrentPos = InStr(LastPos + 1, Text, delimiter, vbBinaryCompare)
    Next i
    
    If CurrentPos = 0 Then
        Field_Read = mid$(Text, LastPos + 1, Len(Text) - LastPos)
    Else
        Field_Read = mid$(Text, LastPos + 1, CurrentPos - LastPos - 1)
    End If
End Function
Public Sub ShowSendTxt()
    If Not frmCantidad.Visible Then
        frmMain.SendTxt.Visible = True
        frmMain.SendTxt.SetFocus
    End If
End Sub

Function AsciiValidos(ByVal cad As String) As Boolean
    Dim car As Byte
    Dim i As Long
    
    cad = LCase$(cad)
    
    For i = 1 To Len(cad)
        car = Asc(mid$(cad, i, 1))
        
        If ((car < 97 Or car > 122) Or car = Asc("º")) And (car <> 255) And (car <> 32) Then
            Exit Function
        End If
    Next i
    
    AsciiValidos = True
End Function
Public Function General_Random_Number(ByVal LowerBound As Long, ByVal UpperBound As Long) As Single
    Randomize timer
    General_Random_Number = (UpperBound - LowerBound) * Rnd + LowerBound
End Function

Public Function LoadInterface(ByVal filename As String) As IPictureDisp
    Extract_File Interface, App.Path & "\Interfaces\", LTrim(filename) & ".jpg", App.Path & "\Interfaces\"
       Set LoadInterface = LoadPicture(App.Path & "\Interfaces\" & filename & ".jpg")
End Function


Public Sub DibujarMiniMapPos()

    frmMain.UserP.Left = UserPos.x
    frmMain.UserP.Top = UserPos.y
    frmMain.Minimap.Refresh
    
    frmMain.Label2.Caption = "Posición: " & UserMap & ", " & UserPos.x & ", " & UserPos.y
    
End Sub

Public Function IsAppActive() As Boolean
    IsAppActive = GetActiveWindow
End Function
Public Function LogError(Desc As String)
On Error Resume Next
    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    
    Open App.Path & "\errores.log" For Append As #nfile
        Print #nfile, Desc
    Close #nfile
End Function

Public Function LogCustom(Desc As String)
On Error Resume Next
    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    
    Open App.Path & "\custom.log" For Append As #nfile
        Print #nfile, Now & " " & Desc
    Close #nfile
End Function

Public Function Sounder() As Boolean
    If frmMain.Visible = True Or frmConnect.Visible = True Then
        Sounder = True
        Exit Function
    End If
    Sounder = False
End Function

'Public Function Get_Extract(ByVal file_type As resource_file_type, ByVal file_name As String) As String
'    Extract_File file_type, App.Path & "\Recursos", LCase$(file_name), Windows_Temp_Dir, False
'    Get_Extract = Windows_Temp_Dir & file_name
'End Function
''
'Public Function LoadInterface(ByVal picture_file_name As String) As IPicture
'        Set LoadInterface = LoadPicture(Get_Extract(Interface, LCase$(picture_file_name)))
'        Call Delete_File(Windows_Temp_Dir & LCase$(picture_file_name))
'End Function
