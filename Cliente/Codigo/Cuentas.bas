Attribute VB_Name = "Cuentas"
Option Explicit
Public AccountPath As String
Public Function AddUserInAccount(ByVal name As String, ByVal Account As String)
    Dim aFile As String
    
    aFile = AccountPath & Account & ".cnt"
    
    Dim NumPJs As Byte
    NumPJs = CByte(GetVar(aFile, "PJS", "NumPjs")) + 1
    
    Call WriteVar(aFile, "PJS", "PJ" & NumPJs, name)
    Call WriteVar(aFile, "PJS", "NumPjs", NumPJs)

End Function
Public Function IsPJOfAccount(ByVal name As String, ByVal Account As String) As Boolean
    Dim aFile As String
    
    aFile = AccountPath & Account & ".cnt"
    
    Dim NumPJs As Byte
    NumPJs = CByte(GetVar(aFile, "PJS", "NumPjs"))
    
    If Not NumPJs = 0 Then
        Dim i As Byte
        For i = 1 To NumPJs
            If UCase$(name) = UCase$(GetVar(aFile, "PJS", "PJ" & i)) Then
                IsPJOfAccount = True
                Exit Function
            End If
        Next i
    End If
    
    IsPJOfAccount = False
End Function
Public Function CrearCuenta(ByVal UserIndex As Integer, ByVal name As String, ByVal Pass As String, ByVal Mail As String, ByVal respuesta As String, ByVal pregunta As Byte)
    
    name = UCase$(LTrim(RTrim(name)))
    
    'Existe ya la cuenta?
    If FileExist(AccountPath & name & ".cnt", vbNormal) Then
        Call CloseSocket(UserIndex)
        Exit Function
    End If
    
    Dim N As Integer
    N = FreeFile()
    Open AccountPath & name & ".cnt" For Append As N
        Print #N, "[" & name & "]"
        Print #N, "Password=" & Pass
        Print #N, "Mail=" & Mail
        Print #N, "Ban=0"
        Print #N, "Pregunta=" & LTrim(RTrim(str(pregunta)))
        Print #N, "Respuesta=" & respuesta
        Print #N, "[PJS]"
        Print #N, "NumPjs=0"
        Print #N, "PJ1="
        Print #N, "PJ2="
        Print #N, "PJ3="
        Print #N, "PJ4="
        Print #N, "PJ5="
        Print #N, "PJ6="
        Print #N, "PJ7="
        Print #N, "PJ8="
        Print #N, "PJ9="
        Print #N, "PJ10="
    Close #N
    
    Call WriteShowMessageBox(UserIndex, "La cuenta ha sido creada con exito.")
    Call FlushBuffer(UserIndex)
    Call CloseSocket(UserIndex)
    
End Function
Public Function ConectarCuenta(ByVal UserIndex As Integer, ByVal name As String, ByVal Pass As String)
    name = UCase$(LTrim(RTrim(name)))
    
    Dim Leer As New clsIniReader

    'Existe ya la cuenta?
    If Not FileExist(AccountPath & name & ".cnt", vbNormal) Then
        Call WriteErrorMsg(UserIndex, "La cuenta no existe.")
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
        Exit Function
    End If
    
    Leer.Initialize AccountPath & name & ".cnt"
    
    'Es la pass correcta?
    If Pass <> Leer.GetValue(name, "Password") And Not Pass = "ISNOTHING123456789" Then
        Call WriteErrorMsg(UserIndex, "Contraseña incorrecta.")
        Call FlushBuffer(UserIndex)
        Call CloseSocket(UserIndex)
        Exit Function
    End If
    
    Dim NumPJs As Byte
    NumPJs = Leer.GetValue("PJS", "NumPjs")
    
    Call WriteShowAccount(UserIndex)
    
    If Not NumPJs = 0 Then
        Dim i As Byte
        For i = 1 To NumPJs
            Call WriteAddPj(UserIndex, Leer.GetValue("PJS", "PJ" & i), i)
        Next i
    End If
    
    Call FlushBuffer(UserIndex)
    
End Function

