Attribute VB_Name = "Anti_Cheat_Citrox"
Public Sub BuscarEngine()
On Error Resume Next
Dim MiObjeto As Object
Set MiObjeto = CreateObject("Wscript.Shell")
Dim X As String
X = "1"
X = MiObjeto.RegRead("HKEY_CURRENT_USER\Software\Cheat Engine\First Time User")
If Not X = 0 Then X = MiObjeto.RegRead("HKEY_USERS\S-1-5-21-343818398-484763869-854245398-500\Software\Cheat Engine\First Time User")
If X = "0" Then
MsgBox "No seas cheater y desistala el cheat engine Chitero xD."
End
End If
Set MiObjeto = Nothing
End Sub
